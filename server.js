const fs = require("fs");
const fsp = require("fs/promises");
const os = require("os");
const path = require("path");
const { randomUUID } = require("crypto");
const express = require("express");
const multer = require("multer");
const archiver = require("archiver");
const csvParser = require("csv-parser");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");
const XlsxStreamReader = require("xlsx-stream-reader");

const app = express();
const port = process.env.PORT || 3000;
const uploadDirectory = path.join(os.tmpdir(), "campaign-p360-uploads");
const generatedDownloads = new Map();
const downloadRetentionMs = 30 * 60 * 1000;

fs.mkdirSync(uploadDirectory, { recursive: true });

const upload = multer({
  dest: uploadDirectory,
  limits: {
    fileSize: 1536 * 1024 * 1024,
  },
});

function sanitizeFilePart(value) {
  return String(value || "")
    .trim()
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ");
}

function normalizeHeader(value) {
  return String(value || "").trim().toLowerCase();
}

function extractCampaignId(campaignValue) {
  const value = String(campaignValue || "").trim();

  if (!value) {
    return null;
  }

  const match = value.match(/(\d+)(?!.*\d)/);
  if (match) {
    return match[1];
  }

  return value.replace(/\s+/g, "_").replace(/[^A-Za-z0-9_-]/g, "");
}

function normalizeCellValue(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date) {
    return value;
  }

  if (typeof value !== "object") {
    return value;
  }

  if (Array.isArray(value.richText)) {
    return value.richText.map((part) => part.text || "").join("");
  }

  if ("result" in value && value.result !== undefined && value.result !== null) {
    return value.result;
  }

  if ("text" in value && value.text) {
    return value.text;
  }

  if ("hyperlink" in value && value.hyperlink) {
    return value.hyperlink;
  }

  if ("formula" in value && value.formula) {
    return value.formula;
  }

  if ("error" in value && value.error) {
    return value.error;
  }

  return String(value);
}

function createCampaignWorkbookStore(outputDirectory, appName) {
  const workbooks = new Map();

  function getEntry(campaignId) {
    if (!workbooks.has(campaignId)) {
      const filename = `${appName} - ${campaignId} - p360.xlsx`;
      const fullPath = path.join(outputDirectory, filename);
      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: fullPath,
        useSharedStrings: false,
        useStyles: false,
      });

      workbooks.set(campaignId, {
        filename,
        fullPath,
        workbook,
        installSheet: null,
        inAppSheet: null,
        installCount: 0,
        inAppCount: 0,
      });
    }

    return workbooks.get(campaignId);
  }

  function appendRow(campaignId, sourceType, headers, rowValues) {
    const entry = getEntry(campaignId);
    const isInstall = sourceType === "install";
    const sheetKey = isInstall ? "installSheet" : "inAppSheet";
    const sheetName = isInstall ? "PA - Install" : "PA - InApp";

    if (!entry[sheetKey]) {
      entry[sheetKey] = entry.workbook.addWorksheet(sheetName);
      entry[sheetKey].addRow(headers).commit();
    }

    entry[sheetKey].addRow(rowValues).commit();

    if (isInstall) {
      entry.installCount += 1;
    } else {
      entry.inAppCount += 1;
    }
  }

  async function finalize() {
    const entries = Array.from(workbooks.values()).sort((left, right) =>
      left.filename.localeCompare(right.filename, undefined, { numeric: true }),
    );

    for (const entry of entries) {
      if (entry.installSheet) {
        entry.installSheet.commit();
      }

      if (entry.inAppSheet) {
        entry.inAppSheet.commit();
      }

      await entry.workbook.commit();
    }

    return entries;
  }

  return {
    appendRow,
    finalize,
  };
}

async function processCsvFile(filePath, sourceType, store) {
  return new Promise((resolve, reject) => {
    let headers = null;
    let campaignIndex = -1;
    let isSettled = false;
    let dataRowCount = 0;

    function fail(error) {
      if (!isSettled) {
        isSettled = true;
        reject(error);
      }
    }

    const stream = fs.createReadStream(filePath);
    const parser = csvParser({
      headers: false,
      skipComments: false,
    });

    parser.on("data", (row) => {
      const rawValues = Object.values(row).map((value) =>
        typeof value === "string" ? value.trim() : value ?? "",
      );

      if (!headers) {
        const detectedCampaignIndex = rawValues.findIndex(
          (value) => normalizeHeader(value) === "campaign",
        );

        if (detectedCampaignIndex === -1) {
          return;
        }

        headers = rawValues;
        campaignIndex = detectedCampaignIndex;
        return;
      }

      const rowValues = headers.map((_, index) => rawValues[index] ?? "");
      const campaignValue = rowValues[campaignIndex];
      const campaignId = extractCampaignId(campaignValue);

      if (!campaignId) {
        return;
      }

      store.appendRow(campaignId, sourceType, headers, rowValues);
      dataRowCount += 1;
    });

    parser.on("end", () => {
      if (!headers) {
        fail(
          new Error(
            'CSV file must contain a row with a "Campaign" column header.',
          ),
        );
        return;
      }

      if (!isSettled) {
        isSettled = true;
        resolve({ dataRowCount });
      }
    });

    parser.on("error", fail);
    stream.on("error", fail);
    stream.pipe(parser);
  });
}

async function processXlsxFile(filePath, sourceType, store) {
  return new Promise((resolve, reject) => {
    let isSettled = false;
    let sawWorksheet = false;
    let dataRowCount = 0;

    function fail(error) {
      if (!isSettled) {
        isSettled = true;
        reject(error);
      }
    }

    const workbookReader = new XlsxStreamReader({
      formatting: false,
      saxTrim: false,
      verbose: false,
    });

    workbookReader.on("worksheet", (worksheetReader) => {
      if (sawWorksheet) {
        worksheetReader.skip();
        return;
      }

      sawWorksheet = true;

      let headers = null;
      let campaignIndex = -1;

      worksheetReader.on("row", (row) => {
        if (!headers) {
          const candidateHeaders = Array.from(
            { length: row.values.length - 1 },
            (_, index) =>
              String(normalizeCellValue(row.values[index + 1]) || "").trim(),
          );
          const detectedCampaignIndex = candidateHeaders.findIndex(
            (header) => normalizeHeader(header) === "campaign",
          );

          if (detectedCampaignIndex === -1) {
            return;
          }

          headers = candidateHeaders;
          campaignIndex = detectedCampaignIndex;
          return;
        }

        const rowValues = headers.map((_, index) =>
          normalizeCellValue(row.values[index + 1]),
        );
        const campaignId = extractCampaignId(rowValues[campaignIndex]);

        if (!campaignId) {
          return;
        }

        store.appendRow(campaignId, sourceType, headers, rowValues);
        dataRowCount += 1;
      });

      worksheetReader.on("end", () => {
        if (!headers) {
          fail(
            new Error(
              'XLSX file must contain a row with a "Campaign" column header.',
            ),
          );
        }
      });

      worksheetReader.on("error", fail);
      worksheetReader.process();
    });

    workbookReader.on("error", fail);
    workbookReader.on("end", () => {
      if (!sawWorksheet) {
        fail(new Error("XLSX file does not contain any sheet."));
        return;
      }

      if (!isSettled) {
        isSettled = true;
        resolve({ dataRowCount });
      }
    });

    fs.createReadStream(filePath).on("error", fail).pipe(workbookReader);
  });
}

async function detectFileType(file) {
  const handle = await fsp.open(file.path, "r");

  try {
    const buffer = Buffer.alloc(4);
    await handle.read(buffer, 0, 4, 0);

    if (
      buffer[0] === 0x50 &&
      buffer[1] === 0x4b &&
      buffer[2] === 0x03 &&
      buffer[3] === 0x04
    ) {
      return "xlsx";
    }
  } finally {
    await handle.close();
  }

  const extension = path.extname(file.originalname || "").toLowerCase();

  if (extension === ".xlsx") {
    return "xlsx";
  }

  if (extension === ".csv") {
    return "csv";
  }

  return "csv";
}

async function readMatrixFromFile(file) {
  const fileType = await detectFileType(file);

  if (fileType === "xlsx") {
    const workbook = XLSX.readFile(file.path, { cellDates: true });
    const firstSheetName = workbook.SheetNames[0];

    if (!firstSheetName) {
      throw new Error("File does not contain any sheet.");
    }

    return {
      fileType,
      sheetName: firstSheetName,
      rows: XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
        header: 1,
        defval: "",
        raw: true,
      }),
    };
  }

  const content = await fsp.readFile(file.path, "utf8");
  const workbook = XLSX.read(content, { raw: true, type: "string" });
  const firstSheetName = workbook.SheetNames[0];

  return {
    fileType,
    sheetName: firstSheetName || "Sheet1",
    rows: XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
      header: 1,
      defval: "",
      raw: true,
    }),
  };
}

function findHeaderRowIndex(rows, targetHeader) {
  return rows.findIndex((row) =>
    row.some((cell) => normalizeHeader(cell) === normalizeHeader(targetHeader)),
  );
}

function findBestHeaderRowIndex(rows) {
  let bestIndex = -1;
  let bestScore = 0;

  rows.slice(0, 50).forEach((row, index) => {
    const nonEmptyValues = row
      .map((cell) => String(cell ?? "").trim())
      .filter((value) => value !== "");
    const uniqueValues = new Set(nonEmptyValues.map((value) => value.toLowerCase()));
    const score = nonEmptyValues.length + uniqueValues.size * 0.1;

    if (nonEmptyValues.length >= 2 && score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestIndex;
}

function padRow(row, columnCount) {
  return Array.from({ length: columnCount }, (_, index) => row[index] ?? "");
}

function extractHeaderInfo(rows, preferredHeader) {
  let headerRowIndex = -1;

  if (preferredHeader) {
    headerRowIndex = findHeaderRowIndex(rows, preferredHeader);
  }

  if (headerRowIndex === -1) {
    headerRowIndex = findBestHeaderRowIndex(rows);
  }

  if (headerRowIndex === -1) {
    return null;
  }

  const headerRow = rows[headerRowIndex];
  const columnCount = Math.max(...rows.map((row) => row.length), headerRow.length, 0);
  const paddedHeaderRow = padRow(headerRow, columnCount);
  const columns = paddedHeaderRow
    .map((value) => String(value ?? "").trim())
    .filter((value) => value !== "");

  return {
    headerRowIndex,
    headerRow,
    columnCount,
    paddedHeaderRow,
    columns,
  };
}

async function processInputFile(file, sourceType, store) {
  const fileType = await detectFileType(file);

  if (fileType === "csv") {
    return processCsvFile(file.path, sourceType, store);
  }

  if (fileType === "xlsx") {
    return processXlsxFile(file.path, sourceType, store);
  }

  throw new Error("Only .xlsx and .csv files are supported.");
}

async function createZipFile(zipPath, files) {
  await new Promise((resolve, reject) => {
    const output = fs.createWriteStream(zipPath);
    const archive = archiver("zip", { zlib: { level: 0 } });

    output.on("close", resolve);
    output.on("error", reject);
    archive.on("error", reject);
    archive.pipe(output);

    for (const file of files) {
      archive.file(file.fullPath, { name: file.filename });
    }

    archive.finalize();
  });
}

function registerGeneratedDownload(zipPath, cleanupTargets, filename) {
  const token = randomUUID();
  const timeout = setTimeout(async () => {
    const download = generatedDownloads.get(token);

    if (!download) {
      return;
    }

    generatedDownloads.delete(token);
    await cleanupPaths(download.cleanupTargets);
  }, downloadRetentionMs);

  if (typeof timeout.unref === "function") {
    timeout.unref();
  }

  generatedDownloads.set(token, {
    zipPath,
    cleanupTargets,
    filename,
    timeout,
  });

  return token;
}

async function cleanupGeneratedDownload(token) {
  const download = generatedDownloads.get(token);

  if (!download) {
    return;
  }

  clearTimeout(download.timeout);
  generatedDownloads.delete(token);
  await cleanupPaths(download.cleanupTargets);
}

async function cleanupPaths(paths) {
  await Promise.all(
    paths.map((targetPath) =>
      fsp.rm(targetPath, { recursive: true, force: true }).catch(() => {}),
    ),
  );
}

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.get("/campaign-p360", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "campaign-p360.html"));
});

app.get("/appsflyer-highlight", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "appsflyer-highlight.html"));
});

app.get("/api/download/:token", async (req, res) => {
  const { token } = req.params;
  const download = generatedDownloads.get(token);

  if (!download) {
    return res.status(404).json({
      error: "Download expired. Please generate the ZIP again.",
    });
  }

  clearTimeout(download.timeout);

  return res.download(download.zipPath, download.filename, async () => {
    await cleanupGeneratedDownload(token);
  });
});

app.use(express.static(path.join(__dirname, "public")));

app.post(
  "/api/extract-columns",
  upload.single("sourceFile"),
  async (req, res) => {
    const sourceFile = req.file;
    const cleanupTargets = [];

    try {
      if (!sourceFile) {
        return res.status(400).json({ error: "A source file is required." });
      }

      cleanupTargets.push(sourceFile.path);

      const { rows } = await readMatrixFromFile(sourceFile);
      const headerInfo = extractHeaderInfo(rows, "AppsFlyer ID");

      if (!headerInfo) {
        throw new Error("Could not detect a header row in the uploaded file.");
      }

      await cleanupPaths(cleanupTargets);

      return res.json({
        columns: headerInfo.columns,
        headerRowNumber: headerInfo.headerRowIndex + 1,
        suggestedColumn:
          headerInfo.columns.find(
            (column) => normalizeHeader(column) === "appsflyer id",
          ) || headerInfo.columns[0] || "",
      });
    } catch (error) {
      await cleanupPaths(cleanupTargets);

      return res
        .status(500)
        .json({ error: error.message || "Something went wrong." });
    }
  },
);

app.post(
  "/api/generate",
  upload.fields([
    { name: "installFile", maxCount: 1 },
    { name: "inAppFile", maxCount: 1 },
  ]),
  async (req, res) => {
    const installFile = req.files?.installFile?.[0];
    const inAppFile = req.files?.inAppFile?.[0];
    const cleanupTargets = [];

    try {
      const appName = sanitizeFilePart(req.body.appName);

      if (!appName) {
        return res.status(400).json({ error: "App name is required." });
      }

      if (!installFile || !inAppFile) {
        return res
          .status(400)
          .json({ error: "Both Install and InApp files are required." });
      }

      cleanupTargets.push(installFile.path, inAppFile.path);

      const outputDirectory = await fsp.mkdtemp(
        path.join(os.tmpdir(), "campaign-p360-output-"),
      );

      cleanupTargets.push(outputDirectory);

      const store = createCampaignWorkbookStore(outputDirectory, appName);

      const installSummary = await processInputFile(installFile, "install", store);
      const inAppSummary = await processInputFile(inAppFile, "inapp", store);

      const generatedFiles = await store.finalize();

      if (generatedFiles.length === 0) {
        throw new Error(
          "No campaign IDs were found. Make sure the Campaign column has values like mobisaturn_1216.",
        );
      }

      const zipFilename = "Campaign - p360.zip";
      const zipPath = path.join(outputDirectory, zipFilename);

      await createZipFile(zipPath, generatedFiles);

      const summaryRows = generatedFiles.map((file) => ({
        fileName: file.filename.replace(/\.xlsx$/i, ""),
        headerRows: 1,
        installRows: file.installCount,
        inAppRows: file.inAppCount,
        totalRows: 1 + file.installCount + file.inAppCount,
      }));
      const totals = summaryRows.reduce(
        (accumulator, row) => ({
          headerRows: accumulator.headerRows + row.headerRows,
          installRows: accumulator.installRows + row.installRows,
          inAppRows: accumulator.inAppRows + row.inAppRows,
          totalRows: accumulator.totalRows + row.totalRows,
        }),
        { headerRows: 0, installRows: 0, inAppRows: 0, totalRows: 0 },
      );
      const downloadToken = registerGeneratedDownload(
        zipPath,
        cleanupTargets,
        zipFilename,
      );

      return res.json({
        zipFilename,
        downloadUrl: `/api/download/${downloadToken}`,
        inputTotals: {
          installRows: installSummary.dataRowCount,
          inAppRows: inAppSummary.dataRowCount,
        },
        outputTotals: totals,
        rows: summaryRows,
      });
    } catch (error) {
      await cleanupPaths(cleanupTargets);

      if (!res.headersSent) {
        res.status(500).json({ error: error.message || "Something went wrong." });
      } else {
        res.destroy(error);
      }
    }
  },
);

app.post(
  "/api/highlight-appsflyer",
  upload.single("sourceFile"),
  async (req, res) => {
    const sourceFile = req.file;
    const cleanupTargets = [];

    try {
      if (!sourceFile) {
        return res.status(400).json({ error: "A source file is required." });
      }

      cleanupTargets.push(sourceFile.path);

      const { rows, sheetName } = await readMatrixFromFile(sourceFile);
      const columnName = String(req.body.columnName || "AppsFlyer ID").trim();
      const headerInfo = extractHeaderInfo(rows, columnName);

      if (!columnName) {
        throw new Error("Column name is required.");
      }

      if (!headerInfo) {
        throw new Error(
          `The file must contain a row with a "${columnName}" column header.`,
        );
      }

      const {
        headerRowIndex,
        columnCount,
        paddedHeaderRow,
      } = headerInfo;
      const selectedColumnIndex = paddedHeaderRow.findIndex(
        (value) => normalizeHeader(value) === normalizeHeader(columnName),
      );

      if (selectedColumnIndex === -1) {
        throw new Error(
          `The file must contain a row with a "${columnName}" column header.`,
        );
      }

      const introRows = rows.slice(0, headerRowIndex).map((row) =>
        padRow(row, columnCount),
      );
      const dataRows = rows
        .slice(headerRowIndex + 1)
        .map((row, originalIndex) => ({
          originalIndex,
          values: padRow(row, columnCount),
        }))
        .filter((row) =>
          row.values.some((value) => String(value ?? "").trim() !== ""),
        );

      const counts = new Map();

      for (const row of dataRows) {
        const id = String(row.values[selectedColumnIndex] ?? "").trim();

        if (!id) {
          continue;
        }

        counts.set(id, (counts.get(id) || 0) + 1);
      }

      const duplicateIds = new Set(
        Array.from(counts.entries())
          .filter(([, count]) => count > 1)
          .map(([id]) => id),
      );

      dataRows.sort((left, right) => {
        const leftId = String(left.values[selectedColumnIndex] ?? "").trim();
        const rightId = String(right.values[selectedColumnIndex] ?? "").trim();

        if (!leftId && !rightId) {
          return left.originalIndex - right.originalIndex;
        }

        if (!leftId) {
          return 1;
        }

        if (!rightId) {
          return -1;
        }

        return (
          leftId.localeCompare(rightId, undefined, { numeric: true }) ||
          left.originalIndex - right.originalIndex
        );
      });

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet(sheetName || "Sheet1");

      for (const row of introRows) {
        worksheet.addRow(row);
      }

      const headerExcelRow = worksheet.addRow(paddedHeaderRow);
      headerExcelRow.font = { bold: true };
      headerExcelRow.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF3E8D8" },
        };
      });

      for (const row of dataRows) {
        const excelRow = worksheet.addRow(row.values);
        const id = String(row.values[selectedColumnIndex] ?? "").trim();

        if (duplicateIds.has(id)) {
          excelRow.eachCell({ includeEmpty: true }, (cell) => {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFFDE68A" },
            };
          });
        }
      }

      worksheet.autoFilter = {
        from: { row: headerExcelRow.number, column: 1 },
        to: { row: headerExcelRow.number, column: columnCount },
      };
      worksheet.views = [{ state: "frozen", ySplit: headerExcelRow.number }];

      const outputDirectory = await fsp.mkdtemp(
        path.join(os.tmpdir(), "appsflyer-highlight-"),
      );
      const originalBaseName =
        path.basename(
          sourceFile.originalname,
          path.extname(sourceFile.originalname || ""),
        ) || "appsflyer";
      const outputFilename = `${originalBaseName} - highlighted.xlsx`;
      const outputPath = path.join(outputDirectory, outputFilename);

      cleanupTargets.push(outputDirectory);

      await workbook.xlsx.writeFile(outputPath);

      res.download(outputPath, outputFilename, async () => {
        await cleanupPaths(cleanupTargets);
      });
    } catch (error) {
      await cleanupPaths(cleanupTargets);

      if (!res.headersSent) {
        res.status(500).json({ error: error.message || "Something went wrong." });
      }
    }
  },
);

app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === "LIMIT_FILE_SIZE") {
      return res.status(413).json({
        error: "Upload too large. Each file can be up to 1.5 GB.",
      });
    }

    return res.status(400).json({ error: error.message });
  }

  if (error) {
    return res
      .status(500)
      .json({ error: error.message || "Something went wrong." });
  }

  return next();
});

if (require.main === module) {
  app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
  });
}

// For serverless platforms like Vercel, export the app request handler directly.
module.exports = app;

// Attach helpers for local testing if needed.
module.exports.cleanupPaths = cleanupPaths;
module.exports.createCampaignWorkbookStore = createCampaignWorkbookStore;
module.exports.extractCampaignId = extractCampaignId;
module.exports.normalizeCellValue = normalizeCellValue;
module.exports.processInputFile = processInputFile;
