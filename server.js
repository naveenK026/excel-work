const fs = require("fs");
const fsp = require("fs/promises");
const os = require("os");
const path = require("path");
const express = require("express");
const multer = require("multer");
const archiver = require("archiver");
const csvParser = require("csv-parser");
const ExcelJS = require("exceljs");
const XlsxStreamReader = require("xlsx-stream-reader");

const app = express();
const port = process.env.PORT || 3000;
const uploadDirectory = path.join(os.tmpdir(), "campaign-p360-uploads");

fs.mkdirSync(uploadDirectory, { recursive: true });

const upload = multer({
  dest: uploadDirectory,
  limits: {
    fileSize: 600 * 1024 * 1024,
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
  await new Promise((resolve, reject) => {
    let headers = null;
    let campaignIndex = -1;
    let isSettled = false;

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
        resolve();
      }
    });

    parser.on("error", fail);
    stream.on("error", fail);
    stream.pipe(parser);
  });
}

async function processXlsxFile(filePath, sourceType, store) {
  await new Promise((resolve, reject) => {
    let isSettled = false;
    let sawWorksheet = false;

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
        resolve();
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

async function processInputFile(file, sourceType, store) {
  const fileType = await detectFileType(file);

  if (fileType === "csv") {
    await processCsvFile(file.path, sourceType, store);
    return;
  }

  if (fileType === "xlsx") {
    await processXlsxFile(file.path, sourceType, store);
    return;
  }

  throw new Error("Only .xlsx and .csv files are supported.");
}

async function cleanupPaths(paths) {
  await Promise.all(
    paths.map((targetPath) =>
      fsp.rm(targetPath, { recursive: true, force: true }).catch(() => {}),
    ),
  );
}

app.use(express.static(path.join(__dirname, "public")));

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
    let didAttachCleanup = false;

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

      await processInputFile(installFile, "install", store);
      await processInputFile(inAppFile, "inapp", store);

      const generatedFiles = await store.finalize();

      if (generatedFiles.length === 0) {
        throw new Error(
          "No campaign IDs were found. Make sure the Campaign column has values like mobisaturn_1216.",
        );
      }

      res.setHeader("Content-Type", "application/zip");
      res.setHeader(
        "Content-Disposition",
        'attachment; filename="Campaign - p360.zip"',
      );

      const archive = archiver("zip", { zlib: { level: 0 } });
      const cleanup = () => cleanupPaths(cleanupTargets);

      if (!didAttachCleanup) {
        didAttachCleanup = true;
        res.on("finish", cleanup);
        res.on("close", cleanup);
      }

      archive.on("error", (error) => {
        res.destroy(error);
      });

      archive.pipe(res);

      for (const file of generatedFiles) {
        archive.file(file.fullPath, { name: file.filename });
      }

      await archive.finalize();
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

if (require.main === module) {
  app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
  });
}

module.exports = {
  app,
  cleanupPaths,
  createCampaignWorkbookStore,
  extractCampaignId,
  normalizeCellValue,
  processInputFile,
};
