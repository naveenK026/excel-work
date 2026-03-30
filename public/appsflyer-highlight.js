const highlightForm = document.getElementById("highlight-form");
const highlightStatus = document.getElementById("highlightStatus");
const highlightButton = document.getElementById("highlightButton");
const sourceFileInput = document.getElementById("sourceFile");
const columnNameInput = document.getElementById("columnName");
const columnOptions = document.getElementById("columnOptions");
const columnHint = document.getElementById("columnHint");

function setHighlightStatus(message, tone = "info") {
  highlightStatus.textContent = message;
  highlightStatus.dataset.tone = tone;
}

async function readErrorMessage(response, fallbackMessage) {
  const text = await response.text();

  if (!text) {
    return fallbackMessage;
  }

  try {
    const payload = JSON.parse(text);
    return payload.error || payload.message || fallbackMessage;
  } catch {
    return text;
  }
}

function populateColumnOptions(columns) {
  columnOptions.innerHTML = "";

  for (const column of columns) {
    const option = document.createElement("option");
    option.value = column;
    columnOptions.appendChild(option);
  }
}

async function detectColumns() {
  const file = sourceFileInput.files[0];

  if (!file) {
    populateColumnOptions([]);
    columnHint.textContent =
      "Upload a file to auto-detect headers. You can also type the column name manually.";
    return;
  }

  const formData = new FormData();
  formData.append("sourceFile", file);

  columnHint.textContent = "Detecting column names...";

  try {
    const response = await fetch("/api/extract-columns", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      throw new Error(
        await readErrorMessage(
          response,
          "Could not detect columns from the uploaded file.",
        ),
      );
    }

    const payload = await response.json();
    populateColumnOptions(payload.columns || []);

    if (payload.suggestedColumn) {
      columnNameInput.value = payload.suggestedColumn;
    }

    columnHint.textContent = `Detected ${payload.columns.length} columns from header row ${payload.headerRowNumber}.`;
  } catch (error) {
    populateColumnOptions([]);
    columnHint.textContent =
      "Could not detect headers automatically. You can still type the column name manually.";
  }
}

sourceFileInput.addEventListener("change", () => {
  detectColumns();
});

highlightForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(highlightForm);
  highlightButton.disabled = true;
  setHighlightStatus(
    `Sorting rows and highlighting duplicate values in "${columnNameInput.value}"...`,
    "info",
  );

  try {
    const response = await fetch("/api/highlight-appsflyer", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const errorMessage = await readErrorMessage(
        response,
        "Unable to create the highlighted file.",
      );

      throw new Error(errorMessage);
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const originalName = sourceFileInput.files[0]?.name || "appsflyer";
    const cleanName = originalName.replace(/\.[^.]+$/, "");

    link.href = url;
    link.download = `${cleanName} - highlighted.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);

    setHighlightStatus(
      "Highlighted workbook created successfully. Your download should start automatically.",
      "success",
    );
  } catch (error) {
    setHighlightStatus(error.message || "Something went wrong.", "error");
  } finally {
    highlightButton.disabled = false;
  }
});
