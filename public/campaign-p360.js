const form = document.getElementById("generator-form");
const statusBox = document.getElementById("status");
const submitButton = document.getElementById("submitButton");
const summarySection = document.getElementById("summarySection");
const summaryTableBody = document.getElementById("summaryTableBody");
const installInputTotal = document.getElementById("installInputTotal");
const inAppInputTotal = document.getElementById("inAppInputTotal");
const installOutputTotal = document.getElementById("installOutputTotal");
const inAppOutputTotal = document.getElementById("inAppOutputTotal");
const headerTotalCell = document.getElementById("headerTotalCell");
const installTotalCell = document.getElementById("installTotalCell");
const inAppTotalCell = document.getElementById("inAppTotalCell");
const grandTotalCell = document.getElementById("grandTotalCell");
const downloadNote = document.getElementById("downloadNote");

function setStatus(message, tone = "info") {
  statusBox.textContent = message;
  statusBox.dataset.tone = tone;
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

function renderSummary(payload) {
  summaryTableBody.innerHTML = "";

  for (const row of payload.rows) {
    const tableRow = document.createElement("tr");
    tableRow.innerHTML = `
      <td>${row.fileName}</td>
      <td>${row.headerRows}</td>
      <td>${row.installRows}</td>
      <td>${row.inAppRows}</td>
      <td>${row.totalRows}</td>
    `;
    summaryTableBody.appendChild(tableRow);
  }

  installInputTotal.textContent = payload.inputTotals.installRows;
  inAppInputTotal.textContent = payload.inputTotals.inAppRows;
  installOutputTotal.textContent = payload.outputTotals.installRows;
  inAppOutputTotal.textContent = payload.outputTotals.inAppRows;
  headerTotalCell.textContent = payload.outputTotals.headerRows;
  installTotalCell.textContent = payload.outputTotals.installRows;
  inAppTotalCell.textContent = payload.outputTotals.inAppRows;
  grandTotalCell.textContent = payload.outputTotals.totalRows;
  downloadNote.textContent = `${payload.zipFilename} is ready and downloading automatically.`;
  summarySection.hidden = false;
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(form);
  submitButton.disabled = true;
  setStatus("Generating campaign workbooks. This can take a moment...", "info");

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const errorMessage = await readErrorMessage(
        response,
        "Unable to generate the ZIP file.",
      );

      throw new Error(errorMessage);
    }

    const payload = await response.json();
    renderSummary(payload);

    const link = document.createElement("a");
    link.href = payload.downloadUrl;
    document.body.appendChild(link);
    link.click();
    link.remove();

    setStatus(
      "ZIP generated successfully. Summary table is ready below.",
      "success",
    );
  } catch (error) {
    setStatus(error.message || "Something went wrong.", "error");
  } finally {
    submitButton.disabled = false;
  }
});
