const form = document.getElementById("zip-form");
const statusBox = document.getElementById("status");
const submitButton = document.getElementById("submitButton");
const folderInput = document.getElementById("folderInput");

function setStatus(message, tone = "info") {
  statusBox.textContent = message;
  statusBox.dataset.tone = tone;
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const files = Array.from(folderInput.files);
  if (files.length === 0) {
    setStatus("Please select a folder.", "error");
    return;
  }

  submitButton.disabled = true;
  setStatus("Zipping folder...", "info");

  try {
    const folderName = files[0].webkitRelativePath.split("/")[0];
    const zip = new JSZip();

    for (const file of files) {
      zip.file(file.webkitRelativePath, file);
    }

    const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${folderName}.zip`;
    document.body.appendChild(link);
    link.click();
    link.remove();

    setStatus(`${folderName}.zip downloaded successfully.`, "success");
  } catch (error) {
    setStatus(error.message || "Something went wrong.", "error");
  } finally {
    submitButton.disabled = false;
  }
});
