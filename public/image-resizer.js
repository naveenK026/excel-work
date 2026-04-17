const form = document.getElementById("resizer-form");
const statusBox = document.getElementById("status");
const submitButton = document.getElementById("submitButton");
const summarySection = document.getElementById("summarySection");
const summaryBody = document.getElementById("summaryBody");
const folderInput = document.getElementById("folderInput");

function setStatus(message, tone = "info") {
  statusBox.textContent = message;
  statusBox.dataset.tone = tone;
}

async function readErrorMessage(response, fallback) {
  const text = await response.text();
  if (!text) return fallback;
  try {
    const payload = JSON.parse(text);
    return payload.error || payload.message || fallback;
  } catch {
    return text;
  }
}

function renderSummary(summary) {
  summaryBody.innerHTML = "";

  for (const [folder, files] of Object.entries(summary)) {
    const heading = document.createElement("p");
    heading.style.cssText = "font-weight:700;margin:14px 0 6px;";
    heading.textContent = folder;
    summaryBody.appendChild(heading);

    const list = document.createElement("ul");
    list.style.cssText = "margin:0;padding-left:18px;color:var(--muted);line-height:1.8;";
    for (const file of files) {
      const li = document.createElement("li");
      li.textContent = file;
      list.appendChild(li);
    }
    summaryBody.appendChild(list);
  }

  summarySection.hidden = false;
}

async function buildZipFromFiles(files) {
  const zip = new JSZip();

  for (const file of files) {
    // webkitRelativePath looks like: "parentFolder/subfolder/image.jpg"
    // We strip the top-level parent folder so the ZIP starts from subfolders
    const parts = file.webkitRelativePath.split("/");
    const zipPath = parts.slice(1).join("/"); // drop the root folder name

    if (!zipPath) continue; // skip root-level files

    zip.file(zipPath, file);
  }

  return zip.generateAsync({ type: "blob", compression: "STORE" });
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const files = Array.from(folderInput.files);

  if (files.length === 0) {
    setStatus("Please select a folder.", "error");
    return;
  }

  submitButton.disabled = true;
  setStatus("Packaging folder...", "info");

  try {
    const zipBlob = await buildZipFromFiles(files);

    setStatus("Resizing images — this may take a moment...", "info");

    const formData = new FormData();
    formData.append("zipFile", zipBlob, "creatives.zip");
    const folderName = files[0].webkitRelativePath.split("/")[0];
    formData.append("folderName", folderName);

    const response = await fetch("/api/resize-images", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      throw new Error(await readErrorMessage(response, "Unable to resize images."));
    }

    const payload = await response.json();
    renderSummary(payload.summary);

    const link = document.createElement("a");
    link.href = payload.downloadUrl;
    document.body.appendChild(link);
    link.click();
    link.remove();

    setStatus("Done! Your download should start automatically.", "success");
  } catch (error) {
    setStatus(error.message || "Something went wrong.", "error");
  } finally {
    submitButton.disabled = false;
  }
});
