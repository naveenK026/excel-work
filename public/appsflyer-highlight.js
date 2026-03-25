const highlightForm = document.getElementById("highlight-form");
const highlightStatus = document.getElementById("highlightStatus");
const highlightButton = document.getElementById("highlightButton");

function setHighlightStatus(message, tone = "info") {
  highlightStatus.textContent = message;
  highlightStatus.dataset.tone = tone;
}

highlightForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(highlightForm);
  highlightButton.disabled = true;
  setHighlightStatus("Sorting rows and highlighting duplicate AppsFlyer IDs...", "info");

  try {
    const response = await fetch("/api/highlight-appsflyer", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      let errorMessage = "Unable to create the highlighted file.";

      try {
        const payload = await response.json();
        errorMessage = payload.error || errorMessage;
      } catch (error) {
        errorMessage = error.message || errorMessage;
      }

      throw new Error(errorMessage);
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const inputFile = document.getElementById("sourceFile");
    const originalName = inputFile.files[0]?.name || "appsflyer";
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
