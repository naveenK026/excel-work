const form = document.getElementById("generator-form");
const statusBox = document.getElementById("status");
const submitButton = document.getElementById("submitButton");

function setStatus(message, tone = "info") {
  statusBox.textContent = message;
  statusBox.dataset.tone = tone;
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
      let errorMessage = "Unable to generate the ZIP file.";

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

    link.href = url;
    link.download = "Campaign - p360.zip";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);

    setStatus(
      "ZIP generated successfully. Your download should start automatically.",
      "success",
    );
  } catch (error) {
    setStatus(error.message || "Something went wrong.", "error");
  } finally {
    submitButton.disabled = false;
  }
});
