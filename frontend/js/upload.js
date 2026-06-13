document.addEventListener("DOMContentLoaded", () => {
  const form = document.getElementById("upload-form");
  const fileInput = document.getElementById("file-input");
  const uploadButton = document.getElementById("upload-button");
  const statusEl = document.getElementById("status");
  const diagnosticsPanel = document.getElementById("diagnostics-panel");
  const diagnosticsOutput = document.getElementById("diagnostics-output");

  function setStatus(message, type) {
    statusEl.textContent = message;
    statusEl.className = `status ${type || ""}`;
  }

  function setDiagnostics(diagnostics) {
    if (!diagnosticsPanel || !diagnosticsOutput) {
      return;
    }

    if (!diagnostics || Object.keys(diagnostics).length === 0) {
      diagnosticsPanel.hidden = true;
      diagnosticsOutput.textContent = "";
      return;
    }

    diagnosticsPanel.hidden = false;
    diagnosticsPanel.open = true;
    diagnosticsOutput.textContent = Object.entries(diagnostics)
      .map(([key, value]) => `${key}: ${value ?? ""}`)
      .join("\n");
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();

    const file = fileInput.files?.[0];
    if (!file) {
      setStatus("Please choose a .pptx file first.", "error");
      return;
    }

    if (!file.name.toLowerCase().endsWith(".pptx")) {
      setStatus("Only .pptx files are allowed.", "error");
      return;
    }

    uploadButton.disabled = true;
    setStatus("Uploading presentation...", "loading");

    try {
      const result = await window.PresentationApi.upload(file);
      const sizeMb = (result.fileSizeBytes / (1024 * 1024)).toFixed(2);
      let message = `${result.message} (${result.fileName}, ${sizeMb} MB)`;

      if (result.slidesRendered) {
        message += ` · ${result.slideCount} slides rendered`;
      } else if (result.renderWarning) {
        message += ` · Warning: ${result.renderWarning}`;
      }

      setStatus(message, result.slidesRendered ? "success" : "error");
      setDiagnostics(result.slidesRendered ? null : result.renderDiagnostics);
      fileInput.value = "";
    } catch (error) {
      setStatus(error.message || "Upload failed.", "error");
      setDiagnostics(null);
    } finally {
      uploadButton.disabled = false;
    }
  });
});
