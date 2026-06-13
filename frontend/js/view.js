document.addEventListener("DOMContentLoaded", async () => {
  const config = window.APP_CONFIG || {};
  const autoplayIntervalMs = config.autoplayIntervalMs || 5000;
  const autoFullscreen = config.autoFullscreen !== false;

  const player = document.getElementById("slideshow-player");
  const overlay = document.getElementById("viewer-overlay");
  const slideImage = document.getElementById("slide-image");
  const prevBtn = document.getElementById("prev-slide");
  const nextBtn = document.getElementById("next-slide");
  const autoplayBtn = document.getElementById("autoplay-toggle");
  const fullscreenBtn = document.getElementById("fullscreen-btn");
  const exitFullscreenBtn = document.getElementById("exit-fullscreen-btn");
  const counterEl = document.getElementById("slide-counter");
  const statusEl = document.getElementById("status");
  const metaEl = document.getElementById("meta");

  let slides = [];
  let currentIndex = 0;
  let autoplayTimer = null;
  let autoplayEnabled = true;

  function setStatus(message, type) {
    statusEl.textContent = message;
    statusEl.className = `status view-chrome ${type || ""}`;
  }

  function showOverlay(message) {
    overlay.textContent = message;
    overlay.classList.remove("hidden");
    slideImage.hidden = true;
  }

  function hideOverlay() {
    overlay.classList.add("hidden");
    slideImage.hidden = false;
  }

  function updateCounter() {
    if (slides.length === 0) {
      counterEl.textContent = "";
      return;
    }

    counterEl.textContent = `Slide ${currentIndex + 1} of ${slides.length}`;
  }

  function showSlide(index) {
    if (slides.length === 0) {
      return;
    }

    currentIndex = (index + slides.length) % slides.length;
    slideImage.src = slides[currentIndex];
    updateCounter();
  }

  function stopAutoplay() {
    if (autoplayTimer) {
      clearInterval(autoplayTimer);
      autoplayTimer = null;
    }
  }

  function startAutoplay() {
    stopAutoplay();

    if (!autoplayEnabled || slides.length <= 1) {
      return;
    }

    autoplayTimer = setInterval(() => {
      showSlide(currentIndex + 1);
    }, autoplayIntervalMs);
  }

  function setAutoplay(enabled) {
    autoplayEnabled = enabled;
    autoplayBtn.textContent = enabled ? "Pause" : "Play";
    autoplayBtn.classList.toggle("active", enabled);

    if (enabled) {
      startAutoplay();
    } else {
      stopAutoplay();
    }
  }

  function restartAutoplayTimer() {
    if (autoplayEnabled) {
      startAutoplay();
    }
  }

  async function enterFullscreen() {
    const target = player;
    if (target.requestFullscreen) {
      await target.requestFullscreen();
    } else if (target.webkitRequestFullscreen) {
      await target.webkitRequestFullscreen();
    }
  }

  async function exitFullscreen() {
    if (document.fullscreenElement) {
      await document.exitFullscreen();
    }
  }

  function handleFullscreenChange() {
    const isFullscreen = Boolean(document.fullscreenElement);
    document.body.classList.toggle("is-fullscreen", isFullscreen);
    exitFullscreenBtn.hidden = !isFullscreen;
  }

  prevBtn.addEventListener("click", () => {
    showSlide(currentIndex - 1);
    restartAutoplayTimer();
  });

  nextBtn.addEventListener("click", () => {
    showSlide(currentIndex + 1);
    restartAutoplayTimer();
  });

  autoplayBtn.addEventListener("click", () => {
    setAutoplay(!autoplayEnabled);
  });

  fullscreenBtn.addEventListener("click", async () => {
    try {
      await enterFullscreen();
    } catch (error) {
      setStatus("Fullscreen is not supported in this browser.", "error");
    }
  });

  exitFullscreenBtn.addEventListener("click", async () => {
    await exitFullscreen();
  });

  document.addEventListener("fullscreenchange", handleFullscreenChange);

  async function init() {
    showOverlay("Loading presentation...");
    setStatus("Fetching slides...", "loading");

    try {
      const viewInfo = await window.PresentationApi.getViewInfo();

      if (!viewInfo.exists) {
        showOverlay("No presentation uploaded yet.");
        setStatus("Upload a .pptx file on the Upload page first.", "error");
        prevBtn.disabled = true;
        nextBtn.disabled = true;
        autoplayBtn.disabled = true;
        fullscreenBtn.disabled = true;
        return;
      }

      if (!viewInfo.slidesAvailable || !viewInfo.slides?.length) {
        showOverlay("Slides are not ready.");
        setStatus(
          "Re-upload the presentation. PowerPoint must be installed on the API server to export slide images.",
          "error"
        );
        prevBtn.disabled = true;
        nextBtn.disabled = true;
        autoplayBtn.disabled = true;
        fullscreenBtn.disabled = true;
        return;
      }

      slides = viewInfo.slides.map((slide) => slide.url);
      const sizeMb = (viewInfo.fileSizeBytes / (1024 * 1024)).toFixed(2);
      const modified = viewInfo.lastModifiedUtc
        ? new Date(viewInfo.lastModifiedUtc).toLocaleString()
        : "unknown";

      metaEl.textContent = `${viewInfo.fileName} · ${viewInfo.slideCount} slides · ${sizeMb} MB · updated ${modified}`;

      await new Promise((resolve, reject) => {
        const image = new Image();
        image.onload = resolve;
        image.onerror = () => reject(new Error("Failed to load the first slide image."));
        image.src = slides[0];
      });

      showSlide(0);
      hideOverlay();
      setStatus(`Autoplay every ${autoplayIntervalMs / 1000} seconds.`, "success");
      setAutoplay(true);

      if (autoFullscreen) {
        try {
          await enterFullscreen();
        } catch {
          // Fullscreen requires a user gesture in some browsers.
        }
      }
    } catch (error) {
      showOverlay("Unable to display presentation.");
      setStatus(error.message || "Failed to load presentation.", "error");
    }
  }

  init();
});
