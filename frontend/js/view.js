document.addEventListener("DOMContentLoaded", async () => {
  const config = window.APP_CONFIG || {};
  const websiteUrl = config.websiteUrl || "https://solscada.tech";
  const websiteDisplayMs = config.websiteDisplayMs || 30000;
  const slideDisplayMs = config.autoplayIntervalMs || 5000;
  const autoFullscreen = config.autoFullscreen !== false;

  const player = document.getElementById("slideshow-player");
  const overlay = document.getElementById("viewer-overlay");
  const websiteSlide = document.getElementById("website-slide");
  const imageSlide = document.getElementById("image-slide");
  const websiteFrame = document.getElementById("website-frame");
  const slideImage = document.getElementById("slide-image");
  const prevBtn = document.getElementById("prev-slide");
  const nextBtn = document.getElementById("next-slide");
  const autoplayBtn = document.getElementById("autoplay-toggle");
  const fullscreenBtn = document.getElementById("fullscreen-btn");
  const exitFullscreenBtn = document.getElementById("exit-fullscreen-btn");
  const counterEl = document.getElementById("slide-counter");
  const statusEl = document.getElementById("status");
  const metaEl = document.getElementById("meta");

  let playlist = [];
  let playlistIndex = 0;
  let advanceTimer = null;
  let autoplayEnabled = true;
  let slideCount = 0;
  let websiteLoaded = false;

  function setStatus(message, type) {
    statusEl.textContent = message;
    statusEl.className = `status view-chrome ${type || ""}`;
  }

  function showOverlay(message) {
    overlay.textContent = message;
    overlay.classList.remove("hidden");
    websiteSlide.classList.remove("active");
    imageSlide.classList.remove("active");
  }

  function loadWebsiteOnce() {
    if (!websiteLoaded) {
      websiteFrame.src = websiteUrl;
      websiteLoaded = true;
    }
  }

  function showWebsite() {
    loadWebsiteOnce();
    websiteSlide.classList.add("active");
    imageSlide.classList.remove("active");
  }

  function showSlideImage(url) {
    websiteSlide.classList.remove("active");
    imageSlide.classList.add("active");
    slideImage.src = url;
  }

  function hideOverlay() {
    overlay.classList.add("hidden");
  }

  function buildPlaylist(slideUrls) {
    playlist = [
      {
        type: "website",
        url: websiteUrl,
        durationMs: websiteDisplayMs,
        label: "SolScada"
      }
    ];

    slideUrls.forEach((url, index) => {
      playlist.push({
        type: "slide",
        url,
        durationMs: slideDisplayMs,
        label: `Slide ${index + 1}`,
        slideNumber: index + 1
      });
    });
  }

  function updateCounter() {
    if (playlist.length === 0) {
      counterEl.textContent = "";
      return;
    }

    const item = playlist[playlistIndex];
    const seconds = item.durationMs / 1000;

    if (item.type === "website") {
      counterEl.textContent = `${item.label} · ${seconds}s`;
      return;
    }

    counterEl.textContent = `${item.label} of ${slideCount} · ${seconds}s`;
  }

  function showPlaylistItem(index) {
    if (playlist.length === 0) {
      return;
    }

    playlistIndex = (index + playlist.length) % playlist.length;
    const item = playlist[playlistIndex];

    if (item.type === "website") {
      showWebsite();
    } else {
      showSlideImage(item.url);
    }

    hideOverlay();
    updateCounter();
  }

  function stopAutoplay() {
    if (advanceTimer) {
      clearTimeout(advanceTimer);
      advanceTimer = null;
    }
  }

  function scheduleNext() {
    stopAutoplay();

    if (!autoplayEnabled || playlist.length === 0) {
      return;
    }

    const item = playlist[playlistIndex];
    advanceTimer = setTimeout(() => {
      showPlaylistItem(playlistIndex + 1);
      scheduleNext();
    }, item.durationMs);
  }

  function setAutoplay(enabled) {
    autoplayEnabled = enabled;
    autoplayBtn.textContent = enabled ? "Pause" : "Play";
    autoplayBtn.classList.toggle("active", enabled);

    if (enabled) {
      scheduleNext();
    } else {
      stopAutoplay();
    }
  }

  function goToItem(index) {
    showPlaylistItem(index);
    scheduleNext();
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
    goToItem(playlistIndex - 1);
  });

  nextBtn.addEventListener("click", () => {
    goToItem(playlistIndex + 1);
  });

  autoplayBtn.addEventListener("click", () => {
    setAutoplay(!autoplayEnabled);
  });

  fullscreenBtn.addEventListener("click", async () => {
    try {
      await enterFullscreen();
    } catch {
      setStatus("Fullscreen is not supported in this browser.", "error");
    }
  });

  exitFullscreenBtn.addEventListener("click", async () => {
    await exitFullscreen();
  });

  document.addEventListener("fullscreenchange", handleFullscreenChange);

  async function preloadSlides(slideUrls) {
    await Promise.all(
      slideUrls.map(
        (url) =>
          new Promise((resolve, reject) => {
            const image = new Image();
            image.onload = resolve;
            image.onerror = () => reject(new Error(`Failed to load slide: ${url}`));
            image.src = url;
          })
      )
    );
  }

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

      const slideUrls = viewInfo.slides.map((slide) => slide.url);
      slideCount = slideUrls.length;
      buildPlaylist(slideUrls);

      const sizeMb = (viewInfo.fileSizeBytes / (1024 * 1024)).toFixed(2);
      const modified = viewInfo.lastModifiedUtc
        ? new Date(viewInfo.lastModifiedUtc).toLocaleString()
        : "unknown";

      metaEl.textContent =
        `${viewInfo.fileName} · ${slideCount} slides · ${sizeMb} MB · updated ${modified} · ` +
        `cycle: SolScada ${websiteDisplayMs / 1000}s → slides ${slideDisplayMs / 1000}s each`;

      await preloadSlides(slideUrls);

      showPlaylistItem(0);
      setStatus(
        `Loop: ${websiteUrl} (${websiteDisplayMs / 1000}s) → ${slideCount} slides (${slideDisplayMs / 1000}s each).`,
        "success"
      );
      setAutoplay(true);

      if (autoFullscreen) {
        try {
          await enterFullscreen();
        } catch {
          // Fullscreen may require a user gesture in some browsers.
        }
      }
    } catch (error) {
      showOverlay("Unable to display presentation.");
      setStatus(error.message || "Failed to load presentation.", "error");
    }
  }

  init();
});
