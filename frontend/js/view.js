document.addEventListener("DOMContentLoaded", async () => {
  const config = window.APP_CONFIG || {};
  const websiteUrl       = config.websiteUrl        || "https://solscada.tech";
  const websiteDisplayMs = config.websiteDisplayMs  || 30000;
  const slideDisplayMs   = config.autoplayIntervalMs || 30000;
  const autoFullscreen   = config.autoFullscreen !== false;

  const player            = document.getElementById("slideshow-player");
  const overlay           = document.getElementById("viewer-overlay");
  const websiteSlide      = document.getElementById("website-slide");
  const imageSlide        = document.getElementById("image-slide");
  const websiteFrame      = document.getElementById("website-frame");
  const slideImage        = document.getElementById("slide-image");
  const prevBtn           = document.getElementById("prev-slide");
  const nextBtn           = document.getElementById("next-slide");
  const autoplayBtn       = document.getElementById("autoplay-toggle");
  const fullscreenBtn     = document.getElementById("fullscreen-btn");
  const exitFullscreenBtn = document.getElementById("exit-fullscreen-btn");
  const counterEl         = document.getElementById("slide-counter");
  const statusEl          = document.getElementById("status");
  const metaEl            = document.getElementById("meta");
  const diagnosticsPanel  = document.getElementById("diagnostics-panel");
  const diagnosticsOutput = document.getElementById("diagnostics-output");

  let playlist        = [];
  let playlistIndex   = 0;
  let advanceTimer    = null;
  let autoplayEnabled = true;
  let slideCount      = 0;
  let websiteLoaded   = false;

  // ── Progress bar ───────────────────────────────────────────────────────────

  // Inject progress bar + pause indicator into player
  const progressTrack = document.createElement("div");
  progressTrack.className = "progress-bar-track";
  const progressFill = document.createElement("div");
  progressFill.className = "progress-bar-fill";
  progressTrack.appendChild(progressFill);
  player.appendChild(progressTrack);

  const pauseIndicator = document.createElement("div");
  pauseIndicator.className = "pause-indicator";
  pauseIndicator.textContent = "⏸";
  player.appendChild(pauseIndicator);

  function startProgress(durationMs) {
    // Reset instantly, then animate to 100% over durationMs
    progressFill.style.transition = "none";
    progressFill.style.width = "0%";
    progressFill.classList.remove("paused");

    requestAnimationFrame(() => requestAnimationFrame(() => {
      progressFill.style.transition = `width ${durationMs}ms linear`;
      progressFill.style.width = "100%";
    }));
  }

  function pauseProgress() {
    const computed = getComputedStyle(progressFill).width;
    const trackWidth = progressTrack.offsetWidth;
    const pct = trackWidth > 0 ? (parseFloat(computed) / trackWidth * 100) : 0;
    progressFill.style.transition = "none";
    progressFill.style.width = pct + "%";
    progressFill.classList.add("paused");
  }

  function resumeProgress(remainingMs) {
    progressFill.classList.remove("paused");
    requestAnimationFrame(() => requestAnimationFrame(() => {
      progressFill.style.transition = `width ${remainingMs}ms linear`;
      progressFill.style.width = "100%";
    }));
  }

  // ── Status / overlay ───────────────────────────────────────────────────────

  function setStatus(message, type) {
    statusEl.textContent = message;
    statusEl.className   = `status view-chrome ${type || ""}`;
  }

  function showOverlay(message) {
    overlay.textContent = message;
    overlay.classList.remove("hidden");
    websiteSlide.classList.remove("active");
    imageSlide.classList.remove("active");
  }

  function hideOverlay() { overlay.classList.add("hidden"); }

  function setDiagnostics(diagnostics) {
    if (!diagnosticsPanel || !diagnosticsOutput) return;
    if (!diagnostics || Object.keys(diagnostics).length === 0) {
      diagnosticsPanel.hidden = true;
      diagnosticsOutput.textContent = "";
      return;
    }
    diagnosticsPanel.hidden = false;
    diagnosticsPanel.open   = true;
    diagnosticsOutput.textContent = Object.entries(diagnostics)
      .map(([k, v]) => `${k}: ${v ?? ""}`)
      .join("\n");
  }

  // ── Slide display ──────────────────────────────────────────────────────────

  function loadWebsiteOnce() {
    if (!websiteLoaded) { websiteFrame.src = websiteUrl; websiteLoaded = true; }
  }

  function showWebsite() {
    loadWebsiteOnce();
    websiteSlide.classList.add("active");
    imageSlide.classList.remove("active");
  }

  function showSlideImage(url) {
    slideImage.src = url;
    imageSlide.classList.add("active");
    websiteSlide.classList.remove("active");
  }

  // ── Playlist ───────────────────────────────────────────────────────────────

  function buildPlaylist(slideUrls) {
    playlist = [
      { type: "website", url: websiteUrl, durationMs: websiteDisplayMs, label: "SolScada" }
    ];
    slideUrls.forEach((url, i) => playlist.push({
      type: "slide", url, durationMs: slideDisplayMs,
      label: `Slide ${i + 1}`, slideNumber: i + 1
    }));
  }

  function updateCounter() {
    if (!playlist.length) { counterEl.textContent = ""; return; }
    const item = playlist[playlistIndex];
    counterEl.textContent = item.type === "website"
      ? `${item.label} · ${item.durationMs / 1000}s`
      : `${item.label} of ${slideCount} · ${item.durationMs / 1000}s`;
  }

  function showPlaylistItem(index) {
    if (!playlist.length) return;
    playlistIndex = (index + playlist.length) % playlist.length;
    const item = playlist[playlistIndex];
    if (item.type === "website") showWebsite();
    else showSlideImage(item.url);
    hideOverlay();
    updateCounter();
    if (autoplayEnabled) startProgress(item.durationMs);
  }

  // ── Autoplay ───────────────────────────────────────────────────────────────

  let slideStartTime   = null;  // when current slide started (or resumed)
  let slideRemainingMs = 0;     // how much time was left when paused

  function stopAutoplay() {
    if (advanceTimer) { clearTimeout(advanceTimer); advanceTimer = null; }
  }

  function scheduleNext(durationMs) {
    stopAutoplay();
    if (!autoplayEnabled || !playlist.length) return;
    const ms = durationMs ?? playlist[playlistIndex].durationMs;
    slideStartTime   = Date.now();
    slideRemainingMs = ms;
    advanceTimer = setTimeout(() => {
      showPlaylistItem(playlistIndex + 1);
      scheduleNext();
    }, ms);
  }

  function setAutoplay(enabled) {
    autoplayEnabled = enabled;
    autoplayBtn.textContent = enabled ? "Pause" : "Play";
    autoplayBtn.classList.toggle("active", enabled);

    // Flash pause/play indicator
    pauseIndicator.textContent = enabled ? "▶" : "⏸";
    pauseIndicator.classList.add("show");
    setTimeout(() => pauseIndicator.classList.remove("show"), 800);

    if (enabled) {
      // Resume from exactly where we paused
      const elapsed   = slideStartTime ? Date.now() - slideStartTime : 0;
      const remaining = Math.max(500, slideRemainingMs - elapsed);
      scheduleNext(remaining);
      resumeProgress(remaining);
    } else {
      // Pause: snapshot how much time is left
      const elapsed    = slideStartTime ? Date.now() - slideStartTime : 0;
      slideRemainingMs = Math.max(0, slideRemainingMs - elapsed);
      slideStartTime   = null; // clear so elapsed doesn't keep growing
      stopAutoplay();
      pauseProgress();
    }
  }

  function goToItem(index) {
    showPlaylistItem(index);
    scheduleNext();
  }

  // ── Fullscreen ─────────────────────────────────────────────────────────────

  async function enterFullscreen() {
    if (player.requestFullscreen)            await player.requestFullscreen();
    else if (player.webkitRequestFullscreen) await player.webkitRequestFullscreen();
  }

  async function exitFullscreen() {
    if (document.fullscreenElement) await document.exitFullscreen();
  }

  function handleFullscreenChange() {
    const isFs = Boolean(document.fullscreenElement);
    document.body.classList.toggle("is-fullscreen", isFs);
    exitFullscreenBtn.hidden = !isFs;
  }

  // ── Keyboard & remote control ──────────────────────────────────────────────
  // TV remotes send standard media keys or arrow/ok keys depending on platform.
  // We handle both sets so it works on Chromecast, Fire Stick, Android TV, etc.

  document.addEventListener("keydown", (e) => {
    switch (e.key) {
      // Play / Pause — spacebar, Enter (OK on most remotes), MediaPlayPause
      case " ":
      case "MediaPlayPause":
        e.preventDefault();
        setAutoplay(!autoplayEnabled);
        break;

      // Next slide — ArrowRight, MediaTrackNext
      case "ArrowRight":
      case "MediaTrackNext":
        e.preventDefault();
        goToItem(playlistIndex + 1);
        break;

      // Previous slide — ArrowLeft, MediaTrackPrevious
      case "ArrowLeft":
      case "MediaTrackPrevious":
        e.preventDefault();
        goToItem(playlistIndex - 1);
        break;

      // Fullscreen toggle — F key or ArrowUp (some remotes)
      case "f":
      case "F":
        e.preventDefault();
        document.fullscreenElement ? exitFullscreen() : enterFullscreen();
        break;

      // Exit fullscreen — Escape (browser handles this natively too)
      case "Escape":
        // browser handles Escape to exit fullscreen automatically
        break;
    }
  });

  // ── Event listeners ────────────────────────────────────────────────────────

  prevBtn.addEventListener("click", () => goToItem(playlistIndex - 1));
  nextBtn.addEventListener("click", () => goToItem(playlistIndex + 1));
  autoplayBtn.addEventListener("click", () => setAutoplay(!autoplayEnabled));
  fullscreenBtn.addEventListener("click", async () => {
    try { await enterFullscreen(); }
    catch { setStatus("Fullscreen is not supported in this browser.", "error"); }
  });
  exitFullscreenBtn.addEventListener("click", () => exitFullscreen());
  document.addEventListener("fullscreenchange", handleFullscreenChange);

  // ── Preload ────────────────────────────────────────────────────────────────

  async function preloadSlides(slideUrls) {
    await Promise.all(slideUrls.map(url => new Promise((resolve, reject) => {
      const img   = new Image();
      img.onload  = resolve;
      img.onerror = () => reject(new Error(`Failed to load slide: ${url}`));
      img.src     = url;
    })));
  }

  // ── Init ───────────────────────────────────────────────────────────────────

  async function init() {
    showOverlay("Loading presentation...");
    setStatus("Fetching slides...", "loading");

    try {
      const viewInfo = await window.PresentationApi.getViewInfo();

      if (!viewInfo.exists) {
        showOverlay("No presentation uploaded yet.");
        setStatus("Upload a .pptx file on the Upload page first.", "error");
        setDiagnostics(null);
        [prevBtn, nextBtn, autoplayBtn, fullscreenBtn].forEach(b => b.disabled = true);
        return;
      }

      if (!viewInfo.slidesAvailable || !viewInfo.slides?.length) {
        showOverlay("Slides are not ready.");
        setStatus(
          viewInfo.renderWarning ||
          "Re-upload the presentation. PowerPoint must be installed on the API server to export slide images.",
          "error"
        );
        setDiagnostics(viewInfo.renderDiagnostics);
        [prevBtn, nextBtn, autoplayBtn, fullscreenBtn].forEach(b => b.disabled = true);
        return;
      }

      const slideUrls = viewInfo.slides.map(s => s.url);
      slideCount = slideUrls.length;
      buildPlaylist(slideUrls);

      const sizeMb   = (viewInfo.fileSizeBytes / (1024 * 1024)).toFixed(2);
      const modified = viewInfo.lastModifiedUtc
        ? new Date(viewInfo.lastModifiedUtc).toLocaleString() : "unknown";

      metaEl.textContent =
        `${viewInfo.fileName} · ${slideCount} slides · ${sizeMb} MB · updated ${modified} · ` +
        `cycle: SolScada ${websiteDisplayMs / 1000}s → slides ${slideDisplayMs / 1000}s each`;

      await preloadSlides(slideUrls);

      showPlaylistItem(0);
      setDiagnostics(null);
      setStatus(
        `Loop: ${websiteUrl} (${websiteDisplayMs/1000}s) → ${slideCount} slides (${slideDisplayMs/1000}s each). ` +
        `Keyboard: Space=play/pause  ◀▶=prev/next  F=fullscreen`,
        "success"
      );
      setAutoplay(true);

      if (autoFullscreen) {
        try { await enterFullscreen(); }
        catch { /* requires user gesture in some browsers */ }
      }
    } catch (error) {
      showOverlay("Unable to display presentation.");
      setStatus(error.message || "Failed to load presentation.", "error");
      setDiagnostics(null);
    }
  }

  init();
});