const isLocalhost = window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";
const apiBaseUrl = isLocalhost ? "http://localhost:5025" : "http://192.168.15.100:75";

window.APP_CONFIG = {
  apiBaseUrl,
  websiteUrl: "https://solscada.tech",
  websiteDisplayMs: 30000,
  autoplayIntervalMs: 5000,
  autoFullscreen: true
};
