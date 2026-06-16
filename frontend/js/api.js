(function () {
  const config = window.APP_CONFIG || {};
  const apiBaseUrl = (config.apiBaseUrl || "").replace(/\/$/, "");

  function getApiUrl(path) {
    return `${apiBaseUrl}${path}`;
  }

  async function parseError(response) {
    try {
      const data = await response.json();
      return data.details || data.error || data.title || response.statusText;
    } catch {
      return response.statusText || "Request failed.";
    }
  }

  async function request(path, options = {}) {
    const url = getApiUrl(path);

    try {
      const response = await fetch(url, options);
      if (!response.ok) {
        throw new Error(await parseError(response));
      }
      return response;
    } catch (error) {
      if (error instanceof TypeError && error.message === "Failed to fetch") {
        const origin = window.location.origin;
        const healthUrl = getApiUrl("/api/health");
        const msg =
          `The browser could not connect to ${url}. ` +
          `This usually means the backend is not running at ${apiBaseUrl}, the port is blocked, ` +
          `or CORS is rejecting this frontend origin (${origin}). ` +
          `Open ${healthUrl} directly on the server to confirm the API returns JSON, ` +
          `then ensure ${origin} is listed in Cors:AllowedOrigins.`;
        console.error(`[API Error] ${msg}`);
        throw new Error(msg);
      }
      console.error(`[API Error] ${path}:`, error);
      throw error;
    }
  }

  window.PresentationApi = {
    apiBaseUrl,
    getApiUrl,

    async upload(file) {
      const formData = new FormData();
      formData.append("file", file);

      const response = await request("/api/presentation/upload", {
        method: "POST",
        body: formData
      });

      return response.json();
    },

    async getViewInfo() {
      const response = await request("/api/presentation/view");
      return response.json();
    },

    async getHealth() {
      const response = await request("/api/health");
      return response.json();
    },

    async fetchPresentationBlob() {
      const response = await request("/api/presentation/file");
      return response.blob();
    }
  };
})();
