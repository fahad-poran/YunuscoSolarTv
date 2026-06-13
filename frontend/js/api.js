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

  window.PresentationApi = {
    apiBaseUrl,
    getApiUrl,

    async upload(file) {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch(getApiUrl("/api/presentation/upload"), {
        method: "POST",
        body: formData
      });

      if (!response.ok) {
        throw new Error(await parseError(response));
      }

      return response.json();
    },

    async getViewInfo() {
      const response = await fetch(getApiUrl("/api/presentation/view"));

      if (!response.ok) {
        throw new Error(await parseError(response));
      }

      return response.json();
    },

    async fetchPresentationBlob() {
      const response = await fetch(getApiUrl("/api/presentation/file"));

      if (!response.ok) {
        throw new Error(await parseError(response));
      }

      return response.blob();
    }
  };
})();
