/** One function talks to the backend: API.call(action, args, { token, timeout }). */
(function () {
  const cfg = window.MR_CONFIG || {};

  async function call(action, args, opts) {
    opts = opts || {};
    if (!cfg.API_URL || cfg.API_URL.indexOf("PASTE_") === 0) {
      throw new Error("API_URL is not set. Open frontend/js/config.js and paste your Apps Script web app address.");
    }
    const ctl = typeof AbortController === "function" ? new AbortController() : null;
    const timer = ctl ? setTimeout(function () { ctl.abort(); }, opts.timeout || 90000) : null;
    let res;
    try {
      res = await fetch(cfg.API_URL, {
        method: "POST",
        // text/plain avoids a CORS preflight, which Apps Script cannot answer
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: action, args: args || [], token: opts.token || "" }),
        signal: ctl ? ctl.signal : undefined
      });
    } catch (e) {
      if (e && e.name === "AbortError") throw new Error("The server is taking too long. Wait a moment, then try again.");
      throw new Error("Cannot reach the server. Check your internet connection and the API_URL in js/config.js.");
    } finally { if (timer) clearTimeout(timer); }
    let json;
    try { json = await res.json(); }
    catch (e) { throw new Error("Unexpected reply from the server. Re-deploy the web app (Anyone can access) and make sure the address ends with /exec."); }
    if (!json.ok) throw new Error(json.error || "Request failed.");
    return json.data;
  }

  window.API = { call: call };
})();
