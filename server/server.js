const express = require("express");
const path = require("path");
const { createProxyMiddleware } = require("http-proxy-middleware");

const app = express();
const port = process.env.FRONTEND_PORT || 3000;
const backendUrl = process.env.BACKEND_URL || "http://localhost:5050";

app.use(
  "/api",
  createProxyMiddleware({
    target: backendUrl,
    changeOrigin: true,
    logLevel: "debug",
    pathRewrite: (path) => `/api${path}`,
    onProxyReq: (proxyReq, req) => {
      console.log(`Proxying ${req.method} ${req.originalUrl} -> ${backendUrl}${proxyReq.path}`);
    },
    onError: (err, req, res) => {
      console.error("Proxy error", err);
      if (res.headersSent) return;
      res.status(502).json({
        ok: false,
        error: "proxy_error",
        detail: String(err && err.message ? err.message : err),
      });
    },
  })
);

const rootDir = path.resolve(__dirname, "..");
app.use(express.static(rootDir));

app.get("*", (_req, res) => {
  res.sendFile(path.join(rootDir, "index.html"));
});

app.listen(port, () => {
  console.log(`Frontend dev server listening on ${port}`);
  console.log(`Proxying /api to ${backendUrl}`);
});
