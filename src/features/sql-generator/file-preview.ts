import type { GeneratedPair, GeneratedFile } from "@/lib/sql-generator";

export function escapeHtml(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function openHtmlPreview(title: string, blocks: string) {
  const win = window.open("", "_blank");
  if (!win) return false;

  const html = `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(title)}</title>
  <style>
    body { font-family: Consolas, monospace; margin: 16px; background: #f7f7f7; color: #0f172a; }
    h1, h2 { font-family: Arial, sans-serif; }
    pre { background: #fff; padding: 12px; border: 1px solid #dbe4ea; border-radius: 10px; overflow: auto; white-space: pre-wrap; }
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  ${blocks}
</body>
</html>`;

  win.document.open();
  win.document.write(html);
  win.document.close();
  return true;
}

export function renderFilesPreview(files: GeneratedFile[]) {
  return files
    .map((file) => `<h2>${escapeHtml(file.path)}</h2><pre>${escapeHtml(file.content)}</pre>`)
    .join("\n");
}

export function renderPairPreview(pair: GeneratedPair) {
  return pair.files
    .map((file) => `<h2>${escapeHtml(file.path)}</h2><pre>${escapeHtml(file.content)}</pre>`)
    .join("\n");
}
