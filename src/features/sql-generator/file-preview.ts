export function buildPreviewTemplateHref(templateKey: string) {
  return `/preview/${encodeURIComponent(templateKey)}`;
}

export function buildPreviewParamsHref(templateKey: string) {
  return `/preview/${encodeURIComponent(templateKey)}/params`;
}

export function buildPreviewFileHref(templateKey: string, filePath: string) {
  const encodedPath = filePath
    .split("/")
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join("/");

  return `/preview/${encodeURIComponent(templateKey)}/files/${encodedPath}`;
}

export function openPreviewRoute(href: string) {
  const win = window.open(href, "_blank", "noopener,noreferrer");
  return Boolean(win);
}
