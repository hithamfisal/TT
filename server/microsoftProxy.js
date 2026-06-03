const MICROSOFT_ALLOWED_HOSTS = [
  "sharepoint.com",
  "onedrive.live.com",
  "1drv.ms",
  "office.com",
  "officeapps.live.com",
];

const MAX_WORKBOOK_BYTES = 80 * 1024 * 1024;

function isAllowedMicrosoftHost(hostname) {
  const host = hostname.toLowerCase();
  return MICROSOFT_ALLOWED_HOSTS.some(
    allowed => host === allowed || host.endsWith(`.${allowed}`)
  );
}

function microsoftShareId(value) {
  return `u!${Buffer.from(value.trim(), "utf8")
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "")}`;
}

function withDownloadParam(value) {
  const url = new URL(value.trim());
  url.searchParams.set("download", "1");
  return url.toString();
}

function validateMicrosoftWorkbookUrl(value) {
  if (!value || typeof value !== "string") {
    throw new Error("Missing Microsoft 365 workbook URL.");
  }

  const parsed = new URL(value.trim());
  if (parsed.protocol !== "https:") {
    throw new Error("Only HTTPS Microsoft 365 links are allowed.");
  }
  if (!isAllowedMicrosoftHost(parsed.hostname)) {
    throw new Error("Only Microsoft 365, OneDrive, and SharePoint links are allowed.");
  }
  return parsed.toString();
}

async function readWorkbookResponse(response) {
  if (!response.ok) {
    throw new Error(`${response.status} ${response.statusText}`.trim());
  }

  const contentLength = Number(response.headers.get("content-length") ?? 0);
  if (contentLength > MAX_WORKBOOK_BYTES) {
    throw new Error("Workbook is too large to download through the proxy.");
  }

  const contentType = response.headers.get("content-type") ?? "";
  const buffer = Buffer.from(await response.arrayBuffer());
  if (buffer.byteLength > MAX_WORKBOOK_BYTES) {
    throw new Error("Workbook is too large to download through the proxy.");
  }

  const isWorkbook =
    buffer[0] === 0x50 &&
    buffer[1] === 0x4b &&
    buffer[2] === 0x03 &&
    buffer[3] === 0x04;
  if (
    !isWorkbook ||
    contentType.includes("text/html") ||
    contentType.includes("application/json")
  ) {
    throw new Error(`not workbook (${contentType || "unknown content"})`);
  }

  return buffer;
}

export async function downloadMicrosoftWorkbook(rawUrl) {
  const url = validateMicrosoftWorkbookUrl(rawUrl);
  const candidates = [
    `https://api.onedrive.com/v1.0/shares/${microsoftShareId(url)}/root/content`,
    withDownloadParam(url),
  ];
  const errors = [];

  for (const candidate of candidates) {
    try {
      const response = await fetch(candidate, { redirect: "follow" });
      return await readWorkbookResponse(response);
    } catch (error) {
      errors.push(error instanceof Error ? error.message : String(error));
    }
  }

  throw new Error(errors.filter(Boolean).join(" | "));
}

export async function handleDownloadWorkbook(req, res) {
  try {
    const workbook = await downloadMicrosoftWorkbook(req.body?.url);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Cache-Control", "no-store");
    res.send(workbook);
  } catch (error) {
    res.status(400).json({
      error:
        "Could not download the Microsoft 365 workbook through the local proxy.",
      details: error instanceof Error ? error.message : String(error),
    });
  }
}
