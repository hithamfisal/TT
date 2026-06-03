export function downloadMicrosoftWorkbook(rawUrl: string): Promise<Buffer>;
export function handleDownloadWorkbook(
  req: { body?: { url?: string } },
  res: {
    setHeader(name: string, value: string): unknown;
    status(code: number): { json(payload: unknown): unknown };
    send(payload: Buffer): unknown;
  }
): Promise<void>;
