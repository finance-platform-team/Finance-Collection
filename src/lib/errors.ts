/**
 * IOperationResult['error'] is typed as `Error | PowerDataRuntimeHttpError`, but the
 * latter is a plain `{ message, status?, requestId? }` object, not an Error subclass —
 * an `instanceof Error` check silently drops its message on every real HTTP/Dataverse
 * failure. Both shapes carry `.message`, so read that directly instead.
 */
export function operationErrorMessage(error: unknown, fallback: string): string {
  if (error && typeof error === 'object' && 'message' in error) {
    const message = (error as { message?: unknown }).message;
    if (typeof message === 'string' && message.trim()) return message;
  }
  return fallback;
}
