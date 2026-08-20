/**
 * IOperationResult['error'] is typed as `Error | PowerDataRuntimeHttpError`, but the
 * latter is a plain `{ message, status?, requestId? }` object, not an Error subclass —
 * an `instanceof Error` check silently drops its message on every real HTTP/Dataverse
 * failure. Both shapes carry `.message`, so read that directly instead.
 *
 * For Dataverse/plugin failures, `.message` is often the raw OData error response body
 * (a JSON string with an `error.message` field, e.g. a duplicate-key or plugin error)
 * rather than a plain sentence — parse it out so the UI shows a readable message.
 */
export function operationErrorMessage(error: unknown, fallback: string): string {
  if (error && typeof error === 'object' && 'message' in error) {
    const message = (error as { message?: unknown }).message;
    if (typeof message === 'string' && message.trim()) {
      try {
        const parsed = JSON.parse(message) as { error?: { message?: string } };
        if (parsed.error?.message) return parsed.error.message;
      } catch {
        /* not JSON — use the raw message as-is */
      }
      return message;
    }
  }
  return fallback;
}
