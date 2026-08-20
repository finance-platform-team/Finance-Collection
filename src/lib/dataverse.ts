import { operationErrorMessage } from './errors';

/* Drains a skip-token-paginated Dataverse getAll() call into a single array —
 * shared by every page that lists more rows than one page can hold. */
export async function fetchAll<T>(
  run: (skipToken?: string) => Promise<{ success: boolean; data?: T[] | null; error?: unknown; skipToken?: string }>,
  fallback: string
): Promise<T[]> {
  let out: T[] = [];
  let token: string | undefined;
  let guard = 0;
  do {
    const res = await run(token);
    if (!res.success) throw new Error(operationErrorMessage(res.error, fallback));
    out = out.concat(res.data || []);
    token = res.skipToken;
    guard++;
  } while (token && guard < 50);
  return out;
}
