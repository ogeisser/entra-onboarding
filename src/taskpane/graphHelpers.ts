/**
 * Shared helpers for Microsoft Graph API calls.
 * Used by createUser, updateUser, and loadUsers.
 */

export const GRAPH_USERS_URL = 'https://graph.microsoft.com/v1.0/users';

/** Reads a cell value from a row as a trimmed string. */
export function cell(row: unknown[], index: number): string {
  const raw = row[index];
  return String(raw ?? '').trim();
}

/** Reads a cell value; returns null if empty (for Graph PATCH). */
export function cellOrNull(row: unknown[], index: number): string | null {
  const v = cell(row, index);
  return v === '' ? null : v;
}

/**
 * Parses a Graph API error response body into a human-readable message.
 * Tries JSON first, falls back to raw text.
 */
export async function parseGraphErrorResponse(
  response: Response
): Promise<string> {
  const text = await response.text();
  try {
    const errJson = JSON.parse(text) as {
      error?: { message?: string; code?: string };
    };
    return (errJson.error?.message ?? text) || `HTTP ${response.status}`;
  } catch {
    return text || `HTTP ${response.status}`;
  }
}

/** Converts an unknown error to a string message. */
export function toErrorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
