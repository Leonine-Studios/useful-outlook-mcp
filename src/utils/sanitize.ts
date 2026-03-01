/**
 * Input sanitization utilities for Graph API requests.
 * 
 * Prevents URL injection (path traversal, query parameter injection)
 * and OData filter injection from LLM-supplied input.
 * 
 * @see Security audit finding M1.2
 */

const SAFE_ID_PATTERN = /^[a-zA-Z0-9_=+.-]+$/;

/**
 * Validate that a value is safe for use as a URL path segment.
 * 
 * Graph API IDs are alphanumeric strings with `=`, `-`, `_`, `+`, `.` characters.
 * Rejects any input containing path traversal (`/`, `..`), query injection
 * (`?`, `#`, `&`), or other URL-semantic characters.
 * 
 * @throws Error if the value contains unsafe characters
 */
export function sanitizePathSegment(value: string, paramName = 'id'): string {
  if (!value || !SAFE_ID_PATTERN.test(value)) {
    throw new Error(
      `Invalid ${paramName}: contains characters not allowed in a resource identifier. ` +
      `Only alphanumeric characters, hyphens, underscores, dots, plus signs, and equals signs are permitted.`
    );
  }
  return value;
}

/**
 * Escape a string value for safe inclusion in an OData $filter expression.
 * 
 * OData string literals are delimited by single quotes. A single quote inside
 * the value must be escaped by doubling it (`'` -> `''`).
 */
export function sanitizeODataString(value: string): string {
  return value.replace(/'/g, "''");
}
