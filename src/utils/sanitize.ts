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
 * (`?`, `#`, `&`), or other URL-semantic characters. Consecutive dots (`..`)
 * are explicitly rejected to prevent directory traversal even though single
 * dots are allowed in the character set.
 * 
 * @throws Error if the value contains unsafe characters or traversal sequences
 */
export function sanitizePathSegment(value: string, paramName = 'id'): string {
  if (!value || !SAFE_ID_PATTERN.test(value) || value.includes('..')) {
    throw new Error(
      `Invalid ${paramName}: contains characters not allowed in a resource identifier. ` +
      `Only alphanumeric characters, hyphens, underscores, dots, plus signs, and equals signs are permitted. ` +
      `Path traversal sequences ("..") are not allowed.`
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

const ISO_DATETIME_PATTERN = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?$/;

/**
 * Validate that a value is a well-formed ISO 8601 datetime for use in
 * unquoted OData $filter expressions (e.g., `receivedDateTime ge ...`).
 * 
 * Rejects values containing spaces, logical operators, or other characters
 * that could inject additional filter clauses.
 * 
 * @throws Error if the value is not a valid ISO 8601 datetime
 */
export function sanitizeODataDatetime(value: string, paramName = 'date'): string {
  if (!ISO_DATETIME_PATTERN.test(value)) {
    throw new Error(
      `Invalid ${paramName}: must be an ISO 8601 datetime (e.g., "2026-01-20T00:00:00Z").`
    );
  }
  return value;
}

const TIMEZONE_PATTERN = /^[a-zA-Z0-9_/+-]+$/;

/**
 * Validate that a value is a safe IANA timezone identifier for use in
 * HTTP header values (e.g., the Prefer header).
 * 
 * Rejects double quotes, spaces, and other characters that could alter
 * header structure.
 * 
 * @throws Error if the value contains unsafe characters
 */
export function sanitizeTimezone(value: string): string {
  if (!TIMEZONE_PATTERN.test(value)) {
    throw new Error(
      'Invalid timeZone: must be a valid IANA timezone identifier (e.g., "Europe/Berlin").'
    );
  }
  return value;
}
