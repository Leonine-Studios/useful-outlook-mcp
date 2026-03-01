/**
 * Mail tools for Microsoft Graph API
 */

import { z } from 'zod';
import { graphRequest, handleGraphResponse, formatErrorResponse } from '../graph/client.js';
import logger from '../utils/logger.js';
import { serializeResponse } from '../utils/tonl.js';
import { sanitizePathSegment, sanitizeODataString } from '../utils/sanitize.js';

// ============================================================================
// Schemas
// ============================================================================

const listMailMessagesSchema = z.object({
  folderId: z.string().optional(),
  top: z.number().min(1).max(50).optional().default(10),
  skip: z.number().min(0).optional(),
  // User-friendly filter parameters
  senderEmail: z.string().optional(),
  receivedAfter: z.string().optional(),
  receivedBefore: z.string().optional(),
  isRead: z.boolean().optional(),
  hasAttachments: z.boolean().optional(),
  importance: z.enum(['low', 'normal', 'high']).optional(),
  orderBy: z.string().optional().default('receivedDateTime desc'),
});

const searchMailSchema = z.object({
  query: z.string().optional(),
  from: z.string().optional(),
  to: z.string().optional(),
  cc: z.string().optional(),
  bcc: z.string().optional(),
  participants: z.string().optional(),
  subject: z.string().optional(),
  body: z.string().optional(),
  attachment: z.string().optional(),
  hasAttachments: z.boolean().optional(),
  importance: z.enum(['low', 'normal', 'high']).optional(),
  received: z.string().optional(),
  folderId: z.string().optional(),
  top: z.number().min(1).max(1000).optional().default(25),
});

const getMailMessageSchema = z.object({
  messageId: z.string(),
  includeConversationHistory: z.boolean().optional().default(false),
});

const sendMailSchema = z.object({
  to: z.array(z.string()).min(1),
  subject: z.string(),
  body: z.string(),
  bodyType: z.enum(['html', 'text']).optional().default('html'),
  cc: z.array(z.string()).optional(),
  bcc: z.array(z.string()).optional(),
  importance: z.enum(['low', 'normal', 'high']).optional().default('normal'),
  saveToSentItems: z.boolean().optional().default(true),
});

const deleteMailMessageSchema = z.object({
  messageId: z.string(),
});

const moveMailMessageSchema = z.object({
  messageId: z.string(),
  destinationFolderId: z.string(),
});

const createDraftMailSchema = z.object({
  to: z.array(z.string()).optional(),
  subject: z.string().optional(),
  body: z.string().optional(),
  bodyType: z.enum(['html', 'text']).optional().default('html'),
  cc: z.array(z.string()).optional(),
  bcc: z.array(z.string()).optional(),
  importance: z.enum(['low', 'normal', 'high']).optional().default('normal'),
});

const replyMailSchema = z.object({
  messageId: z.string(),
  comment: z.string(),
});

const createReplyDraftSchema = z.object({
  messageId: z.string(),
  comment: z.string().optional(),
});

// ============================================================================
// Tool Implementations
// ============================================================================

const listMailFoldersSchema = z.object({
  parentFolderId: z.string().optional(),
});

/**
 * List mail folders (top-level or subfolders of a specific folder)
 */
async function listMailFolders(params: Record<string, unknown>) {
  const { parentFolderId } = listMailFoldersSchema.parse(params);
  
  try {
    // If parentFolderId is provided, list child folders; otherwise list top-level folders
    const endpoint = parentFolderId 
      ? `/me/mailFolders/${sanitizePathSegment(parentFolderId, 'parentFolderId')}/childFolders`
      : '/me/mailFolders';
    
    const response = await graphRequest<{ value: unknown[] }>(endpoint);
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Build OData filter expression from user-friendly parameters
 * 
 * Note: Microsoft Graph API has known limitations:
 * - `eq` on from/emailAddress/address is unreliable, use `startswith()` instead
 * - Filtering on to/cc/bcc is NOT supported with $filter, use $search instead
 */
function buildMailFilter(params: {
  senderEmail?: string;
  receivedAfter?: string;
  receivedBefore?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  importance?: string;
}): string | undefined {
  const filters: string[] = [];
  
  if (params.senderEmail) {
    // Using startswith() instead of eq because Graph API's eq filter on 
    // from/emailAddress/address is unreliable and often returns 0 results
    // See: https://learn.microsoft.com/en-us/answers/questions/2153305/
    filters.push(`startswith(from/emailAddress/address, '${sanitizeODataString(params.senderEmail)}')`);
  }
  if (params.receivedAfter) {
    filters.push(`receivedDateTime ge ${params.receivedAfter}`);
  }
  if (params.receivedBefore) {
    filters.push(`receivedDateTime le ${params.receivedBefore}`);
  }
  if (params.isRead !== undefined) {
    filters.push(`isRead eq ${params.isRead}`);
  }
  if (params.hasAttachments !== undefined) {
    filters.push(`hasAttachments eq ${params.hasAttachments}`);
  }
  if (params.importance) {
    filters.push(`importance eq '${sanitizeODataString(params.importance)}'`);
  }
  
  return filters.length > 0 ? filters.join(' and ') : undefined;
}

/**
 * List mail messages with user-friendly filters
 */
async function listMailMessages(params: Record<string, unknown>) {
  const parsed = listMailMessagesSchema.parse(params);
  const { folderId, top, skip, senderEmail, receivedAfter, receivedBefore, isRead, hasAttachments, importance, orderBy } = parsed;
  
  try {
    const queryParams = new URLSearchParams();
    
    if (top) queryParams.set('$top', String(top));
    if (skip) queryParams.set('$skip', String(skip));
    
    // Build filter from user-friendly parameters
    const filter = buildMailFilter({ senderEmail, receivedAfter, receivedBefore, isRead, hasAttachments, importance });
    if (filter) queryParams.set('$filter', filter);
    
    // Note: Graph API has strict rules about combining $filter with $orderby
    // When using certain filters, orderBy causes "restriction or sort order is too complex" errors
    // Only apply orderBy when no complex filters are used
    const hasComplexFilter = senderEmail !== undefined || 
                             hasAttachments !== undefined || 
                             importance !== undefined;
    if (orderBy && !hasComplexFilter) {
      queryParams.set('$orderby', orderBy);
    } else if (!hasComplexFilter) {
      // Default sort when no complex filters
      queryParams.set('$orderby', 'receivedDateTime desc');
    }
    // When using complex filters (senderEmail, hasAttachments, importance), skip orderBy
    
    queryParams.set('$select', 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,importance,hasAttachments,bodyPreview');
    
    const endpoint = folderId 
      ? `/me/mailFolders/${sanitizePathSegment(folderId, 'folderId')}/messages`
      : '/me/messages';
    
    const url = `${endpoint}?${queryParams.toString()}`;
    
    const response = await graphRequest<{ value: unknown[] }>(url);
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Sanitize and format a KQL value for use inside $search="..."
 * 
 * - Strips unnecessary quotes that agents might add
 * - Email addresses with @ do NOT need quoting in KQL
 * - Values with spaces need escaped quotes (e.g., \"John Doe\")
 *   because they go inside the outer $search="..." wrapper
 */
function formatKqlValue(value: string): string {
  // First, strip any surrounding quotes that agents might have added
  let cleaned = value.trim();
  if ((cleaned.startsWith('"') && cleaned.endsWith('"')) ||
      (cleaned.startsWith("'") && cleaned.endsWith("'"))) {
    cleaned = cleaned.slice(1, -1);
  }
  
  // If value contains spaces, it needs to be quoted in KQL
  // Since the whole search is wrapped in $search="...", we need escaped quotes
  // Final URL: $search="subject:\"Project Name\""
  if (/\s/.test(cleaned)) {
    // Escape any existing quotes and wrap with escaped quotes
    const escaped = cleaned.replace(/"/g, '\\"');
    return `\\"${escaped}\\"`;
  }
  return cleaned;
}

/**
 * Format a free-text query parameter for use inside $search="..."
 * 
 * - If query looks like raw KQL (contains operators), pass through as-is
 * - If query is a multi-word phrase, quote it
 * - Single terms (even with numbers or special chars) pass through unquoted
 */
function formatQueryValue(query: string): string {
  const cleaned = query.trim();
  
  // Check if this looks like raw KQL (contains operators or property searches)
  const hasKqlOperators = /\b(AND|OR|NOT)\b|:/i.test(cleaned);
  
  if (hasKqlOperators) {
    // Raw KQL - pass through as-is
    return cleaned;
  }
  
  // Only quote multi-word phrases (those with spaces)
  // Single terms like "project", "deadline" work better unquoted
  if (/\s/.test(cleaned)) {
    // Multi-word phrase - escape existing quotes and wrap with escaped quotes
    const escaped = cleaned.replace(/"/g, '\\"');
    return `\\"${escaped}\\"`;
  }
  
  // Single term - pass through unquoted
  return cleaned;
}

/**
 * Build KQL query for a field that may contain multiple email addresses
 * Supports comma or semicolon-separated lists
 * Example: "email1@x.com, email2@x.com" -> "field:email1@x.com AND field:email2@x.com"
 */
function buildEmailFieldQuery(field: string, value: string): string {
  // Split by comma or semicolon, trim whitespace
  const emails = value.split(/[,;]/).map(e => e.trim()).filter(e => e.length > 0);
  
  if (emails.length === 0) return '';
  if (emails.length === 1) return `${field}:${formatKqlValue(emails[0])}`;
  
  // Multiple emails: field:email1 AND field:email2 AND field:email3
  return emails.map(email => `${field}:${formatKqlValue(email)}`).join(' AND ');
}

/**
 * Build KQL search query from parameters
 * 
 * Note: KQL (Keyword Query Language) requires quoting values with special characters
 * like email addresses (containing @) or phrases with spaces.
 */
function buildMailSearchQuery(params: {
  query?: string;
  from?: string;
  to?: string;
  cc?: string;
  bcc?: string;
  participants?: string;
  subject?: string;
  body?: string;
  attachment?: string;
  hasAttachments?: boolean;
  importance?: string;
  received?: string;
}): string {
  const parts: string[] = [];
  
  // Add free-text query - formatQueryValue handles KQL operators vs simple phrases
  if (params.query) {
    parts.push(formatQueryValue(params.query));
  }
  
  // Add KQL property filters for email fields (support multiple emails with AND)
  if (params.from) parts.push(buildEmailFieldQuery('from', params.from));
  if (params.to) parts.push(buildEmailFieldQuery('to', params.to));
  if (params.cc) parts.push(buildEmailFieldQuery('cc', params.cc));
  if (params.bcc) parts.push(buildEmailFieldQuery('bcc', params.bcc));
  if (params.participants) parts.push(buildEmailFieldQuery('participants', params.participants));
  if (params.subject) parts.push(`subject:${formatKqlValue(params.subject)}`);
  if (params.body) parts.push(`body:${formatKqlValue(params.body)}`);
  if (params.attachment) parts.push(`attachment:${formatKqlValue(params.attachment)}`);
  if (params.hasAttachments !== undefined) parts.push(`hasAttachments:${params.hasAttachments}`);
  if (params.importance) parts.push(`importance:${params.importance}`);
  if (params.received) {
    const val = params.received.trim();
    // KQL date syntax: use colon for ranges/keywords (received:2023-01-01..2023-12-31, received:today)
    // but NO colon for comparison operators (received>=2023-01-01)
    if (/^[<>=]/.test(val)) {
      parts.push(`received${val}`);
    } else {
      parts.push(`received:${val}`);
    }
  }
  
  // KQL joining logic:
  // - Free-text query + property filters need explicit AND
  // - Multiple email filters generate internal ANDs, so join all with AND for consistency
  // - Check if any part contains AND (means multiple emails were specified)
  const hasComplexFilters = parts.some(part => part.includes(' AND '));
  
  if (params.query && parts.length > 1) {
    // Free-text query with property filters: join with AND
    return parts.join(' AND ');
  } else if (hasComplexFilters) {
    // Multiple emails specified: join all with AND for proper precedence
    return parts.join(' AND ');
  }
  
  // Simple property filters can be space-separated
  return parts.join(' ');
}

/**
 * Search mail messages using KQL $search
 */
async function searchMail(params: Record<string, unknown>) {
  const parsed = searchMailSchema.parse(params);
  const { query, from, to, cc, bcc, participants, subject, body, attachment, hasAttachments, importance, received, folderId, top } = parsed;
  
  try {
    // If folderId is specified, $search is not supported on folder endpoints
    // Use $filter for simple cases, or fetch and filter client-side
    if (folderId) {
      // Build OData filter for folder-specific search (limited capabilities)
      // Note: Using startswith() instead of eq/contains because Graph API is unreliable with exact matches
      const filters: string[] = [];
      
      if (from) filters.push(`startswith(from/emailAddress/address, '${sanitizeODataString(from)}')`);
      if (subject) filters.push(`contains(subject, '${sanitizeODataString(subject)}')`);
      if (hasAttachments !== undefined) filters.push(`hasAttachments eq ${hasAttachments}`);
      if (importance) filters.push(`importance eq '${sanitizeODataString(importance)}'`);
      if (received) {
        // Parse received format like ">=2026-01-06" 
        const match = received.match(/^(>=|<=|>|<|=)?(.+)$/);
        if (match) {
          const [, op, date] = match;
          const operator = op === '>=' ? 'ge' : op === '<=' ? 'le' : op === '>' ? 'gt' : op === '<' ? 'lt' : 'eq';
          filters.push(`receivedDateTime ${operator} ${date}`);
        }
      }
      
      const queryParams = new URLSearchParams();
      if (filters.length > 0) queryParams.set('$filter', filters.join(' and '));
      if (top) queryParams.set('$top', String(top));
      queryParams.set('$select', 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,importance,hasAttachments,bodyPreview');
      
      const url = `/me/mailFolders/${sanitizePathSegment(folderId, 'folderId')}/messages?${queryParams.toString()}`;
      
      const response = await graphRequest<{ value: unknown[] }>(url);
      
      // Note: full-text search (body, attachment content, etc.) is not supported on folder endpoints
      // If those params were specified, add a warning
      const limitedParams = [query, body, attachment, to, cc, bcc, participants].filter(p => p !== undefined);
      if (limitedParams.length > 0) {
        const result = handleGraphResponse(response);
        if (result.content && result.content[0]?.type === 'text') {
          const parsed = JSON.parse(result.content[0].text);
          parsed._warning = 'Full-text search (body, attachment, to, cc, bcc, participants) is not supported when searching within a specific folder. Only from, subject, hasAttachments, importance, and received filters were applied.';
          return {
            content: [{ type: 'text' as const, text: serializeResponse(parsed) }],
          };
        }
      }
      
      return handleGraphResponse(response);
    }
    
    // Mailbox-wide search using KQL $search
    const searchQuery = buildMailSearchQuery({ query, from, to, cc, bcc, participants, subject, body, attachment, hasAttachments, importance, received });
    
    if (!searchQuery) {
      return formatErrorResponse(new Error('At least one search parameter is required'));
    }
    
    logger.debug('search-mail: KQL query', { searchQuery });
    
    const queryParams = new URLSearchParams();
    queryParams.set('$search', `"${searchQuery}"`);
    if (top) queryParams.set('$top', String(top));
    queryParams.set('$select', 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,importance,hasAttachments,bodyPreview');
    
    const url = `/me/messages?${queryParams.toString()}`;
    
    // Note: Some search queries may require ConsistencyLevel header
    // See: https://learn.microsoft.com/en-us/graph/aad-advanced-queries
    const response = await graphRequest<{ value: unknown[] }>(url, {
      headers: {
        'ConsistencyLevel': 'eventual',
      },
    });
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Strip HTML tags and decode entities for plain text extraction
 */
function stripHtml(html: string): string {
  return html
    // Remove style and script tags with content
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    // Replace common block elements with newlines
    .replace(/<\/?(p|div|br|hr|tr|li|h[1-6])[^>]*>/gi, '\n')
    // Remove all remaining HTML tags
    .replace(/<[^>]+>/g, '')
    // Decode common HTML entities
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&rsquo;/gi, "'")
    .replace(/&lsquo;/gi, "'")
    .replace(/&rdquo;/gi, '"')
    .replace(/&ldquo;/gi, '"')
    .replace(/&bull;/gi, '•')
    .replace(/&mdash;/gi, '—')
    .replace(/&ndash;/gi, '–')
    // Collapse multiple newlines
    .replace(/\n{3,}/g, '\n\n')
    // Trim whitespace from each line
    .split('\n')
    .map(line => line.trim())
    .join('\n')
    .trim();
}

/**
 * Get a single mail message by ID
 * Always returns body as plain text to minimize context window usage
 */
async function getMailMessage(params: Record<string, unknown>) {
  const { messageId, includeConversationHistory } = getMailMessageSchema.parse(params);
  
  try {
    // Use uniqueBody (excludes conversation history) by default
    const bodyField = includeConversationHistory ? 'body' : 'uniqueBody';
    const selectFields = `id,subject,from,sender,toRecipients,ccRecipients,bccRecipients,replyTo,receivedDateTime,sentDateTime,isRead,isDraft,importance,hasAttachments,internetMessageId,conversationId,webLink,flag,${bodyField}`;
    
    const url = `/me/messages/${sanitizePathSegment(messageId, 'messageId')}?$select=${encodeURIComponent(selectFields)}`;
    
    // Request plain text body from Graph API
    const response = await graphRequest(url, {
      headers: {
        'Prefer': 'outlook.body-content-type="text"',
      },
    });
    const result = handleGraphResponse(response);
    
    // If API returned HTML anyway (can happen), strip HTML client-side
    if (result.content?.[0]?.type === 'text') {
      try {
        const data = JSON.parse(result.content[0].text);
        
        // Handle both body and uniqueBody fields
        const bodyData = data.body || data.uniqueBody;
        if (bodyData?.contentType?.toLowerCase() === 'html' && bodyData?.content) {
          bodyData.content = stripHtml(bodyData.content);
          bodyData.contentType = 'text';
          
          // Normalize field name to 'body' for consistency
          if (data.uniqueBody) {
            data.body = bodyData;
            delete data.uniqueBody;
          }
          
          return {
            content: [{ type: 'text' as const, text: serializeResponse(data) }],
          };
        }
        
        // Normalize uniqueBody to body if present
        if (data.uniqueBody && !data.body) {
          data.body = data.uniqueBody;
          delete data.uniqueBody;
          return {
            content: [{ type: 'text' as const, text: serializeResponse(data) }],
          };
        }
      } catch {
        // If parsing fails, return original response
      }
    }
    
    return result;
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Send an email
 */
async function sendMail(params: Record<string, unknown>) {
  const { to, subject, body, bodyType, cc, bcc, importance, saveToSentItems } = sendMailSchema.parse(params);
  
  try {
    const message: Record<string, unknown> = {
      subject,
      body: {
        contentType: bodyType === 'html' ? 'HTML' : 'Text',
        content: body,
      },
      toRecipients: to.map(email => ({
        emailAddress: { address: email },
      })),
      importance,
    };
    
    if (cc?.length) {
      message.ccRecipients = cc.map(email => ({
        emailAddress: { address: email },
      }));
    }
    
    if (bcc?.length) {
      message.bccRecipients = bcc.map(email => ({
        emailAddress: { address: email },
      }));
    }
    
    const response = await graphRequest('/me/sendMail', {
      method: 'POST',
      body: {
        message,
        saveToSentItems,
      },
    });
    
    if (response.status === 202 || response.ok) {
      return {
        content: [{
          type: 'text' as const,
          text: serializeResponse({ success: true, message: 'Email sent successfully' }),
        }],
      };
    }
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Delete a mail message
 */
async function deleteMailMessage(params: Record<string, unknown>) {
  const { messageId } = deleteMailMessageSchema.parse(params);
  
  try {
    const response = await graphRequest(`/me/messages/${sanitizePathSegment(messageId, 'messageId')}`, {
      method: 'DELETE',
    });
    
    if (response.status === 204 || response.ok) {
      return {
        content: [{
          type: 'text' as const,
          text: serializeResponse({ success: true, message: 'Message deleted' }),
        }],
      };
    }
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Move a mail message to a different folder
 */
async function moveMailMessage(params: Record<string, unknown>) {
  const { messageId, destinationFolderId } = moveMailMessageSchema.parse(params);
  
  try {
    const response = await graphRequest(`/me/messages/${sanitizePathSegment(messageId, 'messageId')}/move`, {
      method: 'POST',
      body: {
        destinationId: destinationFolderId,
      },
    });
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Create an email draft (saves to Drafts folder)
 */
async function createDraftMail(params: Record<string, unknown>) {
  const { to, subject, body, bodyType, cc, bcc, importance } = createDraftMailSchema.parse(params);
  
  try {
    const message: Record<string, unknown> = {
      importance,
    };
    
    if (subject) {
      message.subject = subject;
    }
    
    if (body) {
      message.body = {
        contentType: bodyType === 'html' ? 'HTML' : 'Text',
        content: body,
      };
    }
    
    if (to?.length) {
      message.toRecipients = to.map(email => ({
        emailAddress: { address: email },
      }));
    }
    
    if (cc?.length) {
      message.ccRecipients = cc.map(email => ({
        emailAddress: { address: email },
      }));
    }
    
    if (bcc?.length) {
      message.bccRecipients = bcc.map(email => ({
        emailAddress: { address: email },
      }));
    }
    
    const response = await graphRequest('/me/messages', {
      method: 'POST',
      body: message,
    });
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Reply to a mail message (sends immediately)
 */
async function replyMail(params: Record<string, unknown>) {
  const { messageId, comment } = replyMailSchema.parse(params);
  
  try {
    const response = await graphRequest(`/me/messages/${sanitizePathSegment(messageId, 'messageId')}/reply`, {
      method: 'POST',
      body: {
        comment,
      },
    });
    
    if (response.status === 202 || response.ok) {
      return {
        content: [{
          type: 'text' as const,
          text: serializeResponse({ success: true, message: 'Reply sent successfully' }),
        }],
      };
    }
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Reply all to a mail message (sends immediately)
 */
async function replyAllMail(params: Record<string, unknown>) {
  const { messageId, comment } = replyMailSchema.parse(params);
  
  try {
    const response = await graphRequest(`/me/messages/${sanitizePathSegment(messageId, 'messageId')}/replyAll`, {
      method: 'POST',
      body: {
        comment,
      },
    });
    
    if (response.status === 202 || response.ok) {
      return {
        content: [{
          type: 'text' as const,
          text: serializeResponse({ success: true, message: 'Reply-all sent successfully' }),
        }],
      };
    }
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Create a reply draft (saves to Drafts folder)
 */
async function createReplyDraft(params: Record<string, unknown>) {
  const { messageId, comment } = createReplyDraftSchema.parse(params);
  
  try {
    const body: Record<string, unknown> = {};
    if (comment) {
      body.comment = comment;
    }
    
    const response = await graphRequest(`/me/messages/${sanitizePathSegment(messageId, 'messageId')}/createReply`, {
      method: 'POST',
      body,
    });
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Create a reply-all draft (saves to Drafts folder)
 */
async function createReplyAllDraft(params: Record<string, unknown>) {
  const { messageId, comment } = createReplyDraftSchema.parse(params);
  
  try {
    const body: Record<string, unknown> = {};
    if (comment) {
      body.comment = comment;
    }
    
    const response = await graphRequest(`/me/messages/${sanitizePathSegment(messageId, 'messageId')}/createReplyAll`, {
      method: 'POST',
      body,
    });
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

// ============================================================================
// Tool Definitions for MCP
// ============================================================================

export const mailToolDefinitions = [
  {
    name: 'list-mail-folders',
    description: `List mail folders. Can list top-level folders or subfolders of a specific folder.

Use this to:
- Get all top-level folders (Inbox, Sent Items, etc.)
- Find subfolders within a folder (e.g., "Done" under Inbox)

Examples:
- List top-level folders: {} (no parameters)
- List subfolders of Inbox: { "parentFolderId": "<inbox-folder-id>" }

To find a subfolder like "Done" under Inbox:
1. First call with no params to get Inbox folder ID
2. Then call with parentFolderId set to the Inbox ID`,
    readOnly: true,
    requiredScopes: ['Mail.Read'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        parentFolderId: {
          type: 'string',
          description: 'Parent folder ID to list subfolders. If not provided, lists top-level folders.',
        },
      },
    },
    handler: listMailFolders,
  },
  {
    name: 'list-mail-messages',
    description: `List and filter mail messages from a folder with structured filters. Defaults to Inbox.

Use this for: chronological browsing, filtering by sender/date/read status/importance/attachments.
For full-text search in body/subject or filtering by TO/CC/BCC recipients, use search-mail instead.

Filtering notes:
- senderEmail uses partial match (not exact)
- Cannot filter by to/cc/bcc recipients - use search-mail for that

Examples:
- Unread emails: { "isRead": false }
- From sender: { "senderEmail": "john@company.com" }
- Last week: { "receivedAfter": "2026-01-13T00:00:00Z" }
- Important with attachments: { "importance": "high", "hasAttachments": true }`,
    readOnly: true,
    requiredScopes: ['Mail.Read'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        folderId: {
          type: 'string',
          description: 'Mail folder ID (default: Inbox). Use list-mail-folders to get folder IDs.',
        },
        senderEmail: {
          type: 'string',
          description: 'Filter by sender email (partial match using startswith). Example: "john.doe@company.com"',
        },
        receivedAfter: {
          type: 'string',
          description: 'Only messages received after this date/time (ISO 8601). Example: "2026-01-13T00:00:00Z"',
        },
        receivedBefore: {
          type: 'string',
          description: 'Only messages received before this date/time (ISO 8601). Example: "2026-01-20T23:59:59Z"',
        },
        isRead: {
          type: 'boolean',
          description: 'Filter by read status. true = read messages, false = unread messages',
        },
        hasAttachments: {
          type: 'boolean',
          description: 'Filter by attachment presence. true = only messages with attachments',
        },
        importance: {
          type: 'string',
          enum: ['low', 'normal', 'high'],
          description: 'Filter by importance level',
        },
        top: {
          type: 'number',
          description: 'Maximum number of messages to return (1-50, default: 10)',
        },
        skip: {
          type: 'number',
          description: 'Number of messages to skip for pagination',
        },
        orderBy: {
          type: 'string',
          description: 'Sort order (default: receivedDateTime desc)',
        },
      },
    },
    handler: listMailMessages,
  },
  {
    name: 'search-mail',
    description: `Search mail using full-text search. Searches body, subject, attachments. Use this for text search and filtering by TO/CC/BCC. For simple filters (sender, date, read/unread), use list-mail-messages instead.

EMAIL ADDRESSES REQUIRED:
- from/to/cc/bcc/participants expect EMAIL ADDRESSES (names don't work reliably)
- If you only have a name: First call with {"query": "John Doe", "top": 5} to find their email, extract it from results, then search again
- Multiple comma-separated emails require ALL to match: {"participants": "alice@x.com, bob@x.com"}

Folder search limitations:
- When folderId is set, only from/subject/hasAttachments/importance/received work (no query/body/attachment/to/cc/bcc/participants)

Examples:
- Find person's email: {"query": "John Doe", "top": 5}
- From sender: {"from": "john@company.com", "query": "project"}
- By recipients: {"to": "alice@company.com", "subject": "budget"}
- Between two people: {"query": "project", "participants": "alice@x.com, bob@x.com"}
- By attachment: {"attachment": "report.pdf"}
- Folder search: {"folderId": "xxx", "from": "alice@company.com"}`,
    readOnly: true,
    requiredScopes: ['Mail.Read'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        query: {
          type: 'string',
          description: 'Free-text search query (searches body, subject, attachments, sender/recipient names). Use this to find emails by PERSON NAME. Example: "John Doe" to find any emails involving that person. Not supported in folder-specific search.',
        },
        from: {
          type: 'string',
          description: 'Sender EMAIL ADDRESS (names are unreliable). Example: "john@company.com". Supports multiple emails separated by commas to require ALL: "email1@x.com, email2@x.com". If you only have a name, use "query" parameter instead.',
        },
        to: {
          type: 'string',
          description: 'Recipient EMAIL ADDRESS in TO field (names are unreliable). Example: "alice@company.com". Supports multiple emails separated by commas to require ALL. If you only have a name, use "query" parameter instead. Not supported in folder-specific search.',
        },
        cc: {
          type: 'string',
          description: 'Recipient EMAIL ADDRESS in CC field (names are unreliable). Supports multiple emails separated by commas to require ALL. If you only have a name, use "query" parameter instead. Not supported in folder-specific search.',
        },
        bcc: {
          type: 'string',
          description: 'Recipient EMAIL ADDRESS in BCC field (names are unreliable). Supports multiple emails separated by commas to require ALL. If you only have a name, use "query" parameter instead. Not supported in folder-specific search.',
        },
        participants: {
          type: 'string',
          description: 'Any participant EMAIL ADDRESS (from/to/cc/bcc combined). Example: "john@x.com, alice@x.com" requires BOTH in the conversation. Supports multiple comma-separated emails to filter at API level. Names are unreliable - use "query" parameter instead. Not supported in folder-specific search.',
        },
        subject: {
          type: 'string',
          description: 'Subject keyword',
        },
        body: {
          type: 'string',
          description: 'Body content keyword. Not supported in folder-specific search.',
        },
        attachment: {
          type: 'string',
          description: 'Attachment filename. Example: "report.pdf". Not supported in folder-specific search.',
        },
        hasAttachments: {
          type: 'boolean',
          description: 'Has attachments',
        },
        importance: {
          type: 'string',
          enum: ['low', 'normal', 'high'],
          description: 'Filter by importance',
        },
        received: {
          type: 'string',
          description: 'Date filter. Examples: ">=2026-01-13", "2026-01-13..2026-01-20"',
        },
        folderId: {
          type: 'string',
          description: 'Folder ID to search within. When specified, only from, subject, hasAttachments, importance, and received filters are supported (full-text search not available on folder endpoints).',
        },
        top: {
          type: 'number',
          description: 'Maximum number of results (default: 25, max: 1000)',
        },
      },
    },
    handler: searchMail,
  },
  {
    name: 'get-mail-message',
    description: `Get full mail message content by ID. Returns body as plain text (HTML stripped) to minimize context window usage.

By default, returns only the unique message body (excludes forwarded/replied conversation history) using Graph API's uniqueBody field. Set includeConversationHistory to true to get the full email thread.

Use list-mail-messages or search-mail first to find message IDs and preview content, then use this tool only when you need the full message body.`,
    readOnly: true,
    requiredScopes: ['Mail.Read'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: {
          type: 'string',
          description: 'The ID of the message to retrieve',
        },
        includeConversationHistory: {
          type: 'boolean',
          description: 'Include full email thread with forwarded/replied messages (default: false). By default, only the unique message content is returned.',
        },
      },
      required: ['messageId'],
    },
    handler: getMailMessage,
  },
  {
    name: 'send-mail',
    description: 'Send an email message immediately.',
    readOnly: false,
    requiredScopes: ['Mail.Send'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        to: {
          type: 'array',
          items: { type: 'string' },
          description: 'Recipient email addresses',
        },
        subject: {
          type: 'string',
          description: 'Email subject',
        },
        body: {
          type: 'string',
          description: `Email body content in HTML (default) or plain text.

HTML formatting (bodyType="html"):
- Use semantic HTML: <p> for paragraphs, <ul>/<li> for lists, <strong> for emphasis
- IMPORTANT: \\n does NOT work in HTML - use <p> or <br> tags instead
- Use bodyType="text" for plain text where \\n line breaks work automatically`,
        },
        bodyType: {
          type: 'string',
          enum: ['html', 'text'],
          description: 'Body content type. Use "html" (default) for formatted emails with structure. Use "text" for plain text where \\n line breaks work automatically.',
        },
        cc: {
          type: 'array',
          items: { type: 'string' },
          description: 'CC recipients',
        },
        bcc: {
          type: 'array',
          items: { type: 'string' },
          description: 'BCC recipients',
        },
        importance: {
          type: 'string',
          enum: ['low', 'normal', 'high'],
          description: 'Email importance (default: normal)',
        },
        saveToSentItems: {
          type: 'boolean',
          description: 'Save to Sent Items folder (default: true)',
        },
      },
      required: ['to', 'subject', 'body'],
    },
    handler: sendMail,
  },
  {
    name: 'delete-mail-message',
    description: 'Delete a mail message (moves to Deleted Items)',
    readOnly: false,
    requiredScopes: ['Mail.ReadWrite'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: {
          type: 'string',
          description: 'The ID of the message to delete',
        },
      },
      required: ['messageId'],
    },
    handler: deleteMailMessage,
  },
  {
    name: 'move-mail-message',
    description: 'Move a mail message to a different folder',
    readOnly: false,
    requiredScopes: ['Mail.ReadWrite'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: {
          type: 'string',
          description: 'The ID of the message to move',
        },
        destinationFolderId: {
          type: 'string',
          description: 'The ID of the destination folder',
        },
      },
      required: ['messageId', 'destinationFolderId'],
    },
    handler: moveMailMessage,
  },
  {
    name: 'create-draft-mail',
    description: 'Create an email draft and save it to the Drafts folder. Returns the draft message ID which can be used to send or update it later.',
    readOnly: false,
    requiredScopes: ['Mail.ReadWrite'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        to: {
          type: 'array',
          items: { type: 'string' },
          description: 'Recipient email addresses (optional for drafts)',
        },
        subject: {
          type: 'string',
          description: 'Email subject',
        },
        body: {
          type: 'string',
          description: `Email body content in HTML (default) or plain text.

HTML formatting (bodyType="html"):
- Use semantic HTML: <p> for paragraphs, <ul>/<li> for lists, <strong> for emphasis
- IMPORTANT: \\n does NOT work in HTML - use <p> or <br> tags instead
- Use bodyType="text" for plain text where \\n line breaks work automatically`,
        },
        bodyType: {
          type: 'string',
          enum: ['html', 'text'],
          description: 'Body content type. Use "html" (default) for formatted emails with structure. Use "text" for plain text where \\n line breaks work automatically.',
        },
        cc: {
          type: 'array',
          items: { type: 'string' },
          description: 'CC recipients',
        },
        bcc: {
          type: 'array',
          items: { type: 'string' },
          description: 'BCC recipients',
        },
        importance: {
          type: 'string',
          enum: ['low', 'normal', 'high'],
          description: 'Email importance (default: normal)',
        },
      },
    },
    handler: createDraftMail,
  },
  {
    name: 'reply-mail',
    description: 'Reply to a mail message. Sends the reply immediately to the original sender.',
    readOnly: false,
    requiredScopes: ['Mail.Send'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: {
          type: 'string',
          description: 'The ID of the message to reply to',
        },
        comment: {
          type: 'string',
          description: 'The reply body content',
        },
      },
      required: ['messageId', 'comment'],
    },
    handler: replyMail,
  },
  {
    name: 'reply-all-mail',
    description: 'Reply to all recipients of a mail message. Sends the reply immediately to all original recipients.',
    readOnly: false,
    requiredScopes: ['Mail.Send'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: {
          type: 'string',
          description: 'The ID of the message to reply to',
        },
        comment: {
          type: 'string',
          description: 'The reply body content',
        },
      },
      required: ['messageId', 'comment'],
    },
    handler: replyAllMail,
  },
  {
    name: 'create-reply-draft',
    description: 'Create a reply draft to a mail message. Saves the draft to the Drafts folder for review before sending.',
    readOnly: false,
    requiredScopes: ['Mail.ReadWrite'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: {
          type: 'string',
          description: 'The ID of the message to reply to',
        },
        comment: {
          type: 'string',
          description: 'The reply body content (optional)',
        },
      },
      required: ['messageId'],
    },
    handler: createReplyDraft,
  },
  {
    name: 'create-reply-all-draft',
    description: 'Create a reply-all draft to a mail message. Saves the draft to the Drafts folder for review before sending.',
    readOnly: false,
    requiredScopes: ['Mail.ReadWrite'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: {
          type: 'string',
          description: 'The ID of the message to reply to',
        },
        comment: {
          type: 'string',
          description: 'The reply body content (optional)',
        },
      },
      required: ['messageId'],
    },
    handler: createReplyAllDraft,
  },
];
