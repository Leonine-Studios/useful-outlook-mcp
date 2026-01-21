/**
 * Mail tools for Microsoft Graph API
 */

import { z } from 'zod';
import { graphRequest, handleGraphResponse, formatErrorResponse } from '../graph/client.js';
import logger from '../utils/logger.js';

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
      ? `/me/mailFolders/${parentFolderId}/childFolders`
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
    filters.push(`startswith(from/emailAddress/address, '${params.senderEmail}')`);
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
    filters.push(`importance eq '${params.importance}'`);
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
      ? `/me/mailFolders/${folderId}/messages`
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
  
  // Add KQL property filters - formatKqlValue handles quoting and sanitization
  if (params.from) parts.push(`from:${formatKqlValue(params.from)}`);
  if (params.to) parts.push(`to:${formatKqlValue(params.to)}`);
  if (params.cc) parts.push(`cc:${formatKqlValue(params.cc)}`);
  if (params.bcc) parts.push(`bcc:${formatKqlValue(params.bcc)}`);
  if (params.participants) parts.push(`participants:${formatKqlValue(params.participants)}`);
  if (params.subject) parts.push(`subject:${formatKqlValue(params.subject)}`);
  if (params.body) parts.push(`body:${formatKqlValue(params.body)}`);
  if (params.attachment) parts.push(`attachment:${formatKqlValue(params.attachment)}`);
  if (params.hasAttachments !== undefined) parts.push(`hasAttachments:${params.hasAttachments}`);
  if (params.importance) parts.push(`importance:${params.importance}`);
  if (params.received) parts.push(`received:${params.received}`);
  
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
      
      if (from) filters.push(`startswith(from/emailAddress/address, '${from}')`);
      if (subject) filters.push(`contains(subject, '${subject}')`);
      if (hasAttachments !== undefined) filters.push(`hasAttachments eq ${hasAttachments}`);
      if (importance) filters.push(`importance eq '${importance}'`);
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
      
      const url = `/me/mailFolders/${folderId}/messages?${queryParams.toString()}`;
      
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
            content: [{ type: 'text' as const, text: JSON.stringify(parsed, null, 2) }],
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
 * Get a single mail message by ID
 */
async function getMailMessage(params: Record<string, unknown>) {
  const { messageId } = getMailMessageSchema.parse(params);
  
  try {
    const response = await graphRequest(`/me/messages/${messageId}`);
    return handleGraphResponse(response);
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
          text: JSON.stringify({ success: true, message: 'Email sent successfully' }),
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
    const response = await graphRequest(`/me/messages/${messageId}`, {
      method: 'DELETE',
    });
    
    if (response.status === 204 || response.ok) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: true, message: 'Message deleted' }),
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
    const response = await graphRequest(`/me/messages/${messageId}/move`, {
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
    const response = await graphRequest(`/me/messages/${messageId}/reply`, {
      method: 'POST',
      body: {
        comment,
      },
    });
    
    if (response.status === 202 || response.ok) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: true, message: 'Reply sent successfully' }),
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
    const response = await graphRequest(`/me/messages/${messageId}/replyAll`, {
      method: 'POST',
      body: {
        comment,
      },
    });
    
    if (response.status === 202 || response.ok) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: true, message: 'Reply-all sent successfully' }),
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
    
    const response = await graphRequest(`/me/messages/${messageId}/createReply`, {
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
    
    const response = await graphRequest(`/me/messages/${messageId}/createReplyAll`, {
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

Use this tool for:
- Browsing recent emails chronologically
- Filtering by sender email address (uses partial match, not exact)
- Filtering by date range (receivedAfter/receivedBefore)
- Filtering by read/unread status
- Filtering by importance or attachments

For full-text search in body/subject or searching by TO/CC/BCC recipients, use search-mail instead.

GRAPH API QUIRKS:
- senderEmail uses startswith() internally (partial match) because Graph API's exact match is unreliable
- Cannot filter by to/cc/bcc recipients with this tool - use search-mail instead
- Avoid running many calls in parallel - Graph may return MailboxConcurrency errors

Examples:
- Get unread emails: { "isRead": false }
- Get emails from specific sender: { "senderEmail": "john@company.com" }
- Get emails from last week: { "receivedAfter": "2026-01-13T00:00:00Z" }
- Get important emails with attachments: { "importance": "high", "hasAttachments": true }`,
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
    description: `Search mail messages using full-text search (KQL). Searches across body, subject, and attachments.

Use this tool for:
- Full-text search in email body or attachments
- Searching by TO/CC/BCC recipients (not possible with list-mail-messages)
- Keyword queries across multiple fields
- Finding emails by attachment filename
- Searching within a specific folder (with limitations - see folderId)

For structured filtering (sender, date ranges, read/unread), use list-mail-messages instead.

CRITICAL - FINDING EMAILS BY PERSON NAME:
- from/to/cc/bcc/participants parameters expect EMAIL ADDRESSES (names are unreliable)
- To find someone's email address when you only have their NAME: Use "query" parameter with the person's name
- Example: To find John Doe's email, use { "query": "John Doe", "top": 5 }, then extract email from results

GRAPH API QUIRKS:
- Cannot combine with sorting ($orderby) - results are ranked by relevance
- $search does not support $skip pagination
- When folderId is set, only limited filters work (from, subject, hasAttachments, importance, received)
- Search results may be limited to ~1000 items
- Avoid running many calls in parallel - Graph may return MailboxConcurrency errors

Examples:
- Search for topic: { "query": "brainstorming session" }
- Find person's email: { "query": "John Doe", "top": 5 }
- Search from specific sender: { "from": "john@company.com", "query": "meeting notes" }
- Search by recipient: { "to": "alice@company.com", "subject": "project" }
- Search with date: { "query": "quarterly report", "received": ">=2026-01-01" }
- Search by attachment: { "attachment": "budget.xlsx" }
- Search in folder: { "folderId": "xxx", "from": "alice" }`,
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
          description: 'Sender EMAIL ADDRESS (names are unreliable). Example: "john@company.com". If you only have a name, use "query" parameter instead.',
        },
        to: {
          type: 'string',
          description: 'Recipient EMAIL ADDRESS in TO field (names are unreliable). Example: "alice@company.com". If you only have a name, use "query" parameter instead. Not supported in folder-specific search.',
        },
        cc: {
          type: 'string',
          description: 'Recipient EMAIL ADDRESS in CC field (names are unreliable). If you only have a name, use "query" parameter instead. Not supported in folder-specific search.',
        },
        bcc: {
          type: 'string',
          description: 'Recipient EMAIL ADDRESS in BCC field (names are unreliable). If you only have a name, use "query" parameter instead. Not supported in folder-specific search.',
        },
        participants: {
          type: 'string',
          description: 'Any participant EMAIL ADDRESS (from/to/cc/bcc combined). Names are unreliable - use "query" parameter instead. Not supported in folder-specific search.',
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
    description: 'Get a single mail message by its ID',
    readOnly: true,
    requiredScopes: ['Mail.Read'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: {
          type: 'string',
          description: 'The ID of the message to retrieve',
        },
      },
      required: ['messageId'],
    },
    handler: getMailMessage,
  },
  {
    name: 'send-mail',
    description: 'Send an email message',
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
          description: 'Email body content',
        },
        bodyType: {
          type: 'string',
          enum: ['html', 'text'],
          description: 'Body content type (default: html)',
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
          description: 'Email body content',
        },
        bodyType: {
          type: 'string',
          enum: ['html', 'text'],
          description: 'Body content type (default: html)',
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
