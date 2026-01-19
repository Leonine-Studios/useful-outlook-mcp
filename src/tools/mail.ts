/**
 * Mail tools for Microsoft Graph API
 */

import { z } from 'zod';
import { graphRequest, handleGraphResponse, formatErrorResponse } from '../graph/client.js';
import logger from '../utils/logger.js';
import { getContextUserId } from '../utils/context.js';

// ============================================================================
// Schemas
// ============================================================================

const listMailMessagesSchema = z.object({
  folderId: z.string().optional(),
  top: z.number().min(1).max(50).optional().default(10),
  skip: z.number().min(0).optional(),
  filter: z.string().optional(),
  search: z.string().optional(),
  orderBy: z.string().optional().default('receivedDateTime desc'),
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

// ============================================================================
// Tool Implementations
// ============================================================================

/**
 * List all mail folders
 */
async function listMailFolders() {
  logger.info('Tool: list-mail-folders', { user: getContextUserId() });
  
  try {
    const response = await graphRequest<{ value: unknown[] }>('/me/mailFolders');
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * List mail messages
 */
async function listMailMessages(params: Record<string, unknown>) {
  const parsed = listMailMessagesSchema.parse(params);
  const { folderId, top, skip, filter, search, orderBy } = parsed;
  
  logger.info('Tool: list-mail-messages', { 
    user: getContextUserId(),
    folderId,
    top,
  });
  
  try {
    const queryParams = new URLSearchParams();
    
    if (top) queryParams.set('$top', String(top));
    if (skip) queryParams.set('$skip', String(skip));
    if (filter) queryParams.set('$filter', filter);
    if (search) queryParams.set('$search', `"${search}"`);
    if (orderBy) queryParams.set('$orderby', orderBy);
    
    queryParams.set('$select', 'id,subject,from,toRecipients,receivedDateTime,isRead,importance,hasAttachments,bodyPreview');
    
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
 * Get a single mail message by ID
 */
async function getMailMessage(params: Record<string, unknown>) {
  const { messageId } = getMailMessageSchema.parse(params);
  
  logger.info('Tool: get-mail-message', { 
    user: getContextUserId(),
    messageId,
  });
  
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
  
  logger.info('Tool: send-mail', { 
    user: getContextUserId(),
    to: to.length,
    subject: subject.substring(0, 50),
  });
  
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
  
  logger.info('Tool: delete-mail-message', { 
    user: getContextUserId(),
    messageId,
  });
  
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
  
  logger.info('Tool: move-mail-message', { 
    user: getContextUserId(),
    messageId,
    destinationFolderId,
  });
  
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

// ============================================================================
// Tool Definitions for MCP
// ============================================================================

export const mailToolDefinitions = [
  {
    name: 'list-mail-folders',
    description: 'List all mail folders for the authenticated user',
    readOnly: true,
    inputSchema: {
      type: 'object' as const,
      properties: {},
    },
    handler: listMailFolders,
  },
  {
    name: 'list-mail-messages',
    description: 'List mail messages from a folder. Defaults to Inbox.',
    readOnly: true,
    inputSchema: {
      type: 'object' as const,
      properties: {
        folderId: {
          type: 'string',
          description: 'Mail folder ID (default: Inbox)',
        },
        top: {
          type: 'number',
          description: 'Maximum number of messages to return (1-50, default: 10)',
        },
        skip: {
          type: 'number',
          description: 'Number of messages to skip for pagination',
        },
        filter: {
          type: 'string',
          description: 'OData filter expression',
        },
        search: {
          type: 'string',
          description: 'Search query',
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
    name: 'get-mail-message',
    description: 'Get a single mail message by its ID',
    readOnly: true,
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
];
