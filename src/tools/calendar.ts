/**
 * Calendar tools for Microsoft Graph API
 */

import { z } from 'zod';
import { graphRequest, handleGraphResponse, formatErrorResponse } from '../graph/client.js';
import logger from '../utils/logger.js';
import { getContextUserId } from '../utils/context.js';

// ============================================================================
// Schemas
// ============================================================================

const listCalendarEventsSchema = z.object({
  calendarId: z.string().optional(),
  top: z.number().min(1).max(50).optional().default(10),
  skip: z.number().min(0).optional(),
  filter: z.string().optional(),
  orderBy: z.string().optional().default('start/dateTime'),
});

const getCalendarEventSchema = z.object({
  eventId: z.string(),
});

const getCalendarViewSchema = z.object({
  startDateTime: z.string(),
  endDateTime: z.string(),
  calendarId: z.string().optional(),
  top: z.number().min(1).max(50).optional().default(10),
});

const createCalendarEventSchema = z.object({
  subject: z.string(),
  start: z.string(),
  end: z.string(),
  timeZone: z.string().optional().default('UTC'),
  body: z.string().optional(),
  bodyType: z.enum(['html', 'text']).optional().default('text'),
  location: z.string().optional(),
  attendees: z.array(z.object({
    email: z.string(),
    type: z.enum(['required', 'optional']).optional().default('required'),
  })).optional(),
  isAllDay: z.boolean().optional().default(false),
  reminderMinutesBeforeStart: z.number().optional(),
  calendarId: z.string().optional(),
});

const updateCalendarEventSchema = z.object({
  eventId: z.string(),
  subject: z.string().optional(),
  start: z.string().optional(),
  end: z.string().optional(),
  timeZone: z.string().optional(),
  body: z.string().optional(),
  location: z.string().optional(),
  attendees: z.array(z.object({
    email: z.string(),
    type: z.enum(['required', 'optional']).optional(),
  })).optional(),
});

const deleteCalendarEventSchema = z.object({
  eventId: z.string(),
});

// ============================================================================
// Tool Implementations
// ============================================================================

/**
 * List all calendars
 */
async function listCalendars() {
  logger.info('Tool: list-calendars', { user: getContextUserId() });
  
  try {
    const response = await graphRequest<{ value: unknown[] }>('/me/calendars');
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * List calendar events
 */
async function listCalendarEvents(params: Record<string, unknown>) {
  const { calendarId, top, skip, filter, orderBy } = listCalendarEventsSchema.parse(params);
  
  logger.info('Tool: list-calendar-events', { 
    user: getContextUserId(),
    calendarId,
    top,
  });
  
  try {
    const queryParams = new URLSearchParams();
    
    if (top) queryParams.set('$top', String(top));
    if (skip) queryParams.set('$skip', String(skip));
    if (filter) queryParams.set('$filter', filter);
    if (orderBy) queryParams.set('$orderby', orderBy);
    
    queryParams.set('$select', 'id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,bodyPreview');
    
    const endpoint = calendarId 
      ? `/me/calendars/${calendarId}/events`
      : '/me/events';
    
    const url = `${endpoint}?${queryParams.toString()}`;
    
    const response = await graphRequest<{ value: unknown[] }>(url);
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Get a single calendar event by ID
 */
async function getCalendarEvent(params: Record<string, unknown>) {
  const { eventId } = getCalendarEventSchema.parse(params);
  
  logger.info('Tool: get-calendar-event', { 
    user: getContextUserId(),
    eventId,
  });
  
  try {
    const response = await graphRequest(`/me/events/${eventId}`);
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Get calendar view for a time range
 */
async function getCalendarView(params: Record<string, unknown>) {
  const { startDateTime, endDateTime, calendarId, top } = getCalendarViewSchema.parse(params);
  
  logger.info('Tool: get-calendar-view', { 
    user: getContextUserId(),
    startDateTime,
    endDateTime,
  });
  
  try {
    const queryParams = new URLSearchParams();
    queryParams.set('startDateTime', startDateTime);
    queryParams.set('endDateTime', endDateTime);
    if (top) queryParams.set('$top', String(top));
    queryParams.set('$select', 'id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,bodyPreview');
    
    const endpoint = calendarId 
      ? `/me/calendars/${calendarId}/calendarView`
      : '/me/calendarView';
    
    const url = `${endpoint}?${queryParams.toString()}`;
    
    const response = await graphRequest<{ value: unknown[] }>(url);
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Create a calendar event
 */
async function createCalendarEvent(params: Record<string, unknown>) {
  const { 
    subject, start, end, timeZone, body, bodyType, 
    location, attendees, isAllDay, reminderMinutesBeforeStart, calendarId 
  } = createCalendarEventSchema.parse(params);
  
  logger.info('Tool: create-calendar-event', { 
    user: getContextUserId(),
    subject,
    start,
    end,
  });
  
  try {
    const event: Record<string, unknown> = {
      subject,
      start: {
        dateTime: start,
        timeZone: timeZone || 'UTC',
      },
      end: {
        dateTime: end,
        timeZone: timeZone || 'UTC',
      },
      isAllDay,
    };
    
    if (body) {
      event.body = {
        contentType: bodyType === 'html' ? 'HTML' : 'Text',
        content: body,
      };
    }
    
    if (location) {
      event.location = { displayName: location };
    }
    
    if (attendees?.length) {
      event.attendees = attendees.map(a => ({
        emailAddress: { address: a.email },
        type: a.type || 'required',
      }));
    }
    
    if (reminderMinutesBeforeStart !== undefined) {
      event.reminderMinutesBeforeStart = reminderMinutesBeforeStart;
      event.isReminderOn = true;
    }
    
    const endpoint = calendarId 
      ? `/me/calendars/${calendarId}/events`
      : '/me/events';
    
    const response = await graphRequest(endpoint, {
      method: 'POST',
      body: event,
    });
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Update a calendar event
 */
async function updateCalendarEvent(params: Record<string, unknown>) {
  const { eventId, subject, start, end, timeZone, body, location, attendees } = updateCalendarEventSchema.parse(params);
  
  logger.info('Tool: update-calendar-event', { 
    user: getContextUserId(),
    eventId,
  });
  
  try {
    const updates: Record<string, unknown> = {};
    
    if (subject !== undefined) updates.subject = subject;
    if (start !== undefined) {
      updates.start = { dateTime: start, timeZone: timeZone || 'UTC' };
    }
    if (end !== undefined) {
      updates.end = { dateTime: end, timeZone: timeZone || 'UTC' };
    }
    if (body !== undefined) {
      updates.body = { contentType: 'Text', content: body };
    }
    if (location !== undefined) {
      updates.location = { displayName: location };
    }
    if (attendees !== undefined) {
      updates.attendees = attendees.map(a => ({
        emailAddress: { address: a.email },
        type: a.type || 'required',
      }));
    }
    
    const response = await graphRequest(`/me/events/${eventId}`, {
      method: 'PATCH',
      body: updates,
    });
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Delete a calendar event
 */
async function deleteCalendarEvent(params: Record<string, unknown>) {
  const { eventId } = deleteCalendarEventSchema.parse(params);
  
  logger.info('Tool: delete-calendar-event', { 
    user: getContextUserId(),
    eventId,
  });
  
  try {
    const response = await graphRequest(`/me/events/${eventId}`, {
      method: 'DELETE',
    });
    
    if (response.status === 204 || response.ok) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: true, message: 'Event deleted' }),
        }],
      };
    }
    
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

// ============================================================================
// Tool Definitions for MCP
// ============================================================================

export const calendarToolDefinitions = [
  {
    name: 'list-calendars',
    description: 'List all calendars for the authenticated user',
    readOnly: true,
    inputSchema: {
      type: 'object' as const,
      properties: {},
    },
    handler: listCalendars,
  },
  {
    name: 'list-calendar-events',
    description: 'List calendar events from a calendar. Defaults to primary calendar.',
    readOnly: true,
    inputSchema: {
      type: 'object' as const,
      properties: {
        calendarId: {
          type: 'string',
          description: 'Calendar ID (default: primary calendar)',
        },
        top: {
          type: 'number',
          description: 'Maximum number of events to return (1-50, default: 10)',
        },
        skip: {
          type: 'number',
          description: 'Number of events to skip for pagination',
        },
        filter: {
          type: 'string',
          description: 'OData filter expression',
        },
        orderBy: {
          type: 'string',
          description: 'Sort order (default: start/dateTime)',
        },
      },
    },
    handler: listCalendarEvents,
  },
  {
    name: 'get-calendar-event',
    description: 'Get a single calendar event by its ID',
    readOnly: true,
    inputSchema: {
      type: 'object' as const,
      properties: {
        eventId: {
          type: 'string',
          description: 'The ID of the event to retrieve',
        },
      },
      required: ['eventId'],
    },
    handler: getCalendarEvent,
  },
  {
    name: 'get-calendar-view',
    description: 'Get calendar events within a specific time range',
    readOnly: true,
    inputSchema: {
      type: 'object' as const,
      properties: {
        startDateTime: {
          type: 'string',
          description: 'Start of time range (ISO 8601 format)',
        },
        endDateTime: {
          type: 'string',
          description: 'End of time range (ISO 8601 format)',
        },
        calendarId: {
          type: 'string',
          description: 'Calendar ID (default: primary calendar)',
        },
        top: {
          type: 'number',
          description: 'Maximum number of events to return',
        },
      },
      required: ['startDateTime', 'endDateTime'],
    },
    handler: getCalendarView,
  },
  {
    name: 'create-calendar-event',
    description: 'Create a new calendar event',
    readOnly: false,
    inputSchema: {
      type: 'object' as const,
      properties: {
        subject: {
          type: 'string',
          description: 'Event title',
        },
        start: {
          type: 'string',
          description: 'Start time (ISO 8601 format)',
        },
        end: {
          type: 'string',
          description: 'End time (ISO 8601 format)',
        },
        timeZone: {
          type: 'string',
          description: 'Time zone (default: UTC)',
        },
        body: {
          type: 'string',
          description: 'Event description',
        },
        bodyType: {
          type: 'string',
          enum: ['html', 'text'],
          description: 'Body content type (default: text)',
        },
        location: {
          type: 'string',
          description: 'Location name',
        },
        isAllDay: {
          type: 'boolean',
          description: 'Is this an all-day event',
        },
        reminderMinutesBeforeStart: {
          type: 'number',
          description: 'Reminder minutes before event',
        },
        calendarId: {
          type: 'string',
          description: 'Calendar ID (default: primary calendar)',
        },
      },
      required: ['subject', 'start', 'end'],
    },
    handler: createCalendarEvent,
  },
  {
    name: 'update-calendar-event',
    description: 'Update an existing calendar event',
    readOnly: false,
    inputSchema: {
      type: 'object' as const,
      properties: {
        eventId: {
          type: 'string',
          description: 'The ID of the event to update',
        },
        subject: {
          type: 'string',
          description: 'Event title',
        },
        start: {
          type: 'string',
          description: 'Start time (ISO 8601 format)',
        },
        end: {
          type: 'string',
          description: 'End time (ISO 8601 format)',
        },
        timeZone: {
          type: 'string',
          description: 'Time zone',
        },
        body: {
          type: 'string',
          description: 'Event description',
        },
        location: {
          type: 'string',
          description: 'Location name',
        },
      },
      required: ['eventId'],
    },
    handler: updateCalendarEvent,
  },
  {
    name: 'delete-calendar-event',
    description: 'Delete a calendar event',
    readOnly: false,
    inputSchema: {
      type: 'object' as const,
      properties: {
        eventId: {
          type: 'string',
          description: 'The ID of the event to delete',
        },
      },
      required: ['eventId'],
    },
    handler: deleteCalendarEvent,
  },
];
