/**
 * Calendar tools for Microsoft Graph API
 */

import { z } from 'zod';
import { graphRequest, handleGraphResponse, formatErrorResponse } from '../graph/client.js';

// ============================================================================
// Schemas
// ============================================================================

const listCalendarEventsSchema = z.object({
  calendarId: z.string().optional(),
  startAfter: z.string().optional(),
  startBefore: z.string().optional(),
  top: z.number().min(1).max(50).optional().default(10),
  skip: z.number().min(0).optional(),
  orderBy: z.string().optional().default('start/dateTime'),
});

const searchCalendarEventsSchema = z.object({
  subject: z.string().optional(),
  organizerEmail: z.string().optional(),
  organizerName: z.string().optional(),
  attendees: z.array(z.string()).optional(),
  isOnlineMeeting: z.boolean().optional(),
  isAllDay: z.boolean().optional(),
  startAfter: z.string().optional(),
  startBefore: z.string().optional(),
  top: z.number().min(1).max(50).optional().default(25),
});

const findMeetingTimesSchema = z.object({
  attendees: z.array(z.object({
    email: z.string(),
    type: z.enum(['required', 'optional']).optional().default('required'),
  })).min(1),
  durationMinutes: z.number().min(15).max(480).default(60),
  searchWindowStart: z.string(),
  searchWindowEnd: z.string(),
  meetingHoursStart: z.string().optional(),
  meetingHoursEnd: z.string().optional(),
  isOnlineMeeting: z.boolean().optional().default(false),
  isOrganizerOptional: z.boolean().optional().default(false),
  maxSuggestions: z.number().min(1).max(50).optional().default(10),
  timeZone: z.string().optional(),
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
  try {
    const response = await graphRequest<{ value: unknown[] }>('/me/calendars');
    return handleGraphResponse(response);
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * List calendar events with date range filtering
 * Uses calendarView for proper recurring event expansion
 */
async function listCalendarEvents(params: Record<string, unknown>) {
  const { calendarId, startAfter, startBefore, top, skip, orderBy } = listCalendarEventsSchema.parse(params);
  
  try {
    const queryParams = new URLSearchParams();
    
    if (top) queryParams.set('$top', String(top));
    if (skip) queryParams.set('$skip', String(skip));
    if (orderBy) queryParams.set('$orderby', orderBy);
    
    queryParams.set('$select', 'id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,isOnlineMeeting,onlineMeetingUrl,bodyPreview');
    
    // If date range is provided, use calendarView for proper recurring event expansion
    if (startAfter && startBefore) {
      queryParams.set('startDateTime', startAfter);
      queryParams.set('endDateTime', startBefore);
      
      const endpoint = calendarId 
        ? `/me/calendars/${calendarId}/calendarView`
        : '/me/calendarView';
      
      const url = `${endpoint}?${queryParams.toString()}`;
      const response = await graphRequest<{ value: unknown[] }>(url);
      return handleGraphResponse(response);
    }
    
    // Otherwise use events endpoint with filter
    if (startAfter || startBefore) {
      const filters: string[] = [];
      if (startAfter) filters.push(`start/dateTime ge '${startAfter}'`);
      if (startBefore) filters.push(`start/dateTime le '${startBefore}'`);
      queryParams.set('$filter', filters.join(' and '));
    }
    
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
 * Build filter expression for calendar event search
 */
/**
 * Format a free-text query parameter for calendar search
 * 
 * - If query looks like raw KQL (contains operators), pass through as-is
 * - If query is a multi-word phrase, quote it
 * - Single terms pass through unquoted
 */
function formatCalendarQuery(query: string): string {
  const cleaned = query.trim();
  
  // Check if this looks like raw KQL (contains operators)
  const hasKqlOperators = /\b(AND|OR|NOT)\b/i.test(cleaned);
  
  if (hasKqlOperators) {
    // Raw KQL - pass through as-is
    return cleaned;
  }
  
  // Only quote multi-word phrases (those with spaces)
  // Single terms like "deadline", "project" work better unquoted
  if (/\s/.test(cleaned)) {
    // Multi-word phrase - escape existing quotes and wrap
    const escaped = cleaned.replace(/"/g, '\\"');
    return `\\"${escaped}\\"`;
  }
  
  // Single term - pass through unquoted
  return cleaned;
}

/**
 * Build OData filter expression for calendar events
 * 
 * Note: Graph API calendar events have severe limitations:
 * - NO filtering on organizer/emailAddress/address (causes 500 error)
 * - organizerEmail must be filtered client-side
 * - location filtering removed (unreliable) - use isOnlineMeeting instead
 */
function buildCalendarFilter(params: {
  subject?: string;
  organizerEmail?: string;
  organizerName?: string;
  isAllDay?: boolean;
  startAfter?: string;
  startBefore?: string;
}): string | undefined {
  const filters: string[] = [];
  
  if (params.subject) {
    filters.push(`contains(subject, '${params.subject}')`);
  }
  // organizerEmail is NOT included - Graph API doesn't support it (500 error)
  // It will be filtered client-side in searchCalendarEvents
  if (params.organizerName) {
    filters.push(`contains(organizer/emailAddress/name, '${params.organizerName}')`);
  }
  if (params.isAllDay !== undefined) {
    filters.push(`isAllDay eq ${params.isAllDay}`);
  }
  if (params.startAfter) {
    filters.push(`start/dateTime ge '${params.startAfter}'`);
  }
  if (params.startBefore) {
    filters.push(`start/dateTime le '${params.startBefore}'`);
  }
  
  return filters.length > 0 ? filters.join(' and ') : undefined;
}

/**
 * Search calendar events with advanced filtering
 */
async function searchCalendarEvents(params: Record<string, unknown>) {
  const parsed = searchCalendarEventsSchema.parse(params);
  const { subject, organizerEmail, organizerName, attendees, isOnlineMeeting, isAllDay, startAfter, startBefore, top } = parsed;
  
  try {
    const queryParams = new URLSearchParams();
    
    if (top) queryParams.set('$top', String(top));
    queryParams.set('$select', 'id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,isOnlineMeeting,onlineMeetingUrl,bodyPreview');
    
    // Build filter from parameters
    const filter = buildCalendarFilter({ subject, organizerEmail, organizerName, isAllDay, startAfter, startBefore });
    
    // If using date range with both dates, use calendarView
    if (startAfter && startBefore) {
      queryParams.set('startDateTime', startAfter);
      queryParams.set('endDateTime', startBefore);
      
      // Apply additional filters if any (except date ones which are in URL params)
      const nonDateFilter = buildCalendarFilter({ subject, organizerEmail, organizerName, isAllDay });
      if (nonDateFilter) queryParams.set('$filter', nonDateFilter);
      
      const url = `/me/calendarView?${queryParams.toString()}`;
      const response = await graphRequest<{ value: unknown[] }>(url);
      
      // Post-filter for properties not supported in $filter
      let events = (response.data as { value: unknown[] })?.value || [];
      
      // Filter by organizerEmail (Graph API doesn't support this in $filter)
      if (organizerEmail) {
        const emailLower = organizerEmail.toLowerCase();
        events = events.filter((e: unknown) => {
          const event = e as { organizer?: { emailAddress?: { address?: string } } };
          return event.organizer?.emailAddress?.address?.toLowerCase().includes(emailLower);
        });
      }
      
      if (isOnlineMeeting !== undefined) {
        events = events.filter((e: unknown) => {
          const event = e as { isOnlineMeeting?: boolean };
          return event.isOnlineMeeting === isOnlineMeeting;
        });
      }
      
      if (attendees?.length) {
        events = events.filter((e: unknown) => {
          const event = e as { attendees?: Array<{ emailAddress?: { address?: string } }> };
          const eventAttendees = event.attendees || [];
          return attendees.some(searchAttendee => 
            eventAttendees.some(ea => 
              ea.emailAddress?.address?.toLowerCase().includes(searchAttendee.toLowerCase())
            )
          );
        });
      }
      
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ value: events }, null, 2),
        }],
      };
    }
    
    // Without date range, use events endpoint
    if (filter) queryParams.set('$filter', filter);
    
    const url = `/me/events?${queryParams.toString()}`;
    const response = await graphRequest<{ value: unknown[] }>(url);
    
    // Post-filter for properties not supported in $filter
    let events = (response.data as { value: unknown[] })?.value || [];
    
    // Filter by organizerEmail (Graph API doesn't support this in $filter)
    if (organizerEmail) {
      const emailLower = organizerEmail.toLowerCase();
      events = events.filter((e: unknown) => {
        const event = e as { organizer?: { emailAddress?: { address?: string } } };
        return event.organizer?.emailAddress?.address?.toLowerCase().includes(emailLower);
      });
    }
    
    if (isOnlineMeeting !== undefined) {
      events = events.filter((e: unknown) => {
        const event = e as { isOnlineMeeting?: boolean };
        return event.isOnlineMeeting === isOnlineMeeting;
      });
    }
    
    if (attendees?.length) {
      events = events.filter((e: unknown) => {
        const event = e as { attendees?: Array<{ emailAddress?: { address?: string } }> };
        const eventAttendees = event.attendees || [];
        return attendees.some(searchAttendee => 
          eventAttendees.some(ea => 
            ea.emailAddress?.address?.toLowerCase().includes(searchAttendee.toLowerCase())
          )
        );
      });
    }
    
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ value: events }, null, 2),
      }],
    };
  } catch (error) {
    return formatErrorResponse(error);
  }
}

/**
 * Check if a time (HH:MM:SS or HH:MM) is within meeting hours constraints
 */
function isWithinMeetingHours(
  dateTimeStr: string,
  meetingHoursStart?: string,
  meetingHoursEnd?: string
): boolean {
  if (!meetingHoursStart || !meetingHoursEnd) return true;
  
  // Extract time from ISO datetime (e.g., "2026-01-27T09:00:00" -> "09:00:00")
  const timeMatch = dateTimeStr.match(/T(\d{2}:\d{2})/);
  if (!timeMatch) return true;
  
  const time = timeMatch[1]; // "HH:MM"
  const startTime = meetingHoursStart.substring(0, 5); // "HH:MM" from "HH:MM:SS"
  const endTime = meetingHoursEnd.substring(0, 5);
  
  return time >= startTime && time < endTime;
}

/**
 * Find available meeting times when all attendees are free
 */
async function findMeetingTimes(params: Record<string, unknown>) {
  const parsed = findMeetingTimesSchema.parse(params);
  const { attendees, durationMinutes, searchWindowStart, searchWindowEnd, meetingHoursStart, meetingHoursEnd, isOnlineMeeting, isOrganizerOptional, maxSuggestions, timeZone } = parsed;
  
  try {
    // Build the request body for findMeetingTimes
    const requestBody: Record<string, unknown> = {
      attendees: attendees.map(a => ({
        emailAddress: { address: a.email },
        type: a.type || 'required',
      })),
      meetingDuration: `PT${durationMinutes}M`,
      maxCandidates: maxSuggestions ? maxSuggestions * 3 : 30, // Request more to allow for filtering
      returnSuggestionReasons: true,
      isOrganizerOptional: isOrganizerOptional ?? false,
    };
    
    // Build time constraint
    const timeConstraint: Record<string, unknown> = {
      activityDomain: 'work',
      timeSlots: [{
        start: {
          dateTime: searchWindowStart,
          timeZone: timeZone || 'UTC',
        },
        end: {
          dateTime: searchWindowEnd,
          timeZone: timeZone || 'UTC',
        },
      }],
    };
    
    requestBody.timeConstraint = timeConstraint;
    
    // Add location constraint for online meeting
    if (isOnlineMeeting) {
      requestBody.locationConstraint = {
        isRequired: false,
        suggestLocation: false,
        locations: [{
          displayName: 'Microsoft Teams Meeting',
          locationUri: '',
        }],
      };
    }
    
    const response = await graphRequest('/me/findMeetingTimes', {
      method: 'POST',
      body: requestBody,
      headers: timeZone ? { 'Prefer': `outlook.timezone="${timeZone}"` } : undefined,
    });
    
    // Client-side filtering for meeting hours constraint
    // Graph API doesn't reliably enforce meetingHoursStart/End, so we filter here
    if (meetingHoursStart && meetingHoursEnd && response.data) {
      const data = response.data as {
        meetingTimeSuggestions?: Array<{
          meetingTimeSlot?: {
            start?: { dateTime?: string };
            end?: { dateTime?: string };
          };
        }>;
        emptySuggestionsReason?: string;
      };
      
      if (data.meetingTimeSuggestions && data.meetingTimeSuggestions.length > 0) {
        // Filter suggestions to only include those within meeting hours
        const filteredSuggestions = data.meetingTimeSuggestions.filter(suggestion => {
          const startTime = suggestion.meetingTimeSlot?.start?.dateTime;
          if (!startTime) return true;
          return isWithinMeetingHours(startTime, meetingHoursStart, meetingHoursEnd);
        });
        
        // Limit to requested maxSuggestions
        const limitedSuggestions = filteredSuggestions.slice(0, maxSuggestions || 10);
        
        // Return filtered response
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({
              ...data,
              meetingTimeSuggestions: limitedSuggestions,
              _filteredByMeetingHours: {
                original: data.meetingTimeSuggestions.length,
                filtered: limitedSuggestions.length,
                meetingHoursStart,
                meetingHoursEnd,
              },
            }, null, 2),
          }],
        };
      }
    }
    
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
    description: 'List all calendars for the authenticated user. Returns calendar IDs that can be used with other calendar tools.',
    readOnly: true,
    requiredScopes: ['Calendars.Read'],
    inputSchema: {
      type: 'object' as const,
      properties: {},
    },
    handler: listCalendars,
  },
  {
    name: 'list-calendar-events',
    description: `List calendar events with simple date range filtering. Uses calendarView for proper recurring event expansion.

Use this tool for:
- Browsing upcoming events chronologically
- Getting events in a specific date range
- Simple event listing without complex filters

For searching by content/organizer/location/attendees, use search-calendar-events instead.

Examples:
- Get events this week: { "startAfter": "2026-01-20T00:00:00Z", "startBefore": "2026-01-27T00:00:00Z" }
- Get next 20 events: { "top": 20 }`,
    readOnly: true,
    requiredScopes: ['Calendars.Read'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        calendarId: {
          type: 'string',
          description: 'Calendar ID (default: primary calendar). Use list-calendars to get IDs.',
        },
        startAfter: {
          type: 'string',
          description: 'Events starting after this date/time (ISO 8601). Example: "2026-01-20T00:00:00Z"',
        },
        startBefore: {
          type: 'string',
          description: 'Events starting before this date/time (ISO 8601). Example: "2026-01-27T23:59:59Z"',
        },
        top: {
          type: 'number',
          description: 'Maximum number of events to return (1-50, default: 10)',
        },
        skip: {
          type: 'number',
          description: 'Number of events to skip for pagination',
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
    name: 'search-calendar-events',
    description: `Search calendar events with advanced filtering by subject, organizer, attendees, online/in-person, and more.

CRITICAL - EMAIL ADDRESSES REQUIRED:
- If user provides PEOPLE NAMES (e.g., "John Doe", "Jane Smith"), names will NOT work
- FIRST use 'search-mail' with "query" parameter (e.g., {"query": "John Doe", "top": 5}) to find emails
- EXTRACT their email addresses from the mail results (look in from/toRecipients fields)
- THEN call this tool with the extracted email addresses
- DO NOT ask the user for email addresses - look them up yourself using search-mail

Use this tool for:
- Filtering by subject keywords
- Filtering by organizer (requires email address)
- Filtering by attendees (requires email addresses)
- Finding Teams/online meetings vs in-person meetings (use isOnlineMeeting)
- Finding all-day events

IMPORTANT:
- Always use DATE RANGE (startAfter + startBefore) for reliable results
- Location filtering is not supported - use isOnlineMeeting to filter Teams vs in-person
- Subject filtering is case-insensitive

Examples:
- Find meetings about Project X: { "subject": "Project X", "startAfter": "2026-01-01T00:00:00Z", "startBefore": "2026-01-31T23:59:59Z" }
- Find meetings organized by Alice: { "organizerEmail": "alice@company.com", "startAfter": "2026-01-01T00:00:00Z" }
- Find all Teams meetings this week: { "isOnlineMeeting": true, "startAfter": "2026-01-20T00:00:00Z", "startBefore": "2026-01-27T00:00:00Z" }
- Find meetings with Bob: { "attendees": ["bob@company.com"], "startAfter": "2026-01-01T00:00:00Z" }
- Find all-day events: { "isAllDay": true, "startAfter": "2026-01-01T00:00:00Z", "startBefore": "2026-01-31T23:59:59Z" }`,
    readOnly: true,
    requiredScopes: ['Calendars.Read'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        subject: {
          type: 'string',
          description: 'Filter by subject containing this text (case-insensitive)',
        },
        organizerEmail: {
          type: 'string',
          description: 'Organizer EMAIL ADDRESS. If you only have a name, use search-mail with query parameter (e.g., {"query": "John Doe"}) to find their email first.',
        },
        organizerName: {
          type: 'string',
          description: 'Organizer name (UNRELIABLE - avoid using)',
        },
        attendees: {
          type: 'array',
          items: { type: 'string' },
          description: 'Attendee EMAIL ADDRESSES. If you only have names, use search-mail with query parameter to find their emails first. OR logic - matches if any attendee is found.',
        },
        isOnlineMeeting: {
          type: 'boolean',
          description: 'Filter for Teams/online meetings (true) or in-person meetings (false)',
        },
        isAllDay: {
          type: 'boolean',
          description: 'Filter for all-day events only (true) or timed events only (false)',
        },
        startAfter: {
          type: 'string',
          description: 'Events starting after this date/time (ISO 8601). Example: "2026-01-20T00:00:00Z"',
        },
        startBefore: {
          type: 'string',
          description: 'Events starting before this date/time (ISO 8601). Example: "2026-01-27T23:59:59Z"',
        },
        top: {
          type: 'number',
          description: 'Maximum number of events to return (default: 25)',
        },
      },
    },
    handler: searchCalendarEvents,
  },
  {
    name: 'find-meeting-times',
    description: `Find available meeting times when all specified attendees are free. Uses Microsoft's scheduling algorithm to suggest optimal time slots.

This is the key tool for scheduling meetings with multiple people. It checks everyone's free/busy status (same access as Outlook) and returns ranked suggestions.

CRITICAL - EMAIL ADDRESSES REQUIRED:
- If user provides PEOPLE NAMES, DO NOT ask the user for emails
- FIRST use 'search-mail' with "query" parameter (e.g., {"query": "Jane Smith", "top": 5}) to find emails
- EXTRACT their email addresses from the mail results (look in from/toRecipients fields)
- THEN call this tool with the extracted email addresses

MEETING HOURS CONSTRAINT:
- Use meetingHoursStart/meetingHoursEnd to limit suggestions to specific hours (e.g., 9 AM - 5 PM)
- These constraints are strictly enforced via client-side filtering

ORGANIZER AVAILABILITY:
- By default, suggestions require the organizer (you) to be free
- If response contains emptySuggestionsReason="OrganizerUnavailable", it means YOU are busy during the entire search window
- DO NOT silently retry with isOrganizerOptional=true
- Instead, tell the user: "I couldn't find times when you're available in this window. Would you like me to find times when the other attendees are free, even if you have conflicts?"
- Only set isOrganizerOptional=true if user explicitly confirms they want to see options despite their conflicts

Returns a list of suggested time slots with:
- confidence score (0-100%) based on attendee availability
- attendee availability status for each slot
- suggestion reasons

Examples:
- Find 1-hour slot with Alice next 2 weeks, 9-11am:
  {
    "attendees": [{"email": "alice@company.com", "type": "required"}],
    "durationMinutes": 60,
    "searchWindowStart": "2026-01-20T00:00:00",
    "searchWindowEnd": "2026-02-03T23:59:59",
    "meetingHoursStart": "09:00:00",
    "meetingHoursEnd": "11:00:00",
    "timeZone": "Europe/Berlin"
  }

- Find 30-min Teams meeting with team:
  {
    "attendees": [
      {"email": "alice@company.com", "type": "required"},
      {"email": "bob@company.com", "type": "required"},
      {"email": "carol@company.com", "type": "optional"}
    ],
    "durationMinutes": 30,
    "searchWindowStart": "2026-01-20T00:00:00",
    "searchWindowEnd": "2026-01-24T23:59:59",
    "isOnlineMeeting": true,
    "maxSuggestions": 5
  }

After finding times, use create-calendar-event to book the meeting.`,
    readOnly: true,
    requiredScopes: ['Calendars.Read.Shared'],
    inputSchema: {
      type: 'object' as const,
      properties: {
        attendees: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              email: { type: 'string', description: 'Attendee email address (e.g., "john@company.com"). If you only have a name, use search-mail to find their email first.' },
              type: { type: 'string', enum: ['required', 'optional'], description: 'required = must attend, optional = nice to have' },
            },
            required: ['email'],
          },
          description: 'List of attendees. Use simplified format: [{"email": "john@company.com", "type": "required"}]. Do NOT use Graph API format like emailAddress.address. If you only have names, use search-mail to find emails first.',
        },
        durationMinutes: {
          type: 'number',
          description: 'Meeting duration in minutes. Common values: 15, 30, 45, 60, 90, 120',
        },
        searchWindowStart: {
          type: 'string',
          description: 'Start of search window (ISO 8601). Example: "2026-01-20T00:00:00"',
        },
        searchWindowEnd: {
          type: 'string',
          description: 'End of search window (ISO 8601). Example: "2026-02-03T23:59:59"',
        },
        meetingHoursStart: {
          type: 'string',
          description: 'Earliest time of day for meetings (HH:MM:SS). Example: "09:00:00" for 9 AM',
        },
        meetingHoursEnd: {
          type: 'string',
          description: 'Latest time of day for meetings (HH:MM:SS). Example: "17:00:00" for 5 PM',
        },
        isOnlineMeeting: {
          type: 'boolean',
          description: 'Suggest as Teams/online meeting (default: false)',
        },
        isOrganizerOptional: {
          type: 'boolean',
          description: 'If true, suggestions can be returned even when you are busy. IMPORTANT: Only set to true after asking the user if they want to see times despite their conflicts. Never set this silently. Default: false',
        },
        maxSuggestions: {
          type: 'number',
          description: 'Maximum number of time slot suggestions to return (default: 10, max: 50)',
        },
        timeZone: {
          type: 'string',
          description: 'Time zone for the constraints. Example: "Europe/Berlin", "America/New_York", "UTC"',
        },
      },
      required: ['attendees', 'searchWindowStart', 'searchWindowEnd'],
    },
    handler: findMeetingTimes,
  },
  {
    name: 'get-calendar-event',
    description: 'Get a single calendar event by its ID',
    readOnly: true,
    requiredScopes: ['Calendars.Read'],
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
    requiredScopes: ['Calendars.Read'],
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
    requiredScopes: ['Calendars.ReadWrite'],
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
        attendees: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              email: { type: 'string', description: 'Attendee email address (e.g., "john@company.com")' },
              type: { type: 'string', enum: ['required', 'optional'], description: 'Attendance type (default: required)' },
            },
            required: ['email'],
          },
          description: 'List of attendees to invite. Use simplified format: [{"email": "john@company.com", "type": "required"}].',
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
    requiredScopes: ['Calendars.ReadWrite'],
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
        attendees: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              email: { type: 'string', description: 'Attendee email address (e.g., "john@company.com")' },
              type: { type: 'string', enum: ['required', 'optional'], description: 'Attendance type (default: required)' },
            },
            required: ['email'],
          },
          description: 'List of attendees (replaces existing). Use simplified format: [{"email": "john@company.com", "type": "required"}]. Do NOT use Graph API format like emailAddress.address.',
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
    requiredScopes: ['Calendars.ReadWrite'],
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
