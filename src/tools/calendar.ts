/**
 * Calendar tools for Microsoft Graph API
 */

import { z } from 'zod';
import { graphRequest, handleGraphResponse, formatErrorResponse } from '../graph/client.js';
import { serializeResponse } from '../utils/tonl.js';

// ============================================================================
// Day of Week Helper - Prevents LLM date calculation errors
// ============================================================================

const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/**
 * Get day of week from an ISO 8601 date string
 */
function getDayOfWeek(dateTimeStr: string): { day: string; date: string } | null {
  if (!dateTimeStr) return null;
  
  // Parse the date - handle both "2026-01-29T08:00:00.0000000" and "2026-01-29T08:00:00Z"
  const dateMatch = dateTimeStr.match(/^(\d{4}-\d{2}-\d{2})/);
  if (!dateMatch) return null;
  
  const datePart = dateMatch[1];
  const date = new Date(datePart + 'T12:00:00Z'); // Use noon UTC to avoid timezone edge cases
  
  if (isNaN(date.getTime())) return null;
  
  const dayIndex = date.getUTCDay();
  return {
    day: DAYS[dayIndex],
    date: datePart,
  };
}

/**
 * Enrich a calendar event with day of week information
 * Adds _dayInfo field to help LLMs avoid date calculation errors
 */
function enrichEventWithDayInfo(event: Record<string, unknown>): Record<string, unknown> {
  const start = event.start as { dateTime?: string } | undefined;
  const end = event.end as { dateTime?: string } | undefined;
  
  const startDayInfo = start?.dateTime ? getDayOfWeek(start.dateTime) : null;
  const endDayInfo = end?.dateTime ? getDayOfWeek(end.dateTime) : null;
  
  return {
    ...event,
    _dayInfo: {
      start: startDayInfo,
      end: endDayInfo,
    },
  };
}

/**
 * Enrich an array of calendar events with day of week information
 */
function enrichEventsWithDayInfo(events: unknown[]): unknown[] {
  return events.map(event => enrichEventWithDayInfo(event as Record<string, unknown>));
}

// ============================================================================
// Room Management
// ============================================================================

/**
 * Fetch all rooms from Microsoft Graph API
 * Returns minimal room data to reduce token usage
 */
async function fetchAllRooms(): Promise<Array<{
  displayName: string;
  emailAddress: string;
  building?: string;
  city?: string;
}>> {
  try {
    // Fetch rooms using /places endpoint
    const response = await graphRequest<{ value: unknown[] }>('/places/microsoft.graph.room');
    const rooms = (response.data as { value: unknown[] })?.value || [];
    
    // Extract only essential fields
    return rooms.map((room: unknown) => {
      const r = room as {
        displayName?: string;
        emailAddress?: string;
        building?: string;
        address?: { city?: string };
      };
      
      return {
        displayName: r.displayName || 'Unknown Room',
        emailAddress: r.emailAddress || '',
        building: r.building,
        city: r.address?.city,
      };
    }).filter(r => r.emailAddress); // Only include rooms with email
  } catch (error) {
    // If room fetch fails, return empty array (graceful degradation)
    return [];
  }
}

/**
 * Process findMeetingTimes response to reduce token usage
 * - Filters out busy attendees (keeps only free/tentative)
 * - Groups free rooms by location with limited examples
 * - Includes email addresses for booking
 */
function optimizeMeetingTimesResponse(
  suggestions: Array<Record<string, unknown>>,
  roomMetadata: Array<{ emailAddress: string; displayName: string; building?: string; city?: string }>
): Array<Record<string, unknown>> {
  // Create room lookup map
  const roomMap = new Map(
    roomMetadata.map(r => [r.emailAddress.toLowerCase(), r])
  );
  
  const MAX_ROOM_EXAMPLES = 5; // Limit examples per location
  
  return suggestions.map(suggestion => {
    const attendeeAvail = (suggestion.attendeeAvailability as Array<Record<string, unknown>>) || [];
    
    // Filter to only free/tentative attendees
    const availableAttendees = attendeeAvail.filter(a => {
      const availability = (a.availability as string) || '';
      return availability === 'free' || availability === 'tentative';
    });
    
    // Group free rooms by location with email addresses
    const roomsByLocation: Record<string, Array<{ name: string; email: string }>> = {};
    const nonRoomAttendees: Array<Record<string, unknown>> = [];
    
    for (const attendee of availableAttendees) {
      const email = ((attendee.attendee as Record<string, unknown>)?.emailAddress as Record<string, unknown>)?.address as string;
      const emailLower = email?.toLowerCase() || '';
      const room = roomMap.get(emailLower);
      
      if (room) {
        // This is a room - group by location
        const location = room.city || room.building || 'Unknown';
        if (!roomsByLocation[location]) {
          roomsByLocation[location] = [];
        }
        roomsByLocation[location].push({
          name: room.displayName || email,
          email: room.emailAddress,
        });
      } else {
        // This is a person - keep as-is
        nonRoomAttendees.push(attendee);
      }
    }
    
    // Build final room structure with counts and limited examples
    const freeRoomsByLocation: Record<string, { count: number; rooms: Array<{ name: string; email: string }> }> = {};
    for (const [location, rooms] of Object.entries(roomsByLocation)) {
      freeRoomsByLocation[location] = {
        count: rooms.length,
        rooms: rooms.slice(0, MAX_ROOM_EXAMPLES), // Limit to 5 examples
      };
    }
    
    // Build minimal suggestion object - remove redundant fields
    return {
      meetingTimeSlot: suggestion.meetingTimeSlot,
      _dayInfo: suggestion._dayInfo,
      _freeRoomsByLocation: freeRoomsByLocation,
      attendeeAvailability: nonRoomAttendees,
      // Remove: confidence, suggestionReason, locations, organizerAvailability
    };
  });
}

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
    type: z.enum(['required', 'optional', 'resource']).optional().default('required'),
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
  timeZone: z.string().optional().default('Europe/Berlin'),
  body: z.string().optional(),
  bodyType: z.enum(['html', 'text']).optional().default('text'),
  location: z.string().optional(),
  attendees: z.array(z.object({
    email: z.string(),
    type: z.enum(['required', 'optional']).optional().default('required'),
  })).optional(),
  isAllDay: z.boolean().optional().default(false),
  isOnlineMeeting: z.boolean().optional().default(true),
  reminderMinutesBeforeStart: z.number().optional(),
  calendarId: z.string().optional(),
});

const createDraftCalendarEventSchema = z.object({
  subject: z.string(),
  start: z.string(),
  end: z.string(),
  timeZone: z.string().optional().default('Europe/Berlin'),
  body: z.string().optional(),
  bodyType: z.enum(['html', 'text']).optional().default('text'),
  location: z.string().optional(),
  attendees: z.array(z.object({
    email: z.string(),
    type: z.enum(['required', 'optional']).optional().default('required'),
  })).optional(),
  isAllDay: z.boolean().optional().default(false),
  isOnlineMeeting: z.boolean().optional().default(true),
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
      
      // Enrich events with day of week info
      const data = response.data as { value?: unknown[] } | undefined;
      if (data?.value) {
        data.value = enrichEventsWithDayInfo(data.value);
      }
      
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
    
    // Enrich events with day of week info
    const data = response.data as { value?: unknown[] } | undefined;
    if (data?.value) {
      data.value = enrichEventsWithDayInfo(data.value);
    }
    
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
      
      // Enrich events with day of week info
      events = enrichEventsWithDayInfo(events);
      
      return {
        content: [{
          type: 'text' as const,
          text: serializeResponse({ value: events }),
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
    
    // Enrich events with day of week info
    events = enrichEventsWithDayInfo(events);
    
    return {
      content: [{
        type: 'text' as const,
        text: serializeResponse({ value: events }),
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
  let { attendees, durationMinutes, searchWindowStart, searchWindowEnd, meetingHoursStart, meetingHoursEnd, isOnlineMeeting, isOrganizerOptional, maxSuggestions, timeZone } = parsed;
  
  // Auto-fetch rooms for in-person meetings
  let roomMetadata: Array<{ emailAddress: string; displayName: string; building?: string; city?: string }> = [];
  if (isOnlineMeeting === false) {
    const rooms = await fetchAllRooms();
    
    // Store room metadata for later response processing
    roomMetadata = rooms;
    
    // Add rooms as resource attendees
    const roomAttendees = rooms.map(room => ({
      email: room.emailAddress,
      type: 'resource' as const,
    }));
    
    // Merge with existing attendees
    attendees = [...attendees, ...roomAttendees];
  }
  
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
          timeZone: timeZone || 'Europe/Berlin',
        },
        end: {
          dateTime: searchWindowEnd,
          timeZone: timeZone || 'Europe/Berlin',
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
        
        // Optimize response: filter busy attendees and group rooms by location
        const optimizedSuggestions = isOnlineMeeting === false 
          ? optimizeMeetingTimesResponse(limitedSuggestions, roomMetadata)
          : limitedSuggestions;
        
        // Enrich suggestions with day of week info
        const enrichedSuggestions = optimizedSuggestions.map(suggestion => {
          const timeSlot = suggestion.meetingTimeSlot as { start?: { dateTime?: string }; end?: { dateTime?: string } } | undefined;
          const startTime = timeSlot?.start?.dateTime;
          const endTime = timeSlot?.end?.dateTime;
          const startDayInfo = startTime ? getDayOfWeek(startTime) : null;
          const endDayInfo = endTime ? getDayOfWeek(endTime) : null;
          return {
            ...suggestion,
            _dayInfo: {
              start: startDayInfo,
              end: endDayInfo,
            },
          };
        });
        
        // Return filtered response
        return {
          content: [{
            type: 'text' as const,
            text: serializeResponse({
              ...data,
              meetingTimeSuggestions: enrichedSuggestions,
              _filteredByMeetingHours: {
                original: data.meetingTimeSuggestions.length,
                filtered: enrichedSuggestions.length,
                meetingHoursStart,
                meetingHoursEnd,
              },
            }),
          }],
        };
      }
    }
    
    // Enrich unfiltered suggestions with day info too
    const data = response.data as {
      meetingTimeSuggestions?: Array<{
        meetingTimeSlot?: {
          start?: { dateTime?: string };
          end?: { dateTime?: string };
        };
      }>;
    } | undefined;
    
    if (data?.meetingTimeSuggestions) {
      // Optimize response: filter busy attendees and group rooms by location
      const optimizedSuggestions = isOnlineMeeting === false
        ? optimizeMeetingTimesResponse(data.meetingTimeSuggestions, roomMetadata)
        : data.meetingTimeSuggestions;
      
      data.meetingTimeSuggestions = optimizedSuggestions.map(suggestion => {
        const timeSlot = suggestion.meetingTimeSlot as { start?: { dateTime?: string }; end?: { dateTime?: string } } | undefined;
        const startTime = timeSlot?.start?.dateTime;
        const endTime = timeSlot?.end?.dateTime;
        const startDayInfo = startTime ? getDayOfWeek(startTime) : null;
        const endDayInfo = endTime ? getDayOfWeek(endTime) : null;
        return {
          ...suggestion,
          _dayInfo: {
            start: startDayInfo,
            end: endDayInfo,
          },
        };
      }) as typeof data.meetingTimeSuggestions;
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
    
    // Enrich event with day of week info
    if (response.data) {
      response.data = enrichEventWithDayInfo(response.data as Record<string, unknown>);
    }
    
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
    
    // Enrich events with day of week info
    const data = response.data as { value?: unknown[] } | undefined;
    if (data?.value) {
      data.value = enrichEventsWithDayInfo(data.value);
    }
    
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
    location, attendees, isAllDay, isOnlineMeeting, reminderMinutesBeforeStart, calendarId 
  } = createCalendarEventSchema.parse(params);
  
  try {
    const event: Record<string, unknown> = {
      subject,
      start: {
        dateTime: start,
        timeZone: timeZone || 'Europe/Berlin',
      },
      end: {
        dateTime: end,
        timeZone: timeZone || 'Europe/Berlin',
      },
      isAllDay,
    };
    
    if (isOnlineMeeting) {
      event.isOnlineMeeting = true;
    }
    
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
 * Create a draft calendar event (without sending meeting invitations)
 */
async function createDraftCalendarEvent(params: Record<string, unknown>) {
  const { 
    subject, start, end, timeZone, body, bodyType, 
    location, attendees, isAllDay, isOnlineMeeting, reminderMinutesBeforeStart, calendarId 
  } = createDraftCalendarEventSchema.parse(params);
  
  try {
    const event: Record<string, unknown> = {
      subject,
      start: {
        dateTime: start,
        timeZone: timeZone || 'Europe/Berlin',
      },
      end: {
        dateTime: end,
        timeZone: timeZone || 'Europe/Berlin',
      },
      isAllDay,
      isDraft: true,
    };
    
    if (isOnlineMeeting) {
      event.isOnlineMeeting = true;
    }
    
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
      updates.start = { dateTime: start, timeZone: timeZone || 'Europe/Berlin' };
    }
    if (end !== undefined) {
      updates.end = { dateTime: end, timeZone: timeZone || 'Europe/Berlin' };
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
          text: serializeResponse({ success: true, message: 'Event deleted' }),
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

Use this for: chronological browsing, getting events in a date range.
For filtering by organizer/attendees/content, use search-calendar-events instead.

Examples:
- Events this week: { "startAfter": "2026-01-20T00:00:00Z", "startBefore": "2026-01-27T00:00:00Z" }
- Next 20 events: { "top": 20 }`,
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
    description: `Search calendar events with filtering by subject, organizer, attendees, online/in-person, all-day.

EMAIL ADDRESSES REQUIRED:
- organizerEmail/attendees expect EMAIL ADDRESSES (names don't work)
- If you only have names: First use search-mail with {"query": "John Doe", "top": 5} to find emails, extract them from results, then call this tool
- DO NOT ask user for emails - look them up yourself

Filtering notes:
- Always use date range (startAfter + startBefore) for reliable results
- attendees uses OR logic (matches if any attendee found)
- Location filtering not supported - use isOnlineMeeting for Teams vs in-person

Examples:
- By subject: { "subject": "Project X", "startAfter": "2026-01-01T00:00:00Z", "startBefore": "2026-01-31T23:59:59Z" }
- By organizer: { "organizerEmail": "alice@company.com", "startAfter": "2026-01-01T00:00:00Z" }
- Teams meetings: { "isOnlineMeeting": true, "startAfter": "2026-01-20T00:00:00Z", "startBefore": "2026-01-27T00:00:00Z" }
- With attendee: { "attendees": ["bob@company.com"], "startAfter": "2026-01-01T00:00:00Z" }`,
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
    description: `Find available meeting times when attendees are free. Checks everyone's free/busy status and returns ranked suggestions.

AUTOMATIC ROOM LOOKUP (when isOnlineMeeting=false):
- System automatically fetches all available rooms
- Includes rooms as resource attendees in availability check
- Results show which rooms are free for each time slot
- No need to manually specify rooms!

EMAIL ADDRESSES REQUIRED:
- If user provides names: First use search-mail with {"query": "Jane Smith", "top": 5} to find emails, extract from results, then call this tool
- DO NOT ask user for emails - look them up yourself

ORGANIZER AVAILABILITY (critical):
- By default, requires organizer (you) to be free
- If emptySuggestionsReason="OrganizerUnavailable": YOU are busy during entire window
- DO NOT silently retry with isOrganizerOptional=true
- Tell user: "I couldn't find times when you're available. Would you like times when other attendees are free, even if you have conflicts?"
- Only set isOrganizerOptional=true if user explicitly confirms

Parameters:
- meetingHoursStart/End: Limit to specific hours (e.g., "09:00:00" to "17:00:00")
- isOnlineMeeting: true (default, Teams meeting), false (in-person with automatic room lookup)
- Returns: confidence score, attendee availability, suggestion reasons

Examples:
- 1-hour, 9-11am: {"attendees": [{"email": "alice@company.com", "type": "required"}], "durationMinutes": 60, "searchWindowStart": "2026-01-20T00:00:00", "searchWindowEnd": "2026-02-03T23:59:59", "meetingHoursStart": "09:00:00", "meetingHoursEnd": "11:00:00"}
- 30-min Teams: {"attendees": [{"email": "alice@company.com"}, {"email": "bob@company.com"}], "durationMinutes": 30, "searchWindowStart": "2026-01-20T00:00:00", "searchWindowEnd": "2026-01-24T23:59:59", "isOnlineMeeting": true}
- In-person: {"attendees": [{"email": "alice@company.com"}], "durationMinutes": 60, "searchWindowStart": "2026-01-20T00:00:00", "searchWindowEnd": "2026-01-24T23:59:59", "isOnlineMeeting": false}

After finding times, use create-calendar-event to book.`,
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
    description: `Create a new calendar event.

TEAMS MEETINGS:
- Set isOnlineMeeting=true (default) to auto-generate Teams meeting link
- Link appears in onlineMeetingUrl field and is included in invitations

IN-PERSON MEETINGS:
- Set isOnlineMeeting=false for physical meetings
- Set location field to room name/address`,
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
          description: 'Time zone (default: Europe/Berlin)',
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
    name: 'create-draft-calendar-event',
    description: `Create a calendar event draft without sending invitations. Event is saved to calendar but attendees NOT notified until user sends from Outlook.

TEAMS MEETINGS:
- Set isOnlineMeeting=true (default) to auto-generate Teams meeting link
- Link appears in onlineMeetingUrl field and is included in invitations

IN-PERSON MEETINGS:
- Set isOnlineMeeting=false for physical meetings
- Set location field to room name/address

Use this to prepare meetings for user review before sending invitations. Safer than create-calendar-event (which sends immediately).

Draft event appears with "[Draft]" indicator. User can send from Outlook or you can use update-calendar-event to modify.`,
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
          description: 'Time zone (default: Europe/Berlin)',
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
        isOnlineMeeting: {
          type: 'boolean',
          description: 'Create as Teams/online meeting (auto-generates meeting link). Default: true',
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
          description: 'List of attendees to invite. Invitations are NOT sent until the meeting is published from Outlook.',
        },
      },
      required: ['subject', 'start', 'end'],
    },
    handler: createDraftCalendarEvent,
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
