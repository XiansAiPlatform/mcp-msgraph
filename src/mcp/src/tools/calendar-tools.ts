import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { logger } from "../logger.js";
import { GraphApiHelper } from "../graph-api-helper.js";

// Helper schema for boolean parameters that accepts both boolean and string values
const zBooleanParam = () => z.union([
  z.boolean(),
  z.string().transform(str => str.toLowerCase() === 'true')
]);

export function registerCalendarTools(
  server: McpServer,
  getApiHelper: () => GraphApiHelper | null
) {
  // Create calendar event tool
  server.tool(
    "create-calendar-event",
    "Create a calendar event using Microsoft Graph API. Requires Calendars.ReadWrite permission.",
    {
      userEmail: z.string().describe("Email address of the user on whose behalf to create the calendar event"),
      subject: z.string().describe("Event subject/title"),
      start: z.object({
        dateTime: z.string().describe("Start date and time in ISO 8601 format (e.g., '2024-01-20T14:00:00')"),
        timeZone: z.string().describe("Time zone (e.g., 'Pacific Standard Time', 'UTC')")
      }).describe("Event start date and time"),
      end: z.object({
        dateTime: z.string().describe("End date and time in ISO 8601 format"),
        timeZone: z.string().describe("Time zone")
      }).describe("Event end date and time"),
      location: z.object({
        displayName: z.string().describe("Location name")
      }).optional().describe("Event location"),
      body: z.object({
        content: z.string().describe("Event description/body"),
        contentType: z.enum(["text", "html"]).optional().default("html")
      }).optional().describe("Event body/description"),
      attendees: z.array(z.object({
        email: z.string().describe("Attendee email address"),
        name: z.string().optional().describe("Attendee display name"),
        type: z.enum(["required", "optional"]).optional().default("required")
      })).optional().describe("List of attendees"),
      isOnlineMeeting: zBooleanParam().optional().default(false).describe("Whether to create an online meeting (Teams)"),
      reminder: z.union([z.number(), z.string().transform(str => parseInt(str, 10))]).optional().describe("Reminder in minutes before the event"),
      isAllDay: zBooleanParam().optional().default(false).describe("Whether this is an all-day event"),
      showAs: z.enum(["free", "tentative", "busy", "oof", "workingElsewhere", "unknown"]).optional().default("busy"),
      sensitivity: z.enum(["normal", "personal", "private", "confidential"]).optional().default("normal")
    },
    async ({ userEmail, subject, start, end, location, body, attendees, isOnlineMeeting, reminder, isAllDay, showAs, sensitivity }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        // Build the event object
        const event: any = {
          subject,
          start,
          end,
          isAllDay,
          showAs,
          sensitivity
        };

        // Add location if provided
        if (location) {
          event.location = location;
        }

        // Add body if provided
        if (body) {
          event.body = body;
        }

        // Add attendees if provided
        if (attendees && attendees.length > 0) {
          event.attendees = attendees.map(attendee => ({
            emailAddress: {
              address: attendee.email,
              name: attendee.name
            },
            type: attendee.type
          }));
        }

        // Add online meeting if requested
        if (isOnlineMeeting) {
          event.isOnlineMeeting = true;
          event.onlineMeetingProvider = "teamsForBusiness";
        }

        // Add reminder if provided
        if (reminder !== undefined) {
          event.reminderMinutesBeforeStart = reminder;
          event.isReminderOn = true;
        }

        logger.info(`Creating calendar event: ${subject} from ${start.dateTime} to ${end.dateTime} on behalf of ${userEmail}`);

        const response = await apiHelper.callGraphApi({
          path: `/users/${userEmail}/events`,
          method: 'post',
          body: event
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const createdEvent = response.data;
        let resultText = `Calendar event created successfully!\n\n`;
        resultText += `Subject: ${createdEvent.subject}\n`;
        resultText += `Start: ${createdEvent.start.dateTime} (${createdEvent.start.timeZone})\n`;
        resultText += `End: ${createdEvent.end.dateTime} (${createdEvent.end.timeZone})\n`;
        if (location) {
          resultText += `Location: ${location.displayName}\n`;
        }
        if (attendees && attendees.length > 0) {
          resultText += `Attendees: ${attendees.map(a => a.email).join(', ')}\n`;
        }
        if (isOnlineMeeting && createdEvent.onlineMeeting) {
          resultText += `\nOnline Meeting URL: ${createdEvent.onlineMeeting.joinUrl}\n`;
        }
        resultText += `\nEvent ID: ${createdEvent.id}`;

        return {
          content: [{ 
            type: "text" as const, 
            text: resultText
          }],
        };

      } catch (error: any) {
        logger.error("Error creating calendar event:", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the Calendars.ReadWrite permission is granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error creating calendar event: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  // List calendar events tool
  server.tool(
    "list-calendar-events",
    "List calendar events for a user. Requires Calendars.Read permission.",
    {
      userId: z.string().describe("User email address or user ID to list calendar events for"),
      startDateTime: z.string().optional().describe("Start date/time in ISO format (e.g., '2024-01-01T00:00:00Z')"),
      endDateTime: z.string().optional().describe("End date/time in ISO format"),
      filter: z.string().optional().describe("OData filter query"),
      select: z.array(z.string()).optional().describe("Properties to include in the response"),
      top: z.union([z.number(), z.string().transform(str => parseInt(str, 10))]).optional().default(10).describe("Number of events to return"),
      orderBy: z.string().optional().default("start/dateTime").describe("Property to sort by"),
      fetchAll: zBooleanParam().optional().default(false).describe("Whether to fetch all events")
    },
    async ({ userId, startDateTime, endDateTime, filter, select, top, orderBy, fetchAll }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        const queryParams: Record<string, string> = {};
        
        // Build filter for date range if provided
        let dateFilter = "";
        if (startDateTime && endDateTime) {
          dateFilter = `start/dateTime ge '${startDateTime}' and start/dateTime le '${endDateTime}'`;
        } else if (startDateTime) {
          dateFilter = `start/dateTime ge '${startDateTime}'`;
        } else if (endDateTime) {
          dateFilter = `start/dateTime le '${endDateTime}'`;
        }

        // Combine with any additional filter
        if (dateFilter && filter) {
          queryParams.$filter = `(${dateFilter}) and (${filter})`;
        } else if (dateFilter) {
          queryParams.$filter = dateFilter;
        } else if (filter) {
          queryParams.$filter = filter;
        }

        if (select && select.length > 0) {
          queryParams.$select = select.join(',');
        }
        if (!fetchAll && top) {
          queryParams.$top = top.toString();
        }
        if (orderBy) {
          queryParams.$orderby = orderBy;
        }

        logger.info(`Listing calendar events for user: ${userId}`);

        const response = await apiHelper.callGraphApi({
          path: `/users/${userId}/events`,
          method: 'get',
          queryParams,
          fetchAll
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const events = response.data.value || [];
        let resultText = `Found ${events.length} calendar event(s):\n\n`;
        
        events.forEach((event: any, index: number) => {
          resultText += `${index + 1}. ${event.subject}\n`;
          resultText += `   Start: ${event.start.dateTime} (${event.start.timeZone})\n`;
          resultText += `   End: ${event.end.dateTime} (${event.end.timeZone})\n`;
          if (event.location?.displayName) {
            resultText += `   Location: ${event.location.displayName}\n`;
          }
          if (event.isOnlineMeeting && event.onlineMeeting?.joinUrl) {
            resultText += `   Online Meeting: ${event.onlineMeeting.joinUrl}\n`;
          }
          if (event.attendees && event.attendees.length > 0) {
            const attendeeList = event.attendees.map((a: any) => a.emailAddress.address).join(', ');
            resultText += `   Attendees: ${attendeeList}\n`;
          }
          resultText += '\n';
        });

        if (!fetchAll && apiHelper.hasMorePages(response.data, 'graph')) {
          resultText += `Note: More events are available. Use 'fetchAll: true' to retrieve all events.`;
        }

        return {
          content: [{ 
            type: "text" as const, 
            text: resultText
          }],
        };

      } catch (error: any) {
        logger.error("Error listing calendar events:", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the Calendars.Read permission is granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error listing calendar events: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );
}
