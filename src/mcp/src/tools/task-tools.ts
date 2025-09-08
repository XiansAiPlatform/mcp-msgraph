import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { logger } from "../logger.js";
import { GraphApiHelper } from "../graph-api-helper.js";

export function registerTaskTools(
  server: McpServer,
  getApiHelper: () => GraphApiHelper | null
) {
  // Create task tool
  server.tool(
    "create-task",
    "Create a task in Microsoft To Do. Requires Tasks.ReadWrite permission.",
    {
      userEmail: z.string().describe("Email address of the user on whose behalf to create the task"),
      title: z.string().describe("Task title"),
      body: z.object({
        content: z.string().describe("Task description"),
        contentType: z.enum(["text", "html"]).optional().default("text")
      }).optional().describe("Task body/description"),
      dueDateTime: z.object({
        dateTime: z.string().describe("Due date and time in ISO 8601 format"),
        timeZone: z.string().describe("Time zone")
      }).optional().describe("Task due date"),
      reminder: z.object({
        dateTime: z.string().describe("Reminder date and time in ISO 8601 format"),
        timeZone: z.string().describe("Time zone")
      }).optional().describe("Task reminder"),
      importance: z.enum(["low", "normal", "high"]).optional().default("normal"),
      categories: z.array(z.string()).optional().describe("Task categories/labels"),
      taskListId: z.string().optional().describe("ID of the task list (defaults to the default list)")
    },
    async ({ userEmail, title, body, dueDateTime, reminder, importance, categories, taskListId }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        // Build the task object
        const task: any = {
          title,
          importance
        };

        if (body) {
          task.body = body;
        }

        if (dueDateTime) {
          task.dueDateTime = dueDateTime;
        }

        if (reminder) {
          task.reminderDateTime = reminder;
        }

        if (categories && categories.length > 0) {
          task.categories = categories;
        }

        // Determine the path based on whether a task list ID is provided
        const path = taskListId 
          ? `/users/${userEmail}/todo/lists/${taskListId}/tasks`
          : `/users/${userEmail}/todo/lists/tasks/tasks`; // Default tasks list

        logger.info(`Creating task: ${title} on behalf of ${userEmail}`);

        const response = await apiHelper.callGraphApi({
          path,
          method: 'post',
          body: task
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const createdTask = response.data;
        let resultText = `Task created successfully!\n\n`;
        resultText += `Title: ${createdTask.title}\n`;
        resultText += `Status: ${createdTask.status}\n`;
        resultText += `Importance: ${createdTask.importance}\n`;
        if (createdTask.dueDateTime) {
          resultText += `Due: ${createdTask.dueDateTime.dateTime} (${createdTask.dueDateTime.timeZone})\n`;
        }
        if (createdTask.reminderDateTime) {
          resultText += `Reminder: ${createdTask.reminderDateTime.dateTime}\n`;
        }
        if (createdTask.categories && createdTask.categories.length > 0) {
          resultText += `Categories: ${createdTask.categories.join(', ')}\n`;
        }
        resultText += `\nTask ID: ${createdTask.id}`;

        return {
          content: [{ 
            type: "text" as const, 
            text: resultText
          }],
        };

      } catch (error: any) {
        logger.error("Error creating task:", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the Tasks.ReadWrite permission is granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error creating task: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );
}
