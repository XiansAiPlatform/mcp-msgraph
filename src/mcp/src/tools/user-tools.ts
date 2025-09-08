import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { logger } from "../logger.js";
import { GraphApiHelper } from "../graph-api-helper.js";

// Helper schema for boolean parameters that accepts both boolean and string values
const zBooleanParam = () => z.union([
  z.boolean(),
  z.string().transform(str => str.toLowerCase() === 'true')
]);

export function registerUserTools(
  server: McpServer,
  getApiHelper: () => GraphApiHelper | null
) {
  // List users tool
  server.tool(
    "list-users",
    "List users in the organization using Microsoft Graph API. Requires User.Read.All or Directory.Read.All permission.",
    {
      filter: z.string().optional().describe("OData filter query (e.g., \"startswith(displayName,'John')\")"),
      select: z.array(z.string()).optional().describe("Properties to include in the response"),
      top: z.union([z.number(), z.string().transform(str => parseInt(str, 10))]).optional().default(10).describe("Number of users to return (default: 10)"),
      orderBy: z.string().optional().describe("Property to sort by (e.g., 'displayName')"),
      fetchAll: zBooleanParam().optional().default(false).describe("Whether to fetch all users (ignores 'top' parameter)")
    },
    async ({ filter, select, top, orderBy, fetchAll }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        const queryParams: Record<string, string> = {};
        
        // Add query parameters
        if (filter) {
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

        // For advanced queries, we might need consistency level
        const needsConsistencyLevel = filter?.includes('endsWith') || orderBy || filter?.includes('$count');

        logger.info(`Listing users with parameters: ${JSON.stringify(queryParams)}`);

        const response = await apiHelper.callGraphApi({
          path: '/users',
          method: 'get',
          queryParams,
          fetchAll,
          consistencyLevel: needsConsistencyLevel ? 'eventual' : undefined
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const users = response.data.value || [];
        let resultText = `Found ${users.length} user(s):\n\n`;
        
        users.forEach((user: any, index: number) => {
          resultText += `${index + 1}. ${user.displayName || 'No Name'}\n`;
          resultText += `   Email: ${user.mail || user.userPrincipalName}\n`;
          resultText += `   ID: ${user.id}\n`;
          if (user.jobTitle) {
            resultText += `   Title: ${user.jobTitle}\n`;
          }
          if (user.department) {
            resultText += `   Department: ${user.department}\n`;
          }
          resultText += '\n';
        });

        if (!fetchAll && apiHelper.hasMorePages(response.data, 'graph')) {
          resultText += `Note: More users are available. Use 'fetchAll: true' to retrieve all users.`;
        }

        return {
          content: [{ 
            type: "text" as const, 
            text: resultText
          }],
        };

      } catch (error: any) {
        logger.error("Error listing users:", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the User.Read.All or Directory.Read.All permission is granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error listing users: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  // Get user profile tool
  server.tool(
    "get-user-profile",
    "Get user profile information using Microsoft Graph API. Specify a user ID/UPN/email.",
    {
      userId: z.string().describe("User ID, UPN (user@domain.com), or email address of the user to get profile for"),
      select: z.array(z.string()).optional().describe("Specific properties to retrieve")
    },
    async ({ userId, select }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        const queryParams: Record<string, string> = {};
        if (select && select.length > 0) {
          queryParams.$select = select.join(',');
        }

        logger.info(`Getting user profile for: ${userId}`);

        const response = await apiHelper.callGraphApi({
          path: `/users/${userId}`,
          method: 'get',
          queryParams
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const user = response.data;
        let resultText = `User Profile:\n\n`;
        resultText += `Display Name: ${user.displayName || 'N/A'}\n`;
        resultText += `Email: ${user.mail || user.userPrincipalName}\n`;
        resultText += `ID: ${user.id}\n`;
        if (user.jobTitle) resultText += `Job Title: ${user.jobTitle}\n`;
        if (user.department) resultText += `Department: ${user.department}\n`;
        if (user.officeLocation) resultText += `Office: ${user.officeLocation}\n`;
        if (user.mobilePhone) resultText += `Mobile: ${user.mobilePhone}\n`;
        if (user.businessPhones && user.businessPhones.length > 0) {
          resultText += `Business Phone(s): ${user.businessPhones.join(', ')}\n`;
        }
        if (user.manager) resultText += `Manager ID: ${user.manager.id}\n`;
        
        resultText += `\nFull Response:\n${JSON.stringify(user, null, 2)}`;

        return {
          content: [{ 
            type: "text" as const, 
            text: resultText
          }],
        };

      } catch (error: any) {
        logger.error("Error getting user profile:", error);
        const errorMessage = error.statusCode === 404
          ? "User not found. Please check the user ID or UPN."
          : error.statusCode === 403 
          ? "Permission denied. Ensure appropriate user read permissions are granted."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error getting user profile: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );
}
