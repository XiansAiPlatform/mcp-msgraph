import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { logger } from "../logger.js";
import { GraphApiHelper } from "../graph-api-helper.js";

// Helper schema for boolean parameters that accepts both boolean and string values
const zBooleanParam = () => z.union([
  z.boolean(),
  z.string().transform(str => str.toLowerCase() === 'true')
]);

export function registerGroupTools(
  server: McpServer,
  getApiHelper: () => GraphApiHelper | null
) {
  // List groups tool
  server.tool(
    "list-groups",
    "List groups in the organization. Requires Group.Read.All or Directory.Read.All permission.",
    {
      filter: z.string().optional().describe("OData filter query (e.g., \"startswith(displayName,'Sales')\")"),
      select: z.array(z.string()).optional().describe("Properties to include in the response"),
      top: z.union([z.number(), z.string().transform(str => parseInt(str, 10))]).optional().default(10).describe("Number of groups to return"),
      orderBy: z.string().optional().describe("Property to sort by"),
      groupTypes: z.array(z.enum(["Unified", "DynamicMembership", "Security"])).optional().describe("Filter by group types"),
      fetchAll: zBooleanParam().optional().default(false).describe("Whether to fetch all groups")
    },
    async ({ filter, select, top, orderBy, groupTypes, fetchAll }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        const queryParams: Record<string, string> = {};
        
        // Build filter for group types if provided
        let typeFilter = "";
        if (groupTypes && groupTypes.length > 0) {
          const typeConditions = groupTypes.map(type => {
            if (type === "Unified") {
              return "groupTypes/any(c:c eq 'Unified')";
            } else if (type === "DynamicMembership") {
              return "groupTypes/any(c:c eq 'DynamicMembership')";
            } else {
              return "securityEnabled eq true and mailEnabled eq false";
            }
          });
          typeFilter = typeConditions.join(' or ');
        }

        // Combine with any additional filter
        if (typeFilter && filter) {
          queryParams.$filter = `(${typeFilter}) and (${filter})`;
        } else if (typeFilter) {
          queryParams.$filter = typeFilter;
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

        // For advanced queries, we might need consistency level
        const needsConsistencyLevel = filter?.includes('endsWith') || orderBy || filter?.includes('$count');

        logger.info(`Listing groups with parameters: ${JSON.stringify(queryParams)}`);

        const response = await apiHelper.callGraphApi({
          path: '/groups',
          method: 'get',
          queryParams,
          fetchAll,
          consistencyLevel: needsConsistencyLevel ? 'eventual' : undefined
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const groups = response.data.value || [];
        let resultText = `Found ${groups.length} group(s):\n\n`;
        
        groups.forEach((group: any, index: number) => {
          resultText += `${index + 1}. ${group.displayName}\n`;
          resultText += `   ID: ${group.id}\n`;
          if (group.description) {
            resultText += `   Description: ${group.description}\n`;
          }
          if (group.mail) {
            resultText += `   Email: ${group.mail}\n`;
          }
          
          // Determine group type
          let groupType = "Security Group";
          if (group.groupTypes && group.groupTypes.includes("Unified")) {
            groupType = "Microsoft 365 Group";
          } else if (group.groupTypes && group.groupTypes.includes("DynamicMembership")) {
            groupType = "Dynamic Group";
          }
          resultText += `   Type: ${groupType}\n`;
          
          if (group.visibility) {
            resultText += `   Visibility: ${group.visibility}\n`;
          }
          resultText += '\n';
        });

        if (!fetchAll && apiHelper.hasMorePages(response.data, 'graph')) {
          resultText += `Note: More groups are available. Use 'fetchAll: true' to retrieve all groups.`;
        }

        return {
          content: [{ 
            type: "text" as const, 
            text: resultText
          }],
        };

      } catch (error: any) {
        logger.error("Error listing groups:", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the Group.Read.All or Directory.Read.All permission is granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error listing groups: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );
}
