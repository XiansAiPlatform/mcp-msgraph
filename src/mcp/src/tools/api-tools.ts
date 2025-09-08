import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { logger } from "../logger.js";
import { AuthMode } from "../auth.js";
import { GraphApiHelper } from "../graph-api-helper.js";
import { getDefaultGraphApiVersion } from "../constants.js";

// Helper schema for boolean parameters that accepts both boolean and string values
const zBooleanParam = () => z.union([
  z.boolean(),
  z.string().transform(str => str.toLowerCase() === 'true')
]);

export function registerApiTools(
  server: McpServer,
  getApiHelper: () => GraphApiHelper | null,
  useGraphBeta: boolean,
  defaultGraphApiVersion: string
) {
  server.tool(
    "call-graph-api",
    "A versatile tool to interact with Microsoft APIs including Microsoft Graph (Entra) and Azure Resource Management. IMPORTANT: For Graph API GET requests using advanced query parameters ($filter, $count, $search, $orderby), you are ADVISED to set 'consistencyLevel: \"eventual\"'.",
    {
      apiType: z.enum(["graph", "azure"]).describe("Type of Microsoft API to query. Options: 'graph' for Microsoft Graph (Entra) or 'azure' for Azure Resource Management."),
      path: z.string().describe("The Azure or Graph API URL path to call (e.g. '/users', '/groups', '/subscriptions')"),
      method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method to use"),
      apiVersion: z.string().optional().describe("Azure Resource Management API version (required for apiType Azure)"),
      subscriptionId: z.string().optional().describe("Azure Subscription ID (for Azure Resource Management)."),
      queryParams: z.record(z.string()).optional().describe("Query parameters for the request"),
      body: z.union([z.record(z.string(), z.any()), z.string(), z.array(z.record(z.string(), z.any()))]).optional().describe("The request body (for POST, PUT, PATCH) - can be an object, string, or array of objects"),
      graphApiVersion: z.enum(["v1.0", "beta"]).optional().default(defaultGraphApiVersion as "v1.0" | "beta").describe(`Microsoft Graph API version to use (default: ${defaultGraphApiVersion})`),
      fetchAll: zBooleanParam().optional().default(false).describe("Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false."),
      consistencyLevel: z.string().optional().describe("Graph API ConsistencyLevel header. ADVISED to be set to 'eventual' for Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby)."),
    },
    async ({
      apiType,
      path,
      method,
      apiVersion,
      subscriptionId,
      queryParams,
      body,
      graphApiVersion,
      fetchAll,
      consistencyLevel
    }: {
      apiType: "graph" | "azure";
      path: string;
      method: "get" | "post" | "put" | "patch" | "delete";
      apiVersion?: string;
      subscriptionId?: string;
      queryParams?: Record<string, string>;
      body?: any;
      graphApiVersion: "v1.0" | "beta";
      fetchAll: boolean;
      consistencyLevel?: string;
    }) => {
      // Override graphApiVersion if USE_GRAPH_BETA is explicitly set to false
      const effectiveGraphApiVersion = !useGraphBeta ? "v1.0" : graphApiVersion;
      
      logger.info(`Executing tool with params: apiType=${apiType}, path=${path}, method=${method}, graphApiVersion=${effectiveGraphApiVersion}, fetchAll=${fetchAll}, consistencyLevel=${consistencyLevel}`);
      let determinedUrl: string | undefined;

      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        let responseData: any;

        // --- Microsoft Graph Logic ---
        if (apiType === 'graph') {
          determinedUrl = `https://graph.microsoft.com/${effectiveGraphApiVersion}`; // For error reporting

          const response = await apiHelper.callGraphApi({
            path,
            method,
            queryParams,
            body,
            graphApiVersion: effectiveGraphApiVersion,
            fetchAll,
            consistencyLevel
          });

          if (response.error) {
            throw new Error(response.error);
          }

          responseData = response.data;
        }
        // --- Azure Resource Management Logic (using direct fetch) ---
        else { // apiType === 'azure'
          if (!apiVersion) {
            throw new Error("API version is required for Azure Resource Management queries");
          }
          determinedUrl = "https://management.azure.com"; // For error reporting

          const response = await apiHelper.callAzureApi({
            path,
            method,
            apiVersion,
            subscriptionId,
            queryParams,
            body,
            fetchAll
          });

          if (response.error) {
            throw new Error(response.error);
          }

          responseData = response.data;
        }

        // --- Format and Return Result ---
        // For all requests, format as text
        let resultText = `Result for ${apiType} API (${apiType === 'graph' ? effectiveGraphApiVersion : apiVersion}) - ${method} ${path}:\n\n`;
        resultText += JSON.stringify(responseData, null, 2); // responseData already contains the correct structure for fetchAll Graph case

        // Add pagination note if applicable (only for single page GET)
        if (!fetchAll && method === 'get' && apiHelper) {
           if (apiHelper.hasMorePages(responseData, apiType)) {
               resultText += `\n\nNote: More results are available. To retrieve all pages, add the parameter 'fetchAll: true' to your request.`;
           }
        }

        return {
          content: [{ type: "text" as const, text: resultText }],
        };

      } catch (error: any) {
        logger.error(`Error in tool (apiType: ${apiType}, path: ${path}, method: ${method}):`, error); // Added more context to error log
        // Try to determine the base URL even in case of error
        if (!determinedUrl) {
           determinedUrl = apiType === 'graph'
             ? `https://graph.microsoft.com/${effectiveGraphApiVersion}`
             : "https://management.azure.com";
        }
        // Include error body if available from Graph SDK error
        const errorBody = error.body ? (typeof error.body === 'string' ? error.body : JSON.stringify(error.body)) : 'N/A';
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              error: error instanceof Error ? error.message : String(error),
              statusCode: error.statusCode || 'N/A', // Include status code if available from SDK error
              errorBody: errorBody,
              attemptedBaseUrl: determinedUrl
            }),
          }],
          isError: true
        };
      }
    },
  );
}
