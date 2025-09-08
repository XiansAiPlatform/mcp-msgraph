import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { logger } from "../logger.js";
import { GraphApiHelper } from "../graph-api-helper.js";

// Helper schema for boolean parameters that accepts both boolean and string values
const zBooleanParam = () => z.union([
  z.boolean(),
  z.string().transform(str => str.toLowerCase() === 'true')
]);

export function registerSubscriptionTools(
  server: McpServer,
  getApiHelper: () => GraphApiHelper | null
) {
  server.tool(
    "create-subscription",
    "Create a subscription for change notifications on Microsoft Graph resources. Requires appropriate permissions for the resource being subscribed to.",
    {
      resource: z.string().describe("The resource path to subscribe to (e.g., /me/messages, /users/{user-id}/events, /groups/{group-id}/conversations)"),
      changeType: z.string().describe("Comma-separated list of change types to subscribe to (created, updated, deleted)"),
      notificationUrl: z.string().describe("The HTTPS endpoint URL where notifications will be sent"),
      expirationDateTime: z.string().optional().describe("Expiration date/time in ISO 8601 format (defaults to maximum allowed for the resource)"),
      clientState: z.string().describe("Client state value for security validation (max 255 chars) - required for security"),
      lifecycleNotificationUrl: z.string().optional().describe("Optional HTTPS endpoint URL for lifecycle notifications"),
      latestSupportedTlsVersion: z.enum(["v1_0", "v1_1", "v1_2", "v1_3"]).optional().describe("The latest TLS version supported by the notification endpoint"),
      encryptionCertificate: z.string().optional().describe("Base64 encoded certificate for encrypting resource data in notifications"),
      encryptionCertificateId: z.string().optional().describe("Custom ID for the encryption certificate"),
      includeResourceData: zBooleanParam().optional().describe("Whether to include resource data in notifications (requires encryption certificate)")
    },
    async (args) => {
      logger.info("Creating subscription", { resource: args.resource });
      const apiHelper = getApiHelper();
      if (!apiHelper) {
        throw new Error("Not authenticated. Please run authenticate-user first.");
      }

      try {
        // Calculate default expiration if not provided
        let expirationDateTime = args.expirationDateTime;
        if (!expirationDateTime) {
          // Default to 3 days from now (near maximum for most resources)
          const expiration = new Date();
          expiration.setDate(expiration.getDate() + 3);
          expirationDateTime = expiration.toISOString();
        }

        const subscriptionBody: any = {
          changeType: args.changeType,
          notificationUrl: args.notificationUrl,
          resource: args.resource,
          expirationDateTime: expirationDateTime,
          clientState: args.clientState,
        };

        // Add optional parameters
        if (args.lifecycleNotificationUrl) {
          subscriptionBody.lifecycleNotificationUrl = args.lifecycleNotificationUrl;
        }
        if (args.latestSupportedTlsVersion) {
          subscriptionBody.latestSupportedTlsVersion = args.latestSupportedTlsVersion;
        }
        if (args.encryptionCertificate) {
          subscriptionBody.encryptionCertificate = args.encryptionCertificate;
        }
        if (args.encryptionCertificateId) {
          subscriptionBody.encryptionCertificateId = args.encryptionCertificateId;
        }
        if (args.includeResourceData !== undefined) {
          subscriptionBody.includeResourceData = args.includeResourceData;
        }

        const response = await apiHelper.callGraphApi({
          path: "/subscriptions",
          method: "post",
          body: subscriptionBody,
        });

        if (response.error) {
          throw new Error(response.error);
        }

        return {
          content: [{ 
            type: "text" as const, 
            text: `Subscription created successfully.\nID: ${response.data.id}\nResource: ${response.data.resource}\nChange Type: ${response.data.changeType}\nNotification URL: ${response.data.notificationUrl}\nExpiration: ${response.data.expirationDateTime}\n\nFull response:\n${JSON.stringify(response.data, null, 2)}`
          }],
        };
      } catch (error: any) {
        logger.error("Error creating subscription", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the appropriate permissions are granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error creating subscription: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  server.tool(
    "list-subscriptions",
    "List all active subscriptions for the authenticated application",
    {},
    async () => {
      logger.info("Listing subscriptions");
      const apiHelper = getApiHelper();
      if (!apiHelper) {
        throw new Error("Not authenticated. Please run authenticate-user first.");
      }

      try {
        logger.info("Starting subscription list request");
        const response = await apiHelper.callGraphApi({
          path: "/subscriptions",
          method: "get",
          fetchAll: true,
        });

        logger.info("Received response from Graph API", { 
          hasError: !!response.error, 
          dataType: typeof response.data,
          dataKeys: response.data ? Object.keys(response.data) : null
        });

        if (response.error) {
          throw new Error(response.error);
        }

        // When fetchAll is true, response.data has structure: { '@odata.context': ..., value: [...] }
        // We need to extract the actual array from the value property
        let subscriptions;
        try {
          subscriptions = response.data?.value || response.data || [];
          logger.info("Extracted subscriptions", { 
            isArray: Array.isArray(subscriptions), 
            type: typeof subscriptions,
            length: Array.isArray(subscriptions) ? subscriptions.length : 'N/A'
          });
        } catch (extractError: any) {
          logger.error("Error extracting subscriptions from response", extractError);
          throw new Error(`Failed to extract subscriptions: ${extractError.message}`);
        }

        const count = Array.isArray(subscriptions) ? subscriptions.length : 0;

        let resultText = `Found ${count} active subscription(s)\n\n`;
        
        if (count > 0) {
          try {
            subscriptions.forEach((sub: any, index: number) => {
              resultText += `${index + 1}. ID: ${sub.id}\n`;
              resultText += `   Resource: ${sub.resource}\n`;
              resultText += `   Change Type: ${sub.changeType}\n`;
              resultText += `   Notification URL: ${sub.notificationUrl}\n`;
              resultText += `   Expiration: ${sub.expirationDateTime}\n`;
              resultText += `   Application ID: ${sub.applicationId || 'N/A'}\n`;
              resultText += `   Creator ID: ${sub.creatorId || 'N/A'}\n\n`;
            });
          } catch (forEachError: any) {
            logger.error("Error processing subscriptions in forEach", forEachError);
            resultText += `Error processing subscription data: ${forEachError.message}\n`;
            resultText += `Raw subscription data: ${JSON.stringify(subscriptions, null, 2)}\n`;
          }
        }
        
        return {
          content: [{ 
            type: "text" as const, 
            text: resultText
          }],
        };
      } catch (error: any) {
        logger.error("Error listing subscriptions", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the appropriate permissions are granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error listing subscriptions: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  server.tool(
    "get-subscription",
    "Get details of a specific subscription by ID",
    {
      subscriptionId: z.string().describe("The ID of the subscription to retrieve")
    },
    async (args) => {
      logger.info("Getting subscription", { subscriptionId: args.subscriptionId });
      const apiHelper = getApiHelper();
      if (!apiHelper) {
        throw new Error("Not authenticated. Please run authenticate-user first.");
      }

      try {
        const response = await apiHelper.callGraphApi({
          path: `/subscriptions/${args.subscriptionId}`,
          method: "get",
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const sub = response.data;
        const resultText = `Subscription Details:\n\nID: ${sub.id}\nResource: ${sub.resource}\nChange Type: ${sub.changeType}\nNotification URL: ${sub.notificationUrl}\nExpiration: ${sub.expirationDateTime}\nApplication ID: ${sub.applicationId || 'N/A'}\nCreator ID: ${sub.creatorId || 'N/A'}\nClient State: ${sub.clientState || 'N/A'}\nLifecycle Notification URL: ${sub.lifecycleNotificationUrl || 'N/A'}\nLatest Supported TLS Version: ${sub.latestSupportedTlsVersion || 'N/A'}\nInclude Resource Data: ${sub.includeResourceData || false}\n\nFull response:\n${JSON.stringify(sub, null, 2)}`;
        
        return {
          content: [{ 
            type: "text" as const, 
            text: resultText
          }],
        };
      } catch (error: any) {
        logger.error("Error getting subscription", error);
        const errorMessage = error.statusCode === 404 
          ? "Subscription not found"
          : error.statusCode === 403 
          ? "Permission denied. Ensure the appropriate permissions are granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error getting subscription: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  server.tool(
    "update-subscription",
    "Update/renew a subscription. Primarily used to extend the expiration date.",
    {
      subscriptionId: z.string().describe("The ID of the subscription to update"),
      expirationDateTime: z.string().describe("New expiration date/time in ISO 8601 format"),
      notificationUrl: z.string().optional().describe("New notification URL (if changing)"),
      lifecycleNotificationUrl: z.string().optional().describe("New lifecycle notification URL (if changing)"),
      latestSupportedTlsVersion: z.enum(["v1_0", "v1_1", "v1_2", "v1_3"]).optional().describe("Update the latest TLS version supported"),
      encryptionCertificate: z.string().optional().describe("New encryption certificate (base64 encoded)"),
      encryptionCertificateId: z.string().optional().describe("New encryption certificate ID"),
      includeResourceData: zBooleanParam().optional().describe("Update whether to include resource data")
    },
    async (args) => {
      logger.info("Updating subscription", { subscriptionId: args.subscriptionId });
      const apiHelper = getApiHelper();
      if (!apiHelper) {
        throw new Error("Not authenticated. Please run authenticate-user first.");
      }

      try {
        const updateBody: any = {
          expirationDateTime: args.expirationDateTime,
        };

        // Add optional parameters if provided
        if (args.notificationUrl) {
          updateBody.notificationUrl = args.notificationUrl;
        }
        if (args.lifecycleNotificationUrl) {
          updateBody.lifecycleNotificationUrl = args.lifecycleNotificationUrl;
        }
        if (args.latestSupportedTlsVersion) {
          updateBody.latestSupportedTlsVersion = args.latestSupportedTlsVersion;
        }
        if (args.encryptionCertificate) {
          updateBody.encryptionCertificate = args.encryptionCertificate;
        }
        if (args.encryptionCertificateId) {
          updateBody.encryptionCertificateId = args.encryptionCertificateId;
        }
        if (args.includeResourceData !== undefined) {
          updateBody.includeResourceData = args.includeResourceData;
        }

        const response = await apiHelper.callGraphApi({
          path: `/subscriptions/${args.subscriptionId}`,
          method: "patch",
          body: updateBody,
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const sub = response.data;
        const resultText = `Subscription updated successfully.\n\nUpdated Details:\nID: ${sub.id}\nResource: ${sub.resource}\nChange Type: ${sub.changeType}\nNotification URL: ${sub.notificationUrl}\nExpiration: ${sub.expirationDateTime}\nApplication ID: ${sub.applicationId || 'N/A'}\nCreator ID: ${sub.creatorId || 'N/A'}\n\nFull response:\n${JSON.stringify(sub, null, 2)}`;
        
        return {
          content: [{ 
            type: "text" as const, 
            text: resultText
          }],
        };
      } catch (error: any) {
        logger.error("Error updating subscription", error);
        const errorMessage = error.statusCode === 404 
          ? "Subscription not found"
          : error.statusCode === 403 
          ? "Permission denied. Ensure the appropriate permissions are granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error updating subscription: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  server.tool(
    "delete-subscription",
    "Delete a subscription to stop receiving change notifications",
    {
      subscriptionId: z.string().describe("The ID of the subscription to delete")
    },
    async (args) => {
      logger.info("Deleting subscription", { subscriptionId: args.subscriptionId });
      const apiHelper = getApiHelper();
      if (!apiHelper) {
        throw new Error("Not authenticated. Please run authenticate-user first.");
      }

      try {
        const response = await apiHelper.callGraphApi({
          path: `/subscriptions/${args.subscriptionId}`,
          method: "delete",
        });

        if (response.error) {
          throw new Error(response.error);
        }

        return {
          content: [{ 
            type: "text" as const, 
            text: `Subscription ${args.subscriptionId} deleted successfully` 
          }],
        };
      } catch (error: any) {
        logger.error("Error deleting subscription", error);
        const errorMessage = error.statusCode === 404 
          ? "Subscription not found"
          : error.statusCode === 403 
          ? "Permission denied. Ensure the appropriate permissions are granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error deleting subscription: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  server.tool(
    "reauthorize-subscription",
    "Reauthorize and renew a subscription by updating its expiration date. This performs a PATCH operation to both reauthorize and extend the subscription in one call.",
    {
      subscriptionId: z.string().optional().describe("The ID of the subscription to reauthorize. Either provide this or lifecycleNotificationResource."),
      lifecycleNotificationResource: z.string().optional().describe("The lifecycle notification resource URL that contains the subscription ID and tenant ID"),
      expirationDateTime: z.string().optional().describe("New expiration date/time in ISO 8601 format. If not provided, defaults to 3 days from now."),
      tenantId: z.string().optional().describe("The tenant ID from the lifecycle notification"),
      clientState: z.string().optional().describe("The client state value to validate in the lifecycle notification")
    },
    async (args) => {
      logger.info("Reauthorizing subscription", { 
        subscriptionId: args.subscriptionId, 
        resource: args.lifecycleNotificationResource 
      });
      const apiHelper = getApiHelper();
      if (!apiHelper) {
        throw new Error("Not authenticated. Please run authenticate-user first.");
      }

      try {
        let subscriptionId: string;
        
        if (args.subscriptionId) {
          subscriptionId = args.subscriptionId;
        } else if (args.lifecycleNotificationResource) {
          // Extract subscription ID from the lifecycle notification resource URL
          // Format: https://graph.microsoft.com/v1.0/subscriptions/{subscription-id}?tenantId={tenant-id}
          const subscriptionIdMatch = args.lifecycleNotificationResource.match(/\/subscriptions\/([^?]+)/);
          if (!subscriptionIdMatch) {
            throw new Error("Invalid lifecycle notification resource URL");
          }
          subscriptionId = subscriptionIdMatch[1];
        } else {
          throw new Error("Either subscriptionId or lifecycleNotificationResource must be provided");
        }

        // Calculate expiration date if not provided
        let expirationDateTime = args.expirationDateTime;
        if (!expirationDateTime) {
          // Default to 3 days from now
          const expiration = new Date();
          expiration.setDate(expiration.getDate() + 3);
          expirationDateTime = expiration.toISOString();
        }

        // Perform PATCH operation to reauthorize and renew the subscription
        const updateBody = {
          expirationDateTime: expirationDateTime
        };

        const response = await apiHelper.callGraphApi({
          path: `/subscriptions/${subscriptionId}`,
          method: "patch",
          body: updateBody,
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const sub = response.data;
        const resultText = `Subscription ${subscriptionId} reauthorized and renewed successfully.\n\nUpdated Details:\nID: ${sub.id}\nResource: ${sub.resource}\nChange Type: ${sub.changeType}\nNotification URL: ${sub.notificationUrl}\nNew Expiration: ${sub.expirationDateTime}\nApplication ID: ${sub.applicationId || 'N/A'}\nCreator ID: ${sub.creatorId || 'N/A'}\n\nFull response:\n${JSON.stringify(sub, null, 2)}`;
        
        return {
          content: [{ 
            type: "text" as const, 
            text: resultText
          }],
        };
      } catch (error: any) {
        logger.error("Error reauthorizing subscription", error);
        const errorMessage = error.statusCode === 404 
          ? "Subscription not found"
          : error.statusCode === 403 
          ? "Permission denied. Ensure the appropriate permissions are granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error reauthorizing subscription: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );
}
