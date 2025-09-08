import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { Client } from "@microsoft/microsoft-graph-client";
import { logger } from "../logger.js";
import { AuthManager, AuthConfig, AuthMode } from "../auth.js";
import { LokkaClientId, LokkaDefaultTenantId, LokkaDefaultRedirectUri, getDefaultGraphApiVersion } from "../constants.js";
import { GraphApiHelper } from "../graph-api-helper.js";

interface AuthToolsContext {
  getAuthManager: () => AuthManager | null;
  setAuthManager: (authManager: AuthManager | null) => void;
  getGraphClient: () => Client | null;
  setGraphClient: (client: Client | null) => void;
  getApiHelper: () => GraphApiHelper | null;
  setApiHelper: (helper: GraphApiHelper | null) => void;
  useGraphBeta: boolean;
  defaultGraphApiVersion: string;
}

export function registerAuthTools(
  server: McpServer,
  context: AuthToolsContext
) {
  // Set access token tool
  server.tool(
    "set-access-token",
    "Set or update the access token for Microsoft Graph authentication. Use this when the MCP Client has obtained a fresh token through interactive authentication.",
    {
      accessToken: z.string().describe("The access token obtained from Microsoft Graph authentication"),
      expiresOn: z.string().optional().describe("Token expiration time in ISO format (optional, defaults to 1 hour from now)")
    },
    async ({ accessToken, expiresOn }) => {
      try {
        const expirationDate = expiresOn ? new Date(expiresOn) : undefined;
        const authManager = context.getAuthManager();
        
        if (authManager?.getAuthMode() === AuthMode.ClientProvidedToken) {
          authManager.updateAccessToken(accessToken, expirationDate);
          
          // Reinitialize the Graph client with the new token
          const authProvider = authManager.getGraphAuthProvider();
          const graphClient = Client.initWithMiddleware({
            authProvider: authProvider,
          });
          context.setGraphClient(graphClient);
          
          // Update API helper with new client
          const apiHelper = new GraphApiHelper(
            graphClient, 
            authManager, 
            context.useGraphBeta, 
            context.defaultGraphApiVersion as "v1.0" | "beta"
          );
          context.setApiHelper(apiHelper);
          
          return {
            content: [{ 
              type: "text" as const, 
              text: "Access token updated successfully. You can now make Microsoft Graph requests on behalf of the authenticated user." 
            }],
          };
        } else {
          return {
            content: [{ 
              type: "text" as const, 
              text: "Error: MCP Server is not configured for client-provided token authentication. Set USE_CLIENT_TOKEN=true in environment variables." 
            }],
            isError: true
          };
        }
      } catch (error: any) {
        logger.error("Error setting access token:", error);
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error setting access token: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Get auth status tool
  server.tool(
    "get-auth-status",
    "Check the current authentication status and mode of the MCP Server and also returns the current graph permission scopes of the access token for the current session.",
    {},
    async () => {
      try {
        const authManager = context.getAuthManager();
        const authMode = authManager?.getAuthMode() || "Not initialized";
        const isReady = authManager !== null;
        const tokenStatus = authManager ? await authManager.getTokenStatus() : { isExpired: false };
        
        return {
          content: [{ 
            type: "text" as const, 
            text: JSON.stringify({
              authMode,
              isReady,
              supportsTokenUpdates: authMode === AuthMode.ClientProvidedToken,
              tokenStatus: tokenStatus,
              timestamp: new Date().toISOString()
            }, null, 2)
          }],
        };
      } catch (error: any) {
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error checking auth status: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Add graph permission tool
  server.tool(
    "add-graph-permission",
    "Request additional Microsoft Graph permission scopes by performing a fresh interactive sign-in. This tool only works in interactive authentication mode and should be used if any Graph API call returns permissions related errors.",
    {
      scopes: z.array(z.string()).describe("Array of Microsoft Graph permission scopes to request (e.g., ['User.Read', 'Mail.ReadWrite', 'Directory.Read.All'])")
    },
    async ({ scopes }) => {
      try {
        const authManager = context.getAuthManager();
        
        // Check if we're in interactive mode
        if (!authManager || authManager.getAuthMode() !== AuthMode.Interactive) {
          const currentMode = authManager?.getAuthMode() || "Not initialized";
          const clientId = process.env.CLIENT_ID;
          
          let errorMessage = `Error: add-graph-permission tool is only available in interactive authentication mode. Current mode: ${currentMode}.\n\n`;
          
          if (currentMode === AuthMode.ClientCredentials) {
            errorMessage += `📋 To add permissions in Client Credentials mode:\n`;
            errorMessage += `1. Open the Microsoft Entra admin center (https://entra.microsoft.com)\n`;
            errorMessage += `2. Navigate to Applications > App registrations\n`;
            errorMessage += `3. Find your application${clientId ? ` (Client ID: ${clientId})` : ''}\n`;
            errorMessage += `4. Go to API permissions\n`;
            errorMessage += `5. Click "Add a permission" and select Microsoft Graph\n`;
            errorMessage += `6. Choose "Application permissions" and add the required scopes:\n`;
            errorMessage += `   ${scopes.map(scope => `• ${scope}`).join('\n   ')}\n`;
            errorMessage += `7. Click "Grant admin consent" to approve the permissions\n`;
            errorMessage += `8. Restart the MCP server to use the new permissions`;
          } else if (currentMode === AuthMode.ClientProvidedToken) {
            errorMessage += `📋 To add permissions in Client Provided Token mode:\n`;
            errorMessage += `1. Obtain a new access token that includes the required scopes:\n`;
            errorMessage += `   ${scopes.map(scope => `• ${scope}`).join('\n   ')}\n`;
            errorMessage += `2. When obtaining the token, ensure these scopes are included in the consent prompt\n`;
            errorMessage += `3. Use the set-access-token tool to update the server with the new token\n`;
            errorMessage += `4. The new token will include the additional permissions`;
          } else {
            errorMessage += `To use interactive permission requests, set USE_INTERACTIVE=true in environment variables and restart the server.`;
          }
          
          return {
            content: [{ 
              type: "text" as const, 
              text: errorMessage
            }],
            isError: true
          };
        }

        // Validate scopes array
        if (!scopes || scopes.length === 0) {
          return {
            content: [{ 
              type: "text" as const, 
              text: "Error: At least one permission scope must be specified." 
            }],
            isError: true
          };
        }

        // Validate scope format (basic validation)
        const invalidScopes = scopes.filter(scope => !scope.includes('.') || scope.trim() !== scope);
        if (invalidScopes.length > 0) {
          return {
            content: [{ 
              type: "text" as const, 
              text: `Error: Invalid scope format detected: ${invalidScopes.join(', ')}. Scopes should be in format like 'User.Read' or 'Mail.ReadWrite'.` 
            }],
            isError: true
          };
        }

        logger.info(`Requesting additional Graph permissions: ${scopes.join(', ')}`);

        // Get current configuration with defaults for interactive auth
        const tenantId = process.env.TENANT_ID || LokkaDefaultTenantId;
        const clientId = process.env.CLIENT_ID || LokkaClientId;
        const redirectUri = process.env.REDIRECT_URI || LokkaDefaultRedirectUri;

        logger.info(`Using tenant ID: ${tenantId}, client ID: ${clientId} for interactive authentication`);

        // Create a new interactive credential with the requested scopes
        const { InteractiveBrowserCredential, DeviceCodeCredential } = await import("@azure/identity");
        
        // Clear any existing auth manager to force fresh authentication
        context.setAuthManager(null);
        context.setGraphClient(null);
        
        // Request token with the new scopes - this will trigger interactive authentication
        const scopeString = scopes.map(scope => `https://graph.microsoft.com/${scope}`).join(' ');
        logger.info(`Requesting fresh token with scopes: ${scopeString}`);
        
        console.log(`\n🔐 Requesting Additional Graph Permissions:`);
        console.log(`Scopes: ${scopes.join(', ')}`);
        console.log(`You will be prompted to sign in to grant these permissions.\n`);

        let newCredential;
        let tokenResponse;
        
        try {
          // Try Interactive Browser first - create fresh instance each time
          newCredential = new InteractiveBrowserCredential({
            tenantId: tenantId,
            clientId: clientId,
            redirectUri: redirectUri,
          });
          
          // Request token immediately after creating credential
          tokenResponse = await newCredential.getToken(scopeString);
          
        } catch (error) {
          // Fallback to Device Code flow
          logger.info("Interactive browser failed, falling back to device code flow");
          newCredential = new DeviceCodeCredential({
            tenantId: tenantId,
            clientId: clientId,
            userPromptCallback: (info) => {
              console.log(`\n🔐 Additional Permissions Required:`);
              console.log(`Please visit: ${info.verificationUri}`);
              console.log(`And enter code: ${info.userCode}`);
              console.log(`Requested scopes: ${scopes.join(', ')}\n`);
              return Promise.resolve();
            },
          });
          
          // Request token with device code credential
          tokenResponse = await newCredential.getToken(scopeString);
        }

        if (!tokenResponse) {
          return {
            content: [{ 
              type: "text" as const, 
              text: "Error: Failed to acquire access token with the requested scopes. Please check your permissions and try again." 
            }],
            isError: true
          };
        }

        // Create a completely new auth manager instance with the updated credential
        const authConfig: AuthConfig = {
          mode: AuthMode.Interactive,
          tenantId,
          clientId,
          redirectUri
        };

        // Create a new auth manager instance
        const newAuthManager = new AuthManager(authConfig);
        
        // Manually set the credential to our new one with the additional scopes
        (newAuthManager as any).credential = newCredential;

        // DO NOT call initialize() as it might interfere with our fresh token
        // Instead, directly create the Graph client with the new credential
        const authProvider = newAuthManager.getGraphAuthProvider();
        const graphClient = Client.initWithMiddleware({
          authProvider: authProvider,
        });
        
        // Update context with new instances
        context.setAuthManager(newAuthManager);
        context.setGraphClient(graphClient);
        
        // Update API helper with new client
        const apiHelper = new GraphApiHelper(
          graphClient, 
          newAuthManager, 
          context.useGraphBeta, 
          context.defaultGraphApiVersion as "v1.0" | "beta"
        );
        context.setApiHelper(apiHelper);

        // Get the token status to show the new scopes
        const tokenStatus = await newAuthManager.getTokenStatus();

        logger.info(`Successfully acquired fresh token with additional scopes: ${scopes.join(', ')}`);

        return {
          content: [{ 
            type: "text" as const, 
            text: JSON.stringify({
              message: "Successfully acquired additional Microsoft Graph permissions with fresh authentication",
              requestedScopes: scopes,
              tokenStatus: tokenStatus,
              note: "A fresh sign-in was performed to ensure the new permissions are properly granted",
              timestamp: new Date().toISOString()
            }, null, 2)
          }],
        };

      } catch (error: any) {
        logger.error("Error requesting additional Graph permissions:", error);
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error requesting additional permissions: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );
}
