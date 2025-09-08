#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { Client } from "@microsoft/microsoft-graph-client";
import fetch from 'isomorphic-fetch'; // Required polyfill for Graph client
import { logger } from "./logger.js";
import { AuthManager, AuthConfig, AuthMode } from "./auth.js";
import { LokkaClientId, LokkaDefaultTenantId, LokkaDefaultRedirectUri, getDefaultGraphApiVersion } from "./constants.js";
import { GraphApiHelper } from "./graph-api-helper.js";
import { registerAllTools } from "./tools/index.js";

// Set up global fetch for the Microsoft Graph client
(global as any).fetch = fetch;

// Create server instance
const server = new McpServer({
  name: "MCP Microsoft Graph",
  version: "0.2.0", // Updated version for token-based auth support
});

logger.info("Starting Multi-Microsoft API MCP Server");

// Initialize authentication and clients
let authManager: AuthManager | null = null;
let graphClient: Client | null = null;
let apiHelper: GraphApiHelper | null = null;

// Check USE_GRAPH_BETA environment variable
const useGraphBeta = process.env.USE_GRAPH_BETA !== 'false'; // Default to true unless explicitly set to 'false'
const defaultGraphApiVersion = getDefaultGraphApiVersion();

logger.info(`Graph API default version: ${defaultGraphApiVersion} (USE_GRAPH_BETA=${process.env.USE_GRAPH_BETA || 'undefined'})`);

// Register all tools with the server
registerAllTools({
  server,
  getAuthManager: () => authManager,
  setAuthManager: (manager) => { authManager = manager; },
  getGraphClient: () => graphClient,
  setGraphClient: (client) => { graphClient = client; },
  getApiHelper: () => apiHelper,
  setApiHelper: (helper) => { apiHelper = helper; },
  useGraphBeta,
  defaultGraphApiVersion
});

// Start the server with stdio transport
async function main() {
  // Determine authentication mode based on environment variables
  const useCertificate = process.env.USE_CERTIFICATE === 'true';
  const useInteractive = process.env.USE_INTERACTIVE === 'true';
  const useClientToken = process.env.USE_CLIENT_TOKEN === 'true';
  const initialAccessToken = process.env.ACCESS_TOKEN;
  
  let authMode: AuthMode;
  
  // Ensure only one authentication mode is enabled at a time
  const enabledModes = [
    useClientToken,
    useInteractive,
    useCertificate
  ].filter(Boolean);

  if (enabledModes.length > 1) {
    throw new Error(
      "Multiple authentication modes enabled. Please enable only one of USE_CLIENT_TOKEN, USE_INTERACTIVE, or USE_CERTIFICATE."
    );
  }

  if (useClientToken) {
    authMode = AuthMode.ClientProvidedToken;
    if (!initialAccessToken) {
      logger.info("Client token mode enabled but no initial token provided. Token must be set via set-access-token tool.");
    }
  } else if (useInteractive) {
    authMode = AuthMode.Interactive;
  } else if (useCertificate) {
    authMode = AuthMode.Certificate;
  } else {
    // Check if we have client credentials environment variables
    const hasClientCredentials = process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET;
    
    if (hasClientCredentials) {
      authMode = AuthMode.ClientCredentials;
    } else {
      // Default to interactive mode for better user experience
      authMode = AuthMode.Interactive;
      logger.info("No authentication mode specified and no client credentials found. Defaulting to interactive mode.");
    }
  }

  logger.info(`Starting with authentication mode: ${authMode}`);

  // Get tenant ID and client ID with defaults only for interactive mode
  let tenantId: string | undefined;
  let clientId: string | undefined;
  
  if (authMode === AuthMode.Interactive) {
    // Interactive mode can use defaults
    tenantId = process.env.TENANT_ID || LokkaDefaultTenantId;
    clientId = process.env.CLIENT_ID || LokkaClientId;
    logger.info(`Interactive mode using tenant ID: ${tenantId}, client ID: ${clientId}`);
  } else {
    // All other modes require explicit values from environment variables
    tenantId = process.env.TENANT_ID;
    clientId = process.env.CLIENT_ID;
  }

  const clientSecret = process.env.CLIENT_SECRET;
  const certificatePath = process.env.CERTIFICATE_PATH;
  const certificatePassword = process.env.CERTIFICATE_PASSWORD; // optional

  // Validate required configuration
  if (authMode === AuthMode.ClientCredentials) {
    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Client credentials mode requires explicit TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables");
    }
  } else if (authMode === AuthMode.Certificate) {
    if (!tenantId || !clientId || !certificatePath) {
      throw new Error("Certificate mode requires explicit TENANT_ID, CLIENT_ID, and CERTIFICATE_PATH environment variables");
    }
  }
  // Note: Client token mode can start without a token and receive it later

  const authConfig: AuthConfig = {
    mode: authMode,
    tenantId,
    clientId,
    clientSecret,
    accessToken: initialAccessToken,
    redirectUri: process.env.REDIRECT_URI,
    certificatePath,
    certificatePassword
  };

  authManager = new AuthManager(authConfig);
  
  // Only initialize if we have required config (for client token mode, we can start without a token)
  if (authMode !== AuthMode.ClientProvidedToken || initialAccessToken) {
    await authManager.initialize();
    
    // Initialize Graph Client
    const authProvider = authManager.getGraphAuthProvider();
    graphClient = Client.initWithMiddleware({
      authProvider: authProvider,
    });
    
    // Initialize API helper
    apiHelper = new GraphApiHelper(graphClient, authManager, useGraphBeta, defaultGraphApiVersion);
    
    logger.info(`Authentication initialized successfully using ${authMode} mode`);
  } else {
    logger.info("Started in client token mode. Use set-access-token tool to provide authentication token.");
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  logger.error("Fatal error in main()", error);
  process.exit(1);
});
