import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { Client } from "@microsoft/microsoft-graph-client";
import { AuthManager } from "../auth.js";
import { GraphApiHelper } from "../graph-api-helper.js";

import { registerApiTools } from "./api-tools.js";
import { registerAuthTools } from "./auth-tools.js";
import { registerMailTools } from "./mail-tools.js";
import { registerCalendarTools } from "./calendar-tools.js";
import { registerUserTools } from "./user-tools.js";
import { registerTaskTools } from "./task-tools.js";
import { registerGroupTools } from "./group-tools.js";
import { registerSubscriptionTools } from "./subscription-tools.js";

interface ToolsContext {
  server: McpServer;
  getAuthManager: () => AuthManager | null;
  setAuthManager: (authManager: AuthManager | null) => void;
  getGraphClient: () => Client | null;
  setGraphClient: (client: Client | null) => void;
  getApiHelper: () => GraphApiHelper | null;
  setApiHelper: (helper: GraphApiHelper | null) => void;
  useGraphBeta: boolean;
  defaultGraphApiVersion: string;
}

export function registerAllTools(context: ToolsContext) {
  const { server, getApiHelper, useGraphBeta, defaultGraphApiVersion } = context;

  // Register API tools
  registerApiTools(server, getApiHelper, useGraphBeta, defaultGraphApiVersion);

  // Register authentication tools
  registerAuthTools(server, {
    getAuthManager: context.getAuthManager,
    setAuthManager: context.setAuthManager,
    getGraphClient: context.getGraphClient,
    setGraphClient: context.setGraphClient,
    getApiHelper: context.getApiHelper,
    setApiHelper: context.setApiHelper,
    useGraphBeta: context.useGraphBeta,
    defaultGraphApiVersion: context.defaultGraphApiVersion
  });

  // Register mail tools
  registerMailTools(server, getApiHelper);

  // Register calendar tools
  registerCalendarTools(server, getApiHelper);

  // Register user tools
  registerUserTools(server, getApiHelper);

  // Register task tools
  registerTaskTools(server, getApiHelper);

  // Register group tools
  registerGroupTools(server, getApiHelper);

  // Register subscription tools
  registerSubscriptionTools(server, getApiHelper);
}
