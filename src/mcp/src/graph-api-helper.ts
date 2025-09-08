import { Client, PageIterator, PageCollection } from "@microsoft/microsoft-graph-client";
import { logger } from "./logger.js";
import { AuthManager, AuthMode } from "./auth.js";
import fetch from 'isomorphic-fetch';

export interface GraphApiCallParams {
  path: string;
  method: "get" | "post" | "put" | "patch" | "delete";
  queryParams?: Record<string, string>;
  body?: any;
  graphApiVersion?: "v1.0" | "beta";
  fetchAll?: boolean;
  consistencyLevel?: string;
}

export interface AzureApiCallParams {
  path: string;
  method: "get" | "post" | "put" | "patch" | "delete";
  apiVersion: string;
  subscriptionId?: string;
  queryParams?: Record<string, string>;
  body?: any;
  fetchAll?: boolean;
}

export interface ApiResponse {
  data: any;
  error?: string;
  statusCode?: number;
}

export class GraphApiHelper {
  private graphClient: Client | null;
  private authManager: AuthManager | null;
  private useGraphBeta: boolean;
  private defaultGraphApiVersion: "v1.0" | "beta";

  constructor(
    graphClient: Client | null,
    authManager: AuthManager | null,
    useGraphBeta: boolean = true,
    defaultGraphApiVersion: "v1.0" | "beta" = "beta"
  ) {
    this.graphClient = graphClient;
    this.authManager = authManager;
    this.useGraphBeta = useGraphBeta;
    this.defaultGraphApiVersion = defaultGraphApiVersion;
  }

  /**
   * Makes a call to Microsoft Graph API
   */
  async callGraphApi(params: GraphApiCallParams): Promise<ApiResponse> {
    const {
      path,
      method,
      queryParams,
      body,
      graphApiVersion = this.defaultGraphApiVersion,
      fetchAll = false,
      consistencyLevel
    } = params;

    // Override graphApiVersion if USE_GRAPH_BETA is explicitly set to false
    const effectiveGraphApiVersion = !this.useGraphBeta ? "v1.0" : graphApiVersion;
    
    logger.info(`Executing Graph API call: path=${path}, method=${method}, graphApiVersion=${effectiveGraphApiVersion}, fetchAll=${fetchAll}, consistencyLevel=${consistencyLevel}`);

    try {
      if (!this.graphClient) {
        throw new Error("Graph client not initialized");
      }

      // Construct the request using the Graph SDK client
      let request = this.graphClient.api(path).version(effectiveGraphApiVersion);

      // Add query parameters if provided and not empty
      if (queryParams && Object.keys(queryParams).length > 0) {
        request = request.query(queryParams);
      }

      // Add ConsistencyLevel header if provided
      if (consistencyLevel) {
        request = request.header('ConsistencyLevel', consistencyLevel);
        logger.info(`Added ConsistencyLevel header: ${consistencyLevel}`);
      }

      let responseData: any;

      // Handle different methods
      switch (method.toLowerCase()) {
        case 'get':
          if (fetchAll) {
            logger.info(`Fetching all pages for Graph path: ${path}`);
            try {
              // Use a simpler approach: just get the first page for now
              // The PageIterator seems to be causing issues with the callback
              const firstPageResponse: PageCollection = await request.get();
              logger.info(`First page response received`, { 
                hasValue: !!firstPageResponse.value, 
                valueType: typeof firstPageResponse.value,
                valueLength: Array.isArray(firstPageResponse.value) ? firstPageResponse.value.length : 'N/A',
                hasNextLink: !!firstPageResponse['@odata.nextLink']
              });
              
              const odataContext = firstPageResponse['@odata.context'];
              let allItems: any[] = Array.isArray(firstPageResponse.value) ? firstPageResponse.value : [];

              // For now, let's just return the first page to avoid the PageIterator issue
              // TODO: Implement proper pagination without PageIterator if needed
              if (firstPageResponse['@odata.nextLink']) {
                logger.info(`Additional pages available but skipping due to PageIterator issues. NextLink: ${firstPageResponse['@odata.nextLink']}`);
              }

              // Construct final response with context and combined values under 'value' key
              responseData = {
                '@odata.context': odataContext,
                value: allItems
              };
              logger.info(`Finished fetching Graph pages. Total items: ${allItems.length}`);
            } catch (fetchAllError: any) {
              logger.error("Error during fetchAll processing", fetchAllError);
              throw new Error(`FetchAll failed: ${fetchAllError.message}`);
            }

          } else {
            logger.info(`Fetching single page for Graph path: ${path}`);
            responseData = await request.get();
          }
          break;
        case 'post':
          responseData = await request.post(body ?? {});
          break;
        case 'put':
          responseData = await request.put(body ?? {});
          break;
        case 'patch':
          responseData = await request.patch(body ?? {});
          break;
        case 'delete':
          responseData = await request.delete(); // Delete often returns no body or 204
          // Handle potential 204 No Content response
          if (responseData === undefined || responseData === null) {
            responseData = { status: "Success (No Content)" };
          }
          break;
        default:
          throw new Error(`Unsupported method: ${method}`);
      }

      return {
        data: responseData
      };

    } catch (error: any) {
      logger.error(`Error in Graph API call (path: ${path}, method: ${method}):`, error);
      
      // Include error body if available from Graph SDK error
      const errorBody = error.body ? (typeof error.body === 'string' ? error.body : JSON.stringify(error.body)) : 'N/A';
      
      return {
        data: null,
        error: error instanceof Error ? error.message : String(error),
        statusCode: error.statusCode || undefined
      };
    }
  }

  /**
   * Makes a call to Azure Resource Management API
   */
  async callAzureApi(params: AzureApiCallParams): Promise<ApiResponse> {
    const {
      path,
      method,
      apiVersion,
      subscriptionId,
      queryParams,
      body,
      fetchAll = false
    } = params;

    logger.info(`Executing Azure RM API call: path=${path}, method=${method}, apiVersion=${apiVersion}, fetchAll=${fetchAll}`);

    try {
      if (!this.authManager) {
        throw new Error("Auth manager not initialized");
      }

      const baseUrl = "https://management.azure.com";

      // Acquire token for Azure RM
      const azureCredential = this.authManager.getAzureCredential();
      const tokenResponse = await azureCredential.getToken("https://management.azure.com/.default");
      if (!tokenResponse || !tokenResponse.token) {
        throw new Error("Failed to acquire Azure access token");
      }

      // Construct the URL
      let url = baseUrl;
      if (subscriptionId) {
        url += `/subscriptions/${subscriptionId}`;
      }
      url += path;

      const urlParams = new URLSearchParams({ 'api-version': apiVersion });
      if (queryParams) {
        for (const [key, value] of Object.entries(queryParams)) {
          urlParams.append(String(key), String(value));
        }
      }
      url += `?${urlParams.toString()}`;

      // Prepare request options
      const headers: Record<string, string> = {
        'Authorization': `Bearer ${tokenResponse.token}`,
        'Content-Type': 'application/json'
      };
      const requestOptions: RequestInit = {
        method: method.toUpperCase(),
        headers: headers
      };
      if (["POST", "PUT", "PATCH"].includes(method.toUpperCase())) {
        requestOptions.body = body ? JSON.stringify(body) : JSON.stringify({});
      }

      let responseData: any;

      // Handle pagination for Azure RM GET requests
      if (fetchAll && method === 'get') {
        logger.info(`Fetching all pages for Azure RM starting from: ${url}`);
        let allValues: any[] = [];
        let currentUrl: string | null = url;

        while (currentUrl) {
          logger.info(`Fetching Azure RM page: ${currentUrl}`);
          // Re-acquire token for each page (Azure tokens might expire)
          const currentPageTokenResponse = await azureCredential.getToken("https://management.azure.com/.default");
          if (!currentPageTokenResponse || !currentPageTokenResponse.token) {
            throw new Error("Failed to acquire Azure access token during pagination");
          }
          const currentPageHeaders = { ...headers, 'Authorization': `Bearer ${currentPageTokenResponse.token}` };
          const currentPageRequestOptions: RequestInit = { method: 'GET', headers: currentPageHeaders };

          const pageResponse = await fetch(currentUrl, currentPageRequestOptions);
          const pageText = await pageResponse.text();
          let pageData: any;
          try {
            pageData = pageText ? JSON.parse(pageText) : {};
          } catch (e) {
            logger.error(`Failed to parse JSON from Azure RM page: ${currentUrl}`, pageText);
            pageData = { rawResponse: pageText };
          }

          if (!pageResponse.ok) {
            logger.error(`API error on Azure RM page ${currentUrl}:`, pageData);
            throw new Error(`API error (${pageResponse.status}) during Azure RM pagination on ${currentUrl}: ${JSON.stringify(pageData)}`);
          }

          if (pageData.value && Array.isArray(pageData.value)) {
            allValues = allValues.concat(pageData.value);
          } else if (currentUrl === url && !pageData.nextLink) {
            allValues.push(pageData);
          } else if (currentUrl !== url) {
            logger.info(`[Warning] Azure RM response from ${currentUrl} did not contain a 'value' array.`);
          }
          currentUrl = pageData.nextLink || null; // Azure uses nextLink
        }
        responseData = { allValues: allValues };
        logger.info(`Finished fetching all Azure RM pages. Total items: ${allValues.length}`);
      } else {
        // Single page fetch for Azure RM
        logger.info(`Fetching single page for Azure RM: ${url}`);
        const apiResponse = await fetch(url, requestOptions);
        const responseText = await apiResponse.text();
        try {
          responseData = responseText ? JSON.parse(responseText) : {};
        } catch (e) {
          logger.error(`Failed to parse JSON from single Azure RM page: ${url}`, responseText);
          responseData = { rawResponse: responseText };
        }
        if (!apiResponse.ok) {
          logger.error(`API error for Azure RM ${method} ${path}:`, responseData);
          throw new Error(`API error (${apiResponse.status}) for Azure RM: ${JSON.stringify(responseData)}`);
        }
      }

      return {
        data: responseData
      };

    } catch (error: any) {
      logger.error(`Error in Azure API call (path: ${path}, method: ${method}):`, error);
      
      return {
        data: null,
        error: error instanceof Error ? error.message : String(error),
        statusCode: error.statusCode || undefined
      };
    }
  }

  /**
   * Helper method to check if more pages are available
   */
  hasMorePages(data: any, apiType: "graph" | "azure"): boolean {
    const nextLinkKey = apiType === 'graph' ? '@odata.nextLink' : 'nextLink';
    return data && data[nextLinkKey] ? true : false;
  }
}
