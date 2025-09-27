import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { logger } from "../logger.js";
import { GraphApiHelper } from "../graph-api-helper.js";
import * as fs from "fs";
import * as path from "path";

// Helper schema for boolean parameters that accepts both boolean and string values
const zBooleanParam = () => z.union([
  z.boolean(),
  z.string().transform(str => str.toLowerCase() === 'true')
]);

// Helper function to normalize recipient input
const normalizeRecipients = (recipients: string | Array<{email: string, name?: string}>): Array<{email: string, name?: string}> => {
  if (typeof recipients === 'string') {
    // First, try to parse as JSON in case it's a JSON string from Semantic Kernel
    try {
      const parsed = JSON.parse(recipients);
      if (Array.isArray(parsed)) {
        // If it's an array of strings (email addresses)
        if (parsed.length > 0 && typeof parsed[0] === 'string') {
          return parsed.map((email: string) => ({
            email: email.trim(),
            name: undefined
          }));
        }
        // If it's an array of objects with email/name properties
        return parsed.map((recipient: any) => ({
          email: recipient.email || recipient.address || recipient,
          name: recipient.name || undefined
        }));
      }
      // If it's a single string after parsing
      if (typeof parsed === 'string') {
        return [{
          email: parsed.trim(),
          name: undefined
        }];
      }
    } catch (e) {
      // Not valid JSON, continue with original string processing
    }
    
    // Handle comma-separated email addresses or single email
    return recipients.split(',').map(email => ({
      email: email.trim(),
      name: undefined
    }));
  }
  return recipients;
};

export function registerMailTools(
  server: McpServer,
  getApiHelper: () => GraphApiHelper | null
) {
  server.tool(
    "send-mail",
    "Send an email using Microsoft Graph API. Always returns message ID for tracking. Requires Mail.Send permission.",
    {
      userEmail: z.string().describe("Email address of the user on whose behalf to send the email (use 'me' for authenticated user)"),
      to: z.union([
        z.string().describe("Comma-separated email addresses or single email address"),
        z.array(z.object({
          email: z.string().describe("Email address of the recipient"),
          name: z.string().optional().describe("Display name of the recipient")
        }))
      ]).describe("Recipients (to field) - can be a string of comma-separated emails or an array of recipient objects"),
      subject: z.string().describe("Email subject line"),
      body: z.string().describe("Email body content (supports HTML)"),
      bodyType: z.enum(["text", "html"]).optional().default("html").describe("Type of body content"),
      cc: z.union([
        z.string(),
        z.array(z.object({
          email: z.string().describe("Email address of the CC recipient"),
          name: z.string().optional().describe("Display name of the CC recipient")
        }))
      ]).optional().describe("CC recipients - can be a string of comma-separated emails or an array of recipient objects"),
      bcc: z.union([
        z.string(),
        z.array(z.object({
          email: z.string().describe("Email address of the BCC recipient"),
          name: z.string().optional().describe("Display name of the BCC recipient")
        }))
      ]).optional().describe("BCC recipients - can be a string of comma-separated emails or an array of recipient objects"),
      saveAsDraft: zBooleanParam().optional().default(false).describe("If true, saves as draft and returns message ID for adding attachments. If false, sends immediately.")
    },
    async ({ userEmail, to, subject, body, bodyType, cc, bcc, saveAsDraft }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        // Normalize recipients to consistent format
        const normalizedTo = normalizeRecipients(to);
        const normalizedCc = cc ? normalizeRecipients(cc) : undefined;
        const normalizedBcc = bcc ? normalizeRecipients(bcc) : undefined;

        // Build the email message object
        const message: any = {
          subject,
          body: {
            contentType: bodyType,
            content: body
          },
          toRecipients: normalizedTo.map(recipient => ({
            emailAddress: {
              address: recipient.email,
              name: recipient.name
            }
          }))
        };

        // Add CC recipients if provided
        if (normalizedCc && normalizedCc.length > 0) {
          message.ccRecipients = normalizedCc.map(recipient => ({
            emailAddress: {
              address: recipient.email,
              name: recipient.name
            }
          }));
        }

        // Add BCC recipients if provided
        if (normalizedBcc && normalizedBcc.length > 0) {
          message.bccRecipients = normalizedBcc.map(recipient => ({
            emailAddress: {
              address: recipient.email,
              name: recipient.name
            }
          }));
        }

        const userPart = userEmail === 'me' ? 'me' : `users/${userEmail}`;

        if (saveAsDraft) {
          // Create as draft and return message ID
          logger.info(`Creating draft email for ${normalizedTo.map(r => r.email).join(', ')} with subject: ${subject} on behalf of ${userEmail}`);

          const response = await apiHelper.callGraphApi({
            path: `/${userPart}/messages`,
            method: 'post',
            body: message
          });

          if (response.error) {
            throw new Error(response.error);
          }

          const draftMessage = response.data;

          return {
            content: [{ 
              type: "text" as const, 
              text: `Draft email created successfully!\n\nDraft Details:\n- Message ID: ${draftMessage.id}\n- Subject: ${subject}\n- To: ${normalizedTo.map(r => r.email).join(', ')}${normalizedCc ? `\n- CC: ${normalizedCc.map(r => r.email).join(', ')}` : ''}${normalizedBcc ? `\n- BCC: ${normalizedBcc.map(r => r.email).join(', ')}` : ''}\n- User: ${userEmail}\n\nYou can now add attachments using the message ID: ${draftMessage.id}\nTo send the draft, use the send-draft-message tool.` 
            }],
          };
        } else {
          // Create as draft first to get message ID, then send it
          logger.info(`Creating and sending email to ${normalizedTo.map(r => r.email).join(', ')} with subject: ${subject} on behalf of ${userEmail}`);

          // Step 1: Create the draft message
          const createResponse = await apiHelper.callGraphApi({
            path: `/${userPart}/messages`,
            method: 'post',
            body: message
          });

          if (createResponse.error) {
            throw new Error(createResponse.error);
          }

          const draftMessage = createResponse.data;
          const messageId = draftMessage.id;

          // Step 2: Send the draft message
          const sendResponse = await apiHelper.callGraphApi({
            path: `/${userPart}/messages/${messageId}/send`,
            method: 'post'
          });

          if (sendResponse.error) {
            throw new Error(sendResponse.error);
          }

          return {
            content: [{ 
              type: "text" as const, 
              text: `Email sent successfully on behalf of ${userEmail}!\n\nEmail Details:\n- Message ID: ${messageId}\n- Subject: ${subject}\n- To: ${normalizedTo.map(r => r.email).join(', ')}${normalizedCc ? `\n- CC: ${normalizedCc.map(r => r.email).join(', ')}` : ''}${normalizedBcc ? `\n- BCC: ${normalizedBcc.map(r => r.email).join(', ')}` : ''}` 
            }],
          };
        }

      } catch (error: any) {
        logger.error("Error sending email:", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the Mail.Send permission is granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error sending email: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  server.tool(
    "list-mail-messages",
    "List mail messages from a user's mailbox. Requires Mail.Read permission.",
    {
      userEmail: z.string().describe("Email address of the user whose mail messages to list"),
      folder: z.string().optional().describe("Mail folder ID or well-known folder name (e.g., 'inbox', 'sentitems', 'drafts'). If not specified, lists messages from all folders"),
      top: z.union([z.number(), z.string().transform(str => parseInt(str, 10))]).optional().default(10).describe("Number of messages to return (default: 10, max: 999)"),
      skip: z.union([z.number(), z.string().transform(str => parseInt(str, 10))]).optional().describe("Number of messages to skip for pagination"),
      orderBy: z.string().optional().default("receivedDateTime desc").describe("Order by clause (e.g., 'receivedDateTime desc', 'subject asc')"),
      filter: z.string().optional().describe("OData filter query (e.g., \"from/emailAddress/address eq 'sender@example.com'\")"),
      select: z.string().optional().describe("Comma-separated list of fields to return (e.g., 'id,subject,from,receivedDateTime')"),
      fetchAll: zBooleanParam().optional().default(false).describe("Whether to fetch all pages of results (ignores 'top' parameter when true)")
    },
    async ({ userEmail, folder, top, skip, orderBy, filter, select, fetchAll }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        // Construct the path based on whether a folder is specified
        let path = `/users/${userEmail}/`;
        if (folder) {
          // Handle well-known folder names
          const wellKnownFolders = ['inbox', 'sentitems', 'drafts', 'deleteditems', 'junkemail', 'outbox'];
          const folderId = wellKnownFolders.includes(folder.toLowerCase()) ? folder.toLowerCase() : folder;
          path += `mailFolders/${folderId}/messages`;
        } else {
          path += 'messages';
        }

        // Build query parameters
        const queryParams: Record<string, string> = {};
        if (!fetchAll && top) {
          queryParams['$top'] = String(Math.min(top, 999));
        }
        if (skip) {
          queryParams['$skip'] = String(skip);
        }
        if (orderBy) {
          queryParams['$orderby'] = orderBy;
        }
        if (filter) {
          queryParams['$filter'] = filter;
        }
        if (select) {
          queryParams['$select'] = select;
        }

        logger.info(`Listing mail messages for ${userEmail}${folder ? ` in folder ${folder}` : ''}`);

        const response = await apiHelper.callGraphApi({
          path,
          method: 'get',
          queryParams,
          fetchAll
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const messages = response.data.value || [];
        const messageCount = messages.length;

        // Format the message list for display
        const formattedMessages = messages.map((msg: any) => {
          const from = msg.from?.emailAddress?.address || 'Unknown';
          const subject = msg.subject || '(No subject)';
          const received = msg.receivedDateTime ? new Date(msg.receivedDateTime).toLocaleString() : 'Unknown';
          const hasAttachments = msg.hasAttachments ? ' 📎' : '';
          const isRead = msg.isRead ? '' : ' •';
          
          return `${isRead}${hasAttachments} From: ${from} | Subject: ${subject} | Received: ${received} | ID: ${msg.id}`;
        }).join('\n');

        return {
          content: [{ 
            type: "text" as const, 
            text: `Found ${messageCount} mail message(s) for ${userEmail}${folder ? ` in folder ${folder}` : ''}:\n\n${formattedMessages || 'No messages found.'}` 
          }],
        };

      } catch (error: any) {
        logger.error("Error listing mail messages:", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the Mail.Read permission is granted to the application."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error listing mail messages: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  server.tool(
    "get-mail-message",
    "Get a specific mail message by ID. Requires Mail.Read permission.",
    {
      userEmail: z.string().describe("Email address of the user who owns the mail message"),
      messageId: z.string().describe("ID of the mail message to retrieve"),
      select: z.string().optional().describe("Comma-separated list of fields to return (e.g., 'id,subject,from,body,attachments')"),
      includeAttachments: zBooleanParam().optional().default(false).describe("Whether to include attachment details")
    },
    async ({ userEmail, messageId, select, includeAttachments }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        const queryParams: Record<string, string> = {};
        
        // If includeAttachments is true, expand attachments
        if (includeAttachments) {
          queryParams['$expand'] = 'attachments';
        }
        
        // Handle select parameter
        if (select) {
          queryParams['$select'] = select;
        }

        logger.info(`Getting mail message ${messageId} for ${userEmail}`);

        const response = await apiHelper.callGraphApi({
          path: `/users/${userEmail}/messages/${messageId}`,
          method: 'get',
          queryParams
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const msg = response.data;
        
        // Format the message for display
        let messageDetails = `Subject: ${msg.subject || '(No subject)'}\n`;
        messageDetails += `From: ${msg.from?.emailAddress?.name || ''} <${msg.from?.emailAddress?.address || 'Unknown'}>\n`;
        
        // Format recipients
        if (msg.toRecipients && msg.toRecipients.length > 0) {
          const toAddresses = msg.toRecipients.map((r: any) => 
            `${r.emailAddress.name || ''} <${r.emailAddress.address}>`
          ).join(', ');
          messageDetails += `To: ${toAddresses}\n`;
        }
        
        if (msg.ccRecipients && msg.ccRecipients.length > 0) {
          const ccAddresses = msg.ccRecipients.map((r: any) => 
            `${r.emailAddress.name || ''} <${r.emailAddress.address}>`
          ).join(', ');
          messageDetails += `CC: ${ccAddresses}\n`;
        }
        
        messageDetails += `Date: ${msg.receivedDateTime ? new Date(msg.receivedDateTime).toLocaleString() : 'Unknown'}\n`;
        messageDetails += `Read: ${msg.isRead ? 'Yes' : 'No'}\n`;
        messageDetails += `Has Attachments: ${msg.hasAttachments ? 'Yes' : 'No'}\n`;
        
        if (msg.body) {
          messageDetails += `\nBody (${msg.body.contentType}):\n${msg.body.content}\n`;
        }
        
        // Include attachment details if requested
        if (includeAttachments && msg.attachments && msg.attachments.length > 0) {
          messageDetails += '\nAttachments:\n';
          msg.attachments.forEach((att: any, index: number) => {
            messageDetails += `  ${index + 1}. ${att.name} (${att.contentType}, ${att.size} bytes)\n`;
          });
        }

        return {
          content: [{ 
            type: "text" as const, 
            text: messageDetails
          }],
        };

      } catch (error: any) {
        logger.error("Error getting mail message:", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the Mail.Read permission is granted to the application."
          : error.statusCode === 404 
          ? "Message not found. The message ID may be invalid or the message may have been deleted."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error getting mail message: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );


  server.tool(
    "send-mail-with-attachments",
    "Send an email with file attachments in a single operation. Returns message ID for tracking. Only requires Mail.Send permission (no Mail.ReadWrite needed).",
    {
      userEmail: z.string().describe("Email address of the user on whose behalf to send the email (use 'me' for authenticated user)"),
      to: z.union([
        z.string().describe("Comma-separated email addresses or single email address"),
        z.array(z.object({
          email: z.string().describe("Email address of the recipient"),
          name: z.string().optional().describe("Display name of the recipient")
        }))
      ]).describe("Recipients (to field) - can be a string of comma-separated emails or an array of recipient objects"),
      subject: z.string().describe("Email subject line"),
      body: z.string().describe("Email body content (supports HTML)"),
      bodyType: z.enum(["text", "html"]).optional().default("html").describe("Type of body content"),
      attachments: z.union([
        z.string().describe("Single file path to attach"),
        z.array(z.string()).describe("Array of file paths to attach"),
        z.array(z.object({
          filePath: z.string().describe("Path to the file to attach"),
          name: z.string().optional().describe("Custom name for the attachment (defaults to filename)")
        }))
      ]).describe("File attachments - can be a single file path, array of file paths, or array of attachment objects"),
      cc: z.union([
        z.string(),
        z.array(z.object({
          email: z.string().describe("Email address of the CC recipient"),
          name: z.string().optional().describe("Display name of the CC recipient")
        }))
      ]).optional().describe("CC recipients - can be a string of comma-separated emails or an array of recipient objects"),
      bcc: z.union([
        z.string(),
        z.array(z.object({
          email: z.string().describe("Email address of the BCC recipient"),
          name: z.string().optional().describe("Display name of the BCC recipient")
        }))
      ]).optional().describe("BCC recipients - can be a string of comma-separated emails or an array of recipient objects")
    },
    async ({ userEmail, to, subject, body, bodyType, attachments, cc, bcc }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        // Normalize recipients to consistent format
        const normalizedTo = normalizeRecipients(to);
        const normalizedCc = cc ? normalizeRecipients(cc) : undefined;
        const normalizedBcc = bcc ? normalizeRecipients(bcc) : undefined;

        // Normalize attachments to consistent format
        let normalizedAttachments: Array<{filePath: string, name?: string}> = [];
        
        if (typeof attachments === 'string') {
          // Single file path string
          normalizedAttachments = [{ filePath: attachments }];
        } else if (Array.isArray(attachments)) {
          if (attachments.length > 0 && typeof attachments[0] === 'string') {
            // Array of file path strings
            normalizedAttachments = (attachments as string[]).map(filePath => ({ filePath }));
          } else {
            // Array of attachment objects
            normalizedAttachments = attachments as Array<{filePath: string, name?: string}>;
          }
        }

        // Process attachments
        const processedAttachments = [];
        for (const attachment of normalizedAttachments) {
          const { filePath, name } = attachment;

          // Check if file exists
          if (!fs.existsSync(filePath)) {
            throw new Error(`File not found: ${filePath}`);
          }

          // Get file stats
          const fileStats = fs.statSync(filePath);
          if (!fileStats.isFile()) {
            throw new Error(`Path is not a file: ${filePath}`);
          }

          // Check file size (3MB limit)
          const maxSizeBytes = 3 * 1024 * 1024;
          if (fileStats.size > maxSizeBytes) {
            throw new Error(`File size (${fileStats.size} bytes) exceeds maximum allowed size (${maxSizeBytes} bytes) for file: ${filePath}`);
          }

          // Read file and convert to base64
          const fileBuffer = fs.readFileSync(filePath);
          const contentBytes = fileBuffer.toString('base64');

          // Determine attachment name and content type
          const attachmentName = name || path.basename(filePath);
          const fileExtension = path.extname(filePath).toLowerCase();
          const contentTypeMap: Record<string, string> = {
            '.txt': 'text/plain',
            '.pdf': 'application/pdf',
            '.doc': 'application/msword',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.xls': 'application/vnd.ms-excel',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.ppt': 'application/vnd.ms-powerpoint',
            '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.png': 'image/png',
            '.gif': 'image/gif',
            '.zip': 'application/zip',
            '.json': 'application/json',
            '.xml': 'application/xml',
            '.csv': 'text/csv'
          };
          const contentType = contentTypeMap[fileExtension] || 'application/octet-stream';

          processedAttachments.push({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: attachmentName,
            contentType: contentType,
            contentBytes: contentBytes,
            size: fileStats.size
          });
        }

        // Build the email message object with attachments
        const message: any = {
          subject,
          body: {
            contentType: bodyType,
            content: body
          },
          toRecipients: normalizedTo.map(recipient => ({
            emailAddress: {
              address: recipient.email,
              name: recipient.name
            }
          })),
          attachments: processedAttachments
        };

        // Add CC recipients if provided
        if (normalizedCc && normalizedCc.length > 0) {
          message.ccRecipients = normalizedCc.map(recipient => ({
            emailAddress: {
              address: recipient.email,
              name: recipient.name
            }
          }));
        }

        // Add BCC recipients if provided
        if (normalizedBcc && normalizedBcc.length > 0) {
          message.bccRecipients = normalizedBcc.map(recipient => ({
            emailAddress: {
              address: recipient.email,
              name: recipient.name
            }
          }));
        }

        const userPart = userEmail === 'me' ? 'me' : `users/${userEmail}`;

        logger.info(`Creating and sending email with ${normalizedAttachments.length} attachment(s) to ${normalizedTo.map(r => r.email).join(', ')} with subject: ${subject} on behalf of ${userEmail}`);

        // Step 1: Create the draft message with attachments
        const createResponse = await apiHelper.callGraphApi({
          path: `/${userPart}/messages`,
          method: 'post',
          body: message
        });

        if (createResponse.error) {
          throw new Error(createResponse.error);
        }

        const draftMessage = createResponse.data;
        const messageId = draftMessage.id;

        // Step 2: Send the draft message
        const sendResponse = await apiHelper.callGraphApi({
          path: `/${userPart}/messages/${messageId}/send`,
          method: 'post'
        });

        if (sendResponse.error) {
          throw new Error(sendResponse.error);
        }

        const attachmentSummary = normalizedAttachments.map((att, idx) => 
          `  ${idx + 1}. ${att.name || path.basename(att.filePath)} (${fs.statSync(att.filePath).size} bytes)`
        ).join('\n');

        return {
          content: [{ 
            type: "text" as const, 
            text: `Email with attachments sent successfully!\n\nEmail Details:\n- Message ID: ${messageId}\n- Subject: ${subject}\n- To: ${normalizedTo.map(r => r.email).join(', ')}${normalizedCc ? `\n- CC: ${normalizedCc.map(r => r.email).join(', ')}` : ''}${normalizedBcc ? `\n- BCC: ${normalizedBcc.map(r => r.email).join(', ')}` : ''}\n- User: ${userEmail}\n\nAttachments (${normalizedAttachments.length}):\n${attachmentSummary}` 
          }],
        };

      } catch (error: any) {
        logger.error("Error sending email with attachments:", error);
        let errorMessage = error.message;
        
        if (error.statusCode === 403) {
          errorMessage = "Permission denied. Ensure the Mail.Send permission is granted to the application.";
        } else if (error.statusCode === 413) {
          errorMessage = "One or more attachments are too large. Maximum size is 3MB per attachment.";
        }
        
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error sending email with attachments: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  server.tool(
    "check-message-responses",
    "Check for responses(replies) to a specific message using the In-Reply-To email header. This provides accurate message threading by reading the InternetMessageId of the original message and finding messages with matching In-Reply-To headers. Requires Mail.Read permission.",
    {
      userEmail: z.string().describe("Email address of the user whose mailbox to search"),
      originalMessageId: z.string().describe("ID of the original message to check responses for"),
      useInternetHeaders: zBooleanParam().optional().default(false).describe("If true, uses InternetHeaders (slower but more comprehensive). If false, uses extended properties (faster, In-Reply-To only)"),
      searchDays: z.union([z.number(), z.string().transform(str => parseInt(str, 10))]).optional().default(30).describe("Number of days to search back from today (default: 30)"),
      maxResults: z.union([z.number(), z.string().transform(str => parseInt(str, 10))]).optional().default(100).describe("Maximum number of messages to analyze (default: 100)")
    },
    async ({ userEmail, originalMessageId, useInternetHeaders, searchDays, maxResults }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        // First, get the original message to extract its InternetMessageId
        logger.info(`Getting original message ${originalMessageId} for ${userEmail}`);
        
        const originalMessageResponse = await apiHelper.callGraphApi({
          path: `/users/${userEmail}/messages/${originalMessageId}`,
          method: 'get',
          queryParams: {
            '$select': 'id,subject,sentDateTime,from,toRecipients,ccRecipients,internetMessageId'
          }
        });

        if (originalMessageResponse.error) {
          throw new Error(`Failed to get original message: ${originalMessageResponse.error}`);
        }

        const originalMessage = originalMessageResponse.data;
        const originalInternetMessageId = originalMessage.internetMessageId;
        
        if (!originalInternetMessageId) {
          throw new Error("Original message does not have an InternetMessageId. Cannot search for responses.");
        }

        const originalSubject = originalMessage.subject || '';
        const originalSender = originalMessage.from?.emailAddress?.address;
        const originalSentDate = new Date(originalMessage.sentDateTime);

        // Calculate search date range (from original message date to searchDays from now)
        const searchStartDate = originalSentDate.toISOString();
        const searchEndDate = new Date(Date.now() + (searchDays * 24 * 60 * 60 * 1000)).toISOString();

        logger.info(`Searching for responses to message with InternetMessageId: "${originalInternetMessageId}"`);

        let responses: any[] = [];

        if (useInternetHeaders) {
          // Method 1: Use InternetHeaders (comprehensive but slower)
          logger.info("Using InternetHeaders method for comprehensive header analysis");
          
          const queryParams: Record<string, string> = {
            '$filter': `receivedDateTime ge ${searchStartDate} and receivedDateTime le ${searchEndDate}`,
            '$orderby': 'receivedDateTime desc',
            '$top': String(Math.min(maxResults, 999)),
            '$select': 'id,subject,from,receivedDateTime,toRecipients,ccRecipients,body',
            '$expand': 'internetMessageHeaders'
          };

          const searchResponse = await apiHelper.callGraphApi({
            path: `/users/${userEmail}/messages`,
            method: 'get',
            queryParams
          });

          if (searchResponse.error) {
            throw new Error(searchResponse.error);
          }

          const allMessages = searchResponse.data.value || [];
          
          // Filter messages that have In-Reply-To header matching our InternetMessageId
          responses = allMessages.filter((msg: any) => {
            if (msg.id === originalMessageId) return false; // Skip original message
            
            const headers = msg.internetMessageHeaders || [];
            const inReplyToHeader = headers.find((h: any) => 
              h.name.toLowerCase() === 'in-reply-to'
            );
            
            return inReplyToHeader && inReplyToHeader.value.includes(originalInternetMessageId);
          });

        } else {
          // Method 2: Use Extended Properties (faster, In-Reply-To only)
          logger.info("Using Extended Properties method for In-Reply-To header");
          
          // Extended property ID for In-Reply-To header (String 0x1042)
          const inReplyToPropertyId = 'String 0x1042';
          
          const queryParams: Record<string, string> = {
            '$filter': `receivedDateTime ge ${searchStartDate} and receivedDateTime le ${searchEndDate}`,
            '$orderby': 'receivedDateTime desc',
            '$top': String(Math.min(maxResults, 999)),
            '$select': 'id,subject,from,receivedDateTime,toRecipients,ccRecipients,body',
            '$expand': `singleValueExtendedProperties($filter=id eq '${inReplyToPropertyId}')`
          };

          const searchResponse = await apiHelper.callGraphApi({
            path: `/users/${userEmail}/messages`,
            method: 'get',
            queryParams
          });

          if (searchResponse.error) {
            throw new Error(searchResponse.error);
          }

          const allMessages = searchResponse.data.value || [];
          
          // Filter messages that have In-Reply-To extended property matching our InternetMessageId
          responses = allMessages.filter((msg: any) => {
            if (msg.id === originalMessageId) return false; // Skip original message
            
            const extendedProps = msg.singleValueExtendedProperties || [];
            const inReplyToProp = extendedProps.find((prop: any) => 
              prop.id === inReplyToPropertyId
            );
            
            return inReplyToProp && inReplyToProp.value && inReplyToProp.value.includes(originalInternetMessageId);
          });
        }

        // Sort by received date (most recent first)
        responses.sort((a: any, b: any) => 
          new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime()
        );

        if (responses.length === 0) {
          return {
            content: [{ 
              type: "text" as const, 
              text: `No responses found to the message.\n\nOriginal message details:\n- Message ID: ${originalMessageId}\n- InternetMessageId: ${originalInternetMessageId}\n- Subject: ${originalSubject}\n- Sent: ${originalSentDate.toLocaleString()}\n- From: ${originalMessage.from?.emailAddress?.name || ''} <${originalSender || 'Unknown'}>\n\nMethod used: ${useInternetHeaders ? 'InternetHeaders' : 'Extended Properties (In-Reply-To)'}\nSearch period: ${searchDays} days from message date` 
            }],
          };
        }

        // Format the responses
        let responseText = `Found ${responses.length} response(s) to the message using In-Reply-To header matching:\n\n`;
        responseText += `Original message:\n- ID: ${originalMessageId}\n- InternetMessageId: ${originalInternetMessageId}\n- Subject: ${originalSubject}\n- Sent: ${originalSentDate.toLocaleString()}\n- From: ${originalMessage.from?.emailAddress?.name || ''} <${originalSender || 'Unknown'}>\n\n`;
        responseText += `Responses (${useInternetHeaders ? 'via InternetHeaders' : 'via Extended Properties'}):\n`;

        responses.forEach((msg: any, index: number) => {
          const from = msg.from?.emailAddress?.address || 'Unknown';
          const fromName = msg.from?.emailAddress?.name || '';
          const subject = msg.subject || '(No subject)';
          const received = new Date(msg.receivedDateTime).toLocaleString();
          
          responseText += `${index + 1}. From: ${fromName} <${from}>\n`;
          responseText += `   Subject: ${subject}\n`;
          responseText += `   Received: ${received}\n`;
          responseText += `   Message ID: ${msg.id}\n`;
          responseText += `   Body: ${msg.body.content}\n`;
          
          responseText += '\n';
        });

        return {
          content: [{ 
            type: "text" as const, 
            text: responseText
          }],
        };

      } catch (error: any) {
        logger.error("Error checking message responses:", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the Mail.Read permission is granted to the application."
          : error.statusCode === 404 
          ? "Original message not found. The message ID may be invalid or the message may have been deleted."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error checking message responses: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  server.tool(
    "get-message-thread",
    "Build a complete message thread/conversation chain starting from a message. Uses In-Reply-To and References headers to build the full conversation tree. Requires Mail.Read permission.",
    {
      userEmail: z.string().describe("Email address of the user whose mailbox to search"),
      messageId: z.string().describe("ID of any message in the thread (can be original message or any reply)"),
      maxDepth: z.union([z.number(), z.string().transform(str => parseInt(str, 10))]).optional().default(10).describe("Maximum depth to traverse the thread (default: 10)"),
      includeBody: zBooleanParam().optional().default(false).describe("Whether to include message body content in the thread")
    },
    async ({ userEmail, messageId, maxDepth, includeBody }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        logger.info(`Building message thread for message ${messageId} for ${userEmail}`);

        // First, get the starting message
        const startingMessageResponse = await apiHelper.callGraphApi({
          path: `/users/${userEmail}/messages/${messageId}`,
          method: 'get',
          queryParams: {
            '$select': 'id,subject,sentDateTime,receivedDateTime,from,toRecipients,ccRecipients,internetMessageId,body',
            '$expand': 'internetMessageHeaders'
          }
        });

        if (startingMessageResponse.error) {
          throw new Error(`Failed to get starting message: ${startingMessageResponse.error}`);
        }

        const startingMessage = startingMessageResponse.data;
        const messageMap = new Map<string, any>();
        const threadMessages: any[] = [];
        
        // Add starting message to our collections
        messageMap.set(startingMessage.id, startingMessage);
        threadMessages.push(startingMessage);

        // Extract References header to find the root message
        const headers = startingMessage.internetMessageHeaders || [];
        const referencesHeader = headers.find((h: any) => h.name.toLowerCase() === 'references');
        const inReplyToHeader = headers.find((h: any) => h.name.toLowerCase() === 'in-reply-to');
        
        let rootInternetMessageId = startingMessage.internetMessageId;
        
        // If this message is a reply, find the root message
        if (referencesHeader && referencesHeader.value) {
          // References header contains all previous message IDs in the thread
          const references = referencesHeader.value.match(/<[^>]+>/g) || [];
          if (references.length > 0) {
            rootInternetMessageId = references[0].replace(/[<>]/g, ''); // First reference is usually the root
          }
        } else if (inReplyToHeader && inReplyToHeader.value) {
          // If no References, use In-Reply-To as potential root
          const inReplyTo = inReplyToHeader.value.match(/<[^>]+>/);
          if (inReplyTo) {
            rootInternetMessageId = inReplyTo[0].replace(/[<>]/g, '');
          }
        }

        // Search for all messages in the thread using extended properties
        const inReplyToPropertyId = 'String 0x1042'; // In-Reply-To header
        const referencesPropertyId = 'String 0x1040'; // References header

        // Search for messages that reference our root message or starting message
        const searchIds = [rootInternetMessageId, startingMessage.internetMessageId].filter(id => id);
        const allThreadMessages = new Set<string>();

        for (const searchId of searchIds) {
          // Search using extended properties for In-Reply-To
          const queryParams: Record<string, string> = {
            '$orderby': 'sentDateTime asc',
            '$top': '999',
            '$select': includeBody ? 
              'id,subject,sentDateTime,receivedDateTime,from,toRecipients,ccRecipients,internetMessageId,body' :
              'id,subject,sentDateTime,receivedDateTime,from,toRecipients,ccRecipients,internetMessageId',
            '$expand': `internetMessageHeaders,singleValueExtendedProperties($filter=id eq '${inReplyToPropertyId}' or id eq '${referencesPropertyId}')`
          };

          const searchResponse = await apiHelper.callGraphApi({
            path: `/users/${userEmail}/messages`,
            method: 'get',
            queryParams
          });

          if (searchResponse.error) {
            logger.error(`Error searching for thread messages: ${searchResponse.error}`);
            continue;
          }

          const messages = searchResponse.data.value || [];
          
          // Filter messages that are part of this thread
          for (const msg of messages) {
            const msgHeaders = msg.internetMessageHeaders || [];
            const msgInReplyTo = msgHeaders.find((h: any) => h.name.toLowerCase() === 'in-reply-to');
            const msgReferences = msgHeaders.find((h: any) => h.name.toLowerCase() === 'references');
            
            // Check if this message references our search ID
            const isInThread = 
              (msgInReplyTo && msgInReplyTo.value.includes(searchId)) ||
              (msgReferences && msgReferences.value.includes(searchId)) ||
              msg.internetMessageId === searchId;

            if (isInThread && !messageMap.has(msg.id)) {
              messageMap.set(msg.id, msg);
              threadMessages.push(msg);
              allThreadMessages.add(msg.internetMessageId);
            }
          }
        }

        // Sort all messages by sent date
        threadMessages.sort((a, b) => 
          new Date(a.sentDateTime || a.receivedDateTime).getTime() - 
          new Date(b.sentDateTime || b.receivedDateTime).getTime()
        );

        // Build the thread structure
        const threadStructure: any[] = [];
        const processedIds = new Set<string>();

        for (const msg of threadMessages) {
          if (processedIds.has(msg.id)) continue;
          
          const headers = msg.internetMessageHeaders || [];
          const inReplyTo = headers.find((h: any) => h.name.toLowerCase() === 'in-reply-to');
          
          const threadItem = {
            id: msg.id,
            internetMessageId: msg.internetMessageId,
            subject: msg.subject || '(No subject)',
            from: {
              name: msg.from?.emailAddress?.name || '',
              address: msg.from?.emailAddress?.address || 'Unknown'
            },
            sentDateTime: msg.sentDateTime || msg.receivedDateTime,
            inReplyTo: inReplyTo ? inReplyTo.value.replace(/[<>]/g, '') : null,
            body: includeBody && msg.body ? 
              msg.body.content?.replace(/<[^>]*>/g, '').substring(0, 300) + 
              (msg.body.content?.length > 300 ? '...' : '') : null,
            isStartingMessage: msg.id === messageId
          };
          
          threadStructure.push(threadItem);
          processedIds.add(msg.id);
        }

        if (threadStructure.length === 0) {
          return {
            content: [{ 
              type: "text" as const, 
              text: `No thread found for message ${messageId}. This might be a standalone message with no replies or references.` 
            }],
          };
        }

        // Format the thread for display
        let threadText = `Message Thread (${threadStructure.length} message${threadStructure.length === 1 ? '' : 's'}):\n\n`;
        
        threadStructure.forEach((msg, index) => {
          const sentDate = new Date(msg.sentDateTime).toLocaleString();
          const marker = msg.isStartingMessage ? ' ← STARTING MESSAGE' : '';
          const replyIndicator = msg.inReplyTo ? '↳ ' : '';
          
          threadText += `${replyIndicator}${index + 1}. ${msg.from.name} <${msg.from.address}>${marker}\n`;
          threadText += `   Subject: ${msg.subject}\n`;
          threadText += `   Sent: ${sentDate}\n`;
          threadText += `   Message ID: ${msg.id}\n`;
          if (msg.inReplyTo) {
            threadText += `   In-Reply-To: ${msg.inReplyTo}\n`;
          }
          if (msg.body) {
            threadText += `   Body: ${msg.body}\n`;
          }
          threadText += '\n';
        });

        threadText += `Thread Analysis:\n`;
        threadText += `- Root message: ${threadStructure[0]?.subject || 'Unknown'}\n`;
        threadText += `- Total messages: ${threadStructure.length}\n`;
        threadText += `- Date range: ${new Date(threadStructure[0]?.sentDateTime).toLocaleDateString()} - ${new Date(threadStructure[threadStructure.length - 1]?.sentDateTime).toLocaleDateString()}\n`;

        return {
          content: [{ 
            type: "text" as const, 
            text: threadText
          }],
        };

      } catch (error: any) {
        logger.error("Error building message thread:", error);
        const errorMessage = error.statusCode === 403 
          ? "Permission denied. Ensure the Mail.Read permission is granted to the application."
          : error.statusCode === 404 
          ? "Message not found. The message ID may be invalid or the message may have been deleted."
          : error.message;
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error building message thread: ${errorMessage}` 
          }],
          isError: true
        };
      }
    }
  );

  server.tool(
    "debug-message-access",
    "Debug tool to troubleshoot message access issues. Helps identify why a message ID might not be found. Requires Mail.Read permission.",
    {
      userEmail: z.string().describe("Email address of the user whose mailbox to search"),
      messageId: z.string().describe("ID of the message to debug"),
      searchAllFolders: zBooleanParam().optional().default(false).describe("If true, searches all mail folders for the message")
    },
    async ({ userEmail, messageId, searchAllFolders }) => {
      try {
        const apiHelper = getApiHelper();
        if (!apiHelper) {
          throw new Error("API helper not initialized");
        }

        logger.info(`Debugging message access for ${messageId} in ${userEmail}'s mailbox`);

        let debugInfo = `Debug Report for Message ID: ${messageId}\nUser: ${userEmail}\n\n`;

        // Step 1: Try to get the message directly
        debugInfo += "Step 1: Direct message access attempt\n";
        const directResponse = await apiHelper.callGraphApi({
          path: `/users/${userEmail}/messages/${messageId}`,
          method: 'get',
          queryParams: {
            '$select': 'id,subject,sentDateTime,receivedDateTime,from,parentFolderId,internetMessageId'
          }
        });

        if (directResponse.error) {
          debugInfo += `❌ Direct access failed: ${directResponse.error}\n\n`;
          
          if (searchAllFolders) {
            // Step 2: Search in different folders
            debugInfo += "Step 2: Searching in different mail folders\n";
            const wellKnownFolders = ['inbox', 'sentitems', 'drafts', 'deleteditems', 'junkemail', 'outbox'];
            
            for (const folder of wellKnownFolders) {
              try {
                const folderResponse = await apiHelper.callGraphApi({
                  path: `/users/${userEmail}/mailFolders/${folder}/messages/${messageId}`,
                  method: 'get',
                  queryParams: {
                    '$select': 'id,subject,sentDateTime,receivedDateTime,from,parentFolderId'
                  }
                });

                if (!folderResponse.error) {
                  debugInfo += `✅ Found in folder: ${folder}\n`;
                  const msg = folderResponse.data;
                  debugInfo += `   Subject: ${msg.subject || '(No subject)'}\n`;
                  debugInfo += `   From: ${msg.from?.emailAddress?.address || 'Unknown'}\n`;
                  debugInfo += `   Sent: ${msg.sentDateTime || msg.receivedDateTime || 'Unknown'}\n`;
                  debugInfo += `   Parent Folder ID: ${msg.parentFolderId}\n\n`;
                  break;
                } else {
                  debugInfo += `❌ Not in ${folder}: ${folderResponse.error}\n`;
                }
              } catch (e) {
                debugInfo += `❌ Error checking ${folder}: ${e}\n`;
              }
            }

            // Step 3: Search by partial message ID match
            debugInfo += "\nStep 3: Searching recent messages for partial ID match\n";
            try {
              const recentResponse = await apiHelper.callGraphApi({
                path: `/users/${userEmail}/messages`,
                method: 'get',
                queryParams: {
                  '$top': '50',
                  '$orderby': 'receivedDateTime desc',
                  '$select': 'id,subject,sentDateTime,receivedDateTime,from'
                }
              });

              if (!recentResponse.error) {
                const messages = recentResponse.data.value || [];
                const partialMatches = messages.filter((msg: any) => 
                  msg.id.includes(messageId.substring(0, 20)) || 
                  messageId.includes(msg.id.substring(0, 20))
                );

                if (partialMatches.length > 0) {
                  debugInfo += `Found ${partialMatches.length} partial match(es):\n`;
                  partialMatches.forEach((msg: any, index: number) => {
                    debugInfo += `  ${index + 1}. ID: ${msg.id}\n`;
                    debugInfo += `     Subject: ${msg.subject || '(No subject)'}\n`;
                    debugInfo += `     From: ${msg.from?.emailAddress?.address || 'Unknown'}\n`;
                    debugInfo += `     Date: ${msg.sentDateTime || msg.receivedDateTime || 'Unknown'}\n\n`;
                  });
                } else {
                  debugInfo += "No partial matches found in recent messages\n\n";
                }
              }
            } catch (e) {
              debugInfo += `Error searching recent messages: ${e}\n\n`;
            }
          }
        } else {
          debugInfo += "✅ Direct access successful!\n";
          const msg = directResponse.data;
          debugInfo += `   Subject: ${msg.subject || '(No subject)'}\n`;
          debugInfo += `   From: ${msg.from?.emailAddress?.address || 'Unknown'}\n`;
          debugInfo += `   Sent: ${msg.sentDateTime || msg.receivedDateTime || 'Unknown'}\n`;
          debugInfo += `   Parent Folder ID: ${msg.parentFolderId}\n`;
          debugInfo += `   InternetMessageId: ${msg.internetMessageId || 'Not available'}\n\n`;
        }

        // Step 4: Check user permissions and mailbox access
        debugInfo += "Step 4: User mailbox access verification\n";
        try {
          const userResponse = await apiHelper.callGraphApi({
            path: `/users/${userEmail}`,
            method: 'get',
            queryParams: {
              '$select': 'id,displayName,mail,userPrincipalName'
            }
          });

          if (!userResponse.error) {
            debugInfo += "✅ User exists and is accessible\n";
            debugInfo += `   Display Name: ${userResponse.data.displayName}\n`;
            debugInfo += `   Mail: ${userResponse.data.mail}\n`;
            debugInfo += `   UPN: ${userResponse.data.userPrincipalName}\n\n`;
          } else {
            debugInfo += `❌ User access issue: ${userResponse.error}\n\n`;
          }
        } catch (e) {
          debugInfo += `❌ Error checking user: ${e}\n\n`;
        }

        // Step 5: Recommendations
        debugInfo += "Recommendations:\n";
        if (directResponse.error) {
          if (directResponse.error.includes('404') || directResponse.error.includes('NotFound')) {
            debugInfo += "• Message not found - it may have been deleted, moved, or the ID is incorrect\n";
            debugInfo += "• Try using the search functionality to find the message by subject or sender\n";
            if (!searchAllFolders) {
              debugInfo += "• Run this debug tool again with searchAllFolders=true to check other folders\n";
            }
          } else if (directResponse.error.includes('403') || directResponse.error.includes('Forbidden')) {
            debugInfo += "• Permission issue - ensure the application has Mail.Read permission\n";
            debugInfo += "• Check if the user has granted consent for the application\n";
          } else {
            debugInfo += `• Unexpected error: ${directResponse.error}\n`;
            debugInfo += "• Try using a different message ID or check the API endpoint\n";
          }
        } else {
          debugInfo += "• Message is accessible - the check-message-responses tool should work\n";
        }

        return {
          content: [{ 
            type: "text" as const, 
            text: debugInfo
          }],
        };

      } catch (error: any) {
        logger.error("Error in debug tool:", error);
        return {
          content: [{ 
            type: "text" as const, 
            text: `Debug tool error: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );
}
