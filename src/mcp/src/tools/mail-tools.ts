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
    // Handle comma-separated email addresses
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
    "Send an email using Microsoft Graph API. Can optionally return message ID for adding attachments. Requires Mail.Send permission.",
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
          // Send immediately using the sendMail endpoint
          const requestBody = {
            message
          };

          logger.info(`Sending email to ${normalizedTo.map(r => r.email).join(', ')} with subject: ${subject} on behalf of ${userEmail}`);

          const response = await apiHelper.callGraphApi({
            path: `/${userPart}/sendMail`,
            method: 'post',
            body: requestBody
          });

          if (response.error) {
            throw new Error(response.error);
          }

          return {
            content: [{ 
              type: "text" as const, 
              text: `Email sent successfully on behalf of ${userEmail}!\n\nSubject: ${subject}\nTo: ${normalizedTo.map(r => r.email).join(', ')}${normalizedCc ? `\nCC: ${normalizedCc.map(r => r.email).join(', ')}` : ''}${normalizedBcc ? `\nBCC: ${normalizedBcc.map(r => r.email).join(', ')}` : ''}` 
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
    "Send an email with file attachments in a single operation. Only requires Mail.Send permission (no Mail.ReadWrite needed).",
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
        const requestBody = { message };

        logger.info(`Sending email with ${normalizedAttachments.length} attachment(s) to ${normalizedTo.map(r => r.email).join(', ')} with subject: ${subject} on behalf of ${userEmail}`);

        const response = await apiHelper.callGraphApi({
          path: `/${userPart}/sendMail`,
          method: 'post',
          body: requestBody
        });

        if (response.error) {
          throw new Error(response.error);
        }

        const attachmentSummary = normalizedAttachments.map((att, idx) => 
          `  ${idx + 1}. ${att.name || path.basename(att.filePath)} (${fs.statSync(att.filePath).size} bytes)`
        ).join('\n');

        return {
          content: [{ 
            type: "text" as const, 
            text: `Email with attachments sent successfully!\n\nEmail Details:\n- Subject: ${subject}\n- To: ${normalizedTo.map(r => r.email).join(', ')}${normalizedCc ? `\n- CC: ${normalizedCc.map(r => r.email).join(', ')}` : ''}${normalizedBcc ? `\n- BCC: ${normalizedBcc.map(r => r.email).join(', ')}` : ''}\n- User: ${userEmail}\n\nAttachments (${normalizedAttachments.length}):\n${attachmentSummary}` 
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
}
