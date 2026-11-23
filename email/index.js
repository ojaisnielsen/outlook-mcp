/**
 * Email module for Outlook MCP server
 */
const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleDraftEmail = require('./draft');
const handleMarkAsRead = require('./mark-as-read');
const handleDeleteEmail = require('./delete');
const handleFlagEmail = require('./flag');
const { handleDraftReply, handleSendDraft } = require('./reply');

// Filter parameters shared by list-emails and search-emails
const FILTER_PROPERTIES = {
  receivedAfter: {
    type: "string",
    description: "Only emails received at or after this ISO 8601 date/time (e.g. '2026-11-20T06:00:00Z')"
  },
  receivedBefore: {
    type: "string",
    description: "Only emails received before this ISO 8601 date/time"
  },
  flagStatus: {
    type: "string",
    enum: ["flagged", "notFlagged", "complete"],
    description: "Only emails with this follow-up flag status"
  }
};

// Email tool definitions
const emailTools = [
  {
    name: "list-emails",
    description: "Lists emails in a folder (newest first by default) with sender, ISO date, unread/flag/attachment markers, a short preview and the ID. Does not change read state.",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Email folder to list (e.g., 'inbox', 'sent', 'drafts', or a folder name; default: 'inbox'). An unknown name is an error."
        },
        count: {
          type: "number",
          description: "Number of emails to retrieve (default: 10, max: 500; pages through the folder)"
        },
        unreadOnly: {
          type: "boolean",
          description: "Only unread emails"
        },
        hasAttachments: {
          type: "boolean",
          description: "Only emails with attachments"
        },
        ...FILTER_PROPERTIES,
        oldestFirst: {
          type: "boolean",
          description: "Sort oldest first instead of newest first"
        },
        includePreview: {
          type: "boolean",
          description: "Include a ~200-character body preview per email (default: true)"
        }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search for emails using various criteria",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text to find in emails"
        },
        folder: {
          type: "string",
          description: "Email folder to search in (default: 'inbox')"
        },
        from: {
          type: "string",
          description: "Filter by sender email address or name"
        },
        to: {
          type: "string",
          description: "Filter by recipient email address or name"
        },
        subject: {
          type: "string",
          description: "Filter by email subject"
        },
        hasAttachments: {
          type: "boolean",
          description: "Filter to only emails with attachments"
        },
        unreadOnly: {
          type: "boolean",
          description: "Filter to only unread emails"
        },
        ...FILTER_PROPERTIES,
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 500)"
        },
        strict: {
          type: "boolean",
          description: "Only return emails matching all criteria. Without it, if nothing matches, broader searches are tried and the result is marked with a WARNING."
        }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Reads the content of a specific email. HTML emails are securely sanitized to extract only visible text, preventing prompt injection attacks via hidden content.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to read"
        },
        includeRawHtml: {
          type: "boolean",
          description: "Include raw HTML content (UNSAFE - for debugging only, may contain hidden prompt injection content)"
        }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "send-email",
    description: "Composes and sends a new email. Supports both plain text and HTML content.",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content (plain text or HTML)"
        },
        isHtml: {
          type: "boolean",
          description: "Set to true to send as HTML, false for plain text. If not specified, auto-detects based on <html> tag presence."
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        saveToSentItems: {
          type: "boolean",
          description: "Whether to save the email to sent items"
        }
      },
      required: ["to", "subject", "body"]
    },
    handler: handleSendEmail
  },
  {
    name: "draft-email",
    description: "Creates and saves an email draft in Outlook",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Draft email subject"
        },
        body: {
          type: "string",
          description: "Draft email body content (can be plain text or HTML)"
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        }
      },
      required: []
    },
    handler: handleDraftEmail
  },
  {
    name: "mark-as-read",
    description: "Marks an email as read or unread",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to mark as read/unread"
        },
        isRead: {
          type: "boolean",
          description: "Whether to mark as read (true) or unread (false). Default: true"
        }
      },
      required: ["id"]
    },
    handler: handleMarkAsRead
  },
  {
    name: "delete-email",
    description: "Deletes an email by moving it to Deleted Items (trash). Use permanent=true to hard delete.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to delete"
        },
        permanent: {
          type: "boolean",
          description: "If true, permanently delete the email instead of moving to Deleted Items. Default: false"
        }
      },
      required: ["id"]
    },
    handler: handleDeleteEmail
  },
  {
    name: "flag-email",
    description: "Sets the follow-up flag on one or more emails: flag, unflag (notFlagged) or mark complete. Flagged emails also appear in Microsoft To Do. Does not change read state.",
    inputSchema: {
      type: "object",
      properties: {
        ids: {
          type: "string",
          description: "Email ID, or a comma-separated list of IDs"
        },
        status: {
          type: "string",
          enum: ["flagged", "notFlagged", "complete"],
          description: "Flag status to set (default: flagged)"
        }
      },
      required: ["ids"]
    },
    handler: handleFlagEmail
  },
  {
    name: "draft-reply",
    description: "Creates a reply or reply-all draft in the original thread (with the quoted conversation) and returns its recipients for checking. Never sends.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to reply to (normally the latest message in the thread)"
        },
        body: {
          type: "string",
          description: "Plain-text reply, placed above the quoted thread"
        },
        replyAll: {
          type: "boolean",
          description: "Reply to all recipients (default: true)"
        }
      },
      required: ["id"]
    },
    handler: handleDraftReply
  },
  {
    name: "send-draft",
    description: "Sends an existing draft (e.g. one created by draft-reply or draft-email). Refuses anything that is not a draft.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the draft to send"
        }
      },
      required: ["id"]
    },
    handler: handleSendDraft
  }
];

module.exports = {
  emailTools,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleDraftEmail,
  handleMarkAsRead,
  handleDeleteEmail,
  handleFlagEmail,
  handleDraftReply,
  handleSendDraft
};
