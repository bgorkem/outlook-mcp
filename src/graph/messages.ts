import type { GraphClient } from "./client.js";

export interface MessageSummary {
  id: string;
  subject: string;
  from?: string;
  to: string[];
  receivedDateTime?: string;
  isRead: boolean;
  hasAttachments: boolean;
  bodyPreview?: string;
  webLink?: string;
}

export interface MessageDetail extends MessageSummary {
  bodyContentType: "text" | "html";
  body: string;
  cc: string[];
  bcc: string[];
}

export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
  totalItemCount?: number;
  unreadItemCount?: number;
}

const SELECT_LIST =
  "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,webLink";

const SELECT_DETAIL = `${SELECT_LIST},body`;

export class MessagesApi {
  constructor(private readonly graph: GraphClient) {}

  async listFolders(): Promise<MailFolder[]> {
    const data = await this.graph.request<{ value: any[] }>("/me/mailFolders", {
      query: { $top: 100 },
    });
    return (data.value ?? []).map(toFolder);
  }

  async listMessages(opts: {
    folder?: string;
    top?: number;
    unreadOnly?: boolean;
  }): Promise<MessageSummary[]> {
    const folderId = opts.folder ? await this.resolveFolder(opts.folder) : null;
    const path = folderId ? `/me/mailFolders/${folderId}/messages` : "/me/messages";
    const query: Record<string, string | number> = {
      $top: clampTop(opts.top),
      $select: SELECT_LIST,
      $orderby: "receivedDateTime desc",
    };
    if (opts.unreadOnly) query.$filter = "isRead eq false";
    const data = await this.graph.request<{ value: any[] }>(path, { query });
    return (data.value ?? []).map(toSummary);
  }

  async searchMessages(query: string, top?: number): Promise<MessageSummary[]> {
    const data = await this.graph.request<{ value: any[] }>("/me/messages", {
      query: {
        $top: clampTop(top),
        $select: SELECT_LIST,
        $search: `"${query.replace(/"/g, '\\"')}"`,
      },
      headers: { ConsistencyLevel: "eventual" },
    });
    return (data.value ?? []).map(toSummary);
  }

  async getMessage(id: string): Promise<MessageDetail> {
    const data = await this.graph.request<any>(`/me/messages/${encodeURIComponent(id)}`, {
      query: { $select: SELECT_DETAIL },
    });
    return toDetail(data);
  }

  async moveMessage(id: string, destination: string): Promise<{ id: string }> {
    const destinationId = await this.resolveFolder(destination);
    const data = await this.graph.request<{ id: string }>(
      `/me/messages/${encodeURIComponent(id)}/move`,
      { method: "POST", body: { destinationId } },
    );
    return { id: data.id };
  }

  async setRead(id: string, isRead: boolean): Promise<void> {
    await this.graph.request(`/me/messages/${encodeURIComponent(id)}`, {
      method: "PATCH",
      body: { isRead },
      expectJson: false,
    });
  }

  async setFlag(id: string, flagged: boolean): Promise<void> {
    await this.graph.request(`/me/messages/${encodeURIComponent(id)}`, {
      method: "PATCH",
      body: { flag: { flagStatus: flagged ? "flagged" : "notFlagged" } },
      expectJson: false,
    });
  }

  async createFolder(name: string, parent?: string): Promise<MailFolder> {
    const path = parent
      ? `/me/mailFolders/${await this.resolveFolder(parent)}/childFolders`
      : "/me/mailFolders";
    const data = await this.graph.request<any>(path, {
      method: "POST",
      body: { displayName: name },
    });
    return toFolder(data);
  }

  async renameFolder(id: string, newName: string): Promise<MailFolder> {
    const folderId = await this.resolveFolder(id);
    const data = await this.graph.request<any>(`/me/mailFolders/${folderId}`, {
      method: "PATCH",
      body: { displayName: newName },
    });
    return toFolder(data);
  }

  private folderCache = new Map<string, string>();

  async resolveFolder(input: string): Promise<string> {
    const trimmed = input.trim();
    if (trimmed.length === 0) throw new Error("Folder reference is empty");

    const cached = this.folderCache.get(trimmed.toLowerCase());
    if (cached) return cached;

    if (isWellKnown(trimmed)) {
      const id = trimmed.toLowerCase();
      this.folderCache.set(id, id);
      return id;
    }

    if (looksLikeId(trimmed)) {
      this.folderCache.set(trimmed.toLowerCase(), trimmed);
      return trimmed;
    }

    const folders = await this.listFolders();
    const match = folders.find((f) => f.displayName.toLowerCase() === trimmed.toLowerCase());
    if (!match) {
      throw new Error(
        `Folder "${input}" not found. Use list_folders to see available folders, or pass a well-known name (inbox, archive, drafts, deleteditems, junkemail, sentitems).`,
      );
    }
    this.folderCache.set(trimmed.toLowerCase(), match.id);
    return match.id;
  }
}

const WELL_KNOWN = new Set([
  "inbox",
  "archive",
  "drafts",
  "deleteditems",
  "junkemail",
  "sentitems",
  "outbox",
  "clutter",
  "conflicts",
  "conversationhistory",
  "localfailures",
  "msgfolderroot",
  "recoverableitemsdeletions",
  "scheduled",
  "searchfolders",
  "serverfailures",
  "syncissues",
]);

function isWellKnown(s: string): boolean {
  return WELL_KNOWN.has(s.toLowerCase());
}

function looksLikeId(s: string): boolean {
  // Graph mail folder ids are long base64-ish strings; well-known names are short words
  return s.length > 30 && /[A-Za-z0-9_-]/.test(s);
}

function clampTop(top: number | undefined): number {
  if (!top || top <= 0) return 25;
  return Math.min(top, 50);
}

function recipientsToEmails(arr: any): string[] {
  if (!Array.isArray(arr)) return [];
  return arr
    .map((r) => r?.emailAddress?.address)
    .filter((s): s is string => typeof s === "string");
}

function toSummary(m: any): MessageSummary {
  return {
    id: m.id,
    subject: m.subject ?? "",
    from: m.from?.emailAddress?.address,
    to: recipientsToEmails(m.toRecipients),
    receivedDateTime: m.receivedDateTime,
    isRead: !!m.isRead,
    hasAttachments: !!m.hasAttachments,
    bodyPreview: m.bodyPreview,
    webLink: m.webLink,
  };
}

function toDetail(m: any): MessageDetail {
  const summary = toSummary(m);
  const contentType = (m.body?.contentType ?? "text").toLowerCase() === "html" ? "html" : "text";
  return {
    ...summary,
    cc: recipientsToEmails(m.ccRecipients),
    bcc: recipientsToEmails(m.bccRecipients),
    bodyContentType: contentType,
    body: m.body?.content ?? "",
  };
}

function toFolder(f: any): MailFolder {
  return {
    id: f.id,
    displayName: f.displayName,
    parentFolderId: f.parentFolderId,
    totalItemCount: f.totalItemCount,
    unreadItemCount: f.unreadItemCount,
  };
}
