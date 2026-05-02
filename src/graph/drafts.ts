import type { GraphClient } from "./client.js";

export interface DraftInput {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
  bodyType?: "text" | "html";
}

export interface DraftSummary {
  id: string;
  subject: string;
  to: string[];
  cc: string[];
  bcc: string[];
  bodyPreview?: string;
  lastModifiedDateTime?: string;
  webLink?: string;
}

export interface DraftDetail extends DraftSummary {
  bodyContentType: "text" | "html";
  body: string;
}

const SELECT =
  "id,subject,toRecipients,ccRecipients,bccRecipients,bodyPreview,lastModifiedDateTime,webLink,body,isDraft";

export class DraftsApi {
  constructor(private readonly graph: GraphClient) {}

  async create(input: DraftInput & { replyToMessageId?: string; replyAll?: boolean }): Promise<DraftDetail> {
    if (input.replyToMessageId) {
      return this.createReplyDraft(input.replyToMessageId, !!input.replyAll, input);
    }
    const data = await this.graph.request<any>("/me/messages", {
      method: "POST",
      body: buildMessagePayload(input),
    });
    return toDetail(data);
  }

  private async createReplyDraft(
    parentId: string,
    replyAll: boolean,
    input: DraftInput,
  ): Promise<DraftDetail> {
    const path = replyAll
      ? `/me/messages/${encodeURIComponent(parentId)}/createReplyAll`
      : `/me/messages/${encodeURIComponent(parentId)}/createReply`;
    const created = await this.graph.request<any>(path, { method: "POST", body: {} });
    return this.update(created.id, input);
  }

  async update(id: string, input: DraftInput): Promise<DraftDetail> {
    const payload = buildMessagePayload(input);
    if (Object.keys(payload).length === 0) {
      return this.get(id);
    }
    const data = await this.graph.request<any>(`/me/messages/${encodeURIComponent(id)}`, {
      method: "PATCH",
      body: payload,
    });
    return toDetail(data);
  }

  async get(id: string): Promise<DraftDetail> {
    const data = await this.graph.request<any>(`/me/messages/${encodeURIComponent(id)}`, {
      query: { $select: SELECT },
    });
    return toDetail(data);
  }

  async list(top: number): Promise<DraftSummary[]> {
    const data = await this.graph.request<{ value: any[] }>("/me/mailFolders/drafts/messages", {
      query: {
        $top: Math.min(Math.max(top, 1), 50),
        $select: SELECT,
        $orderby: "lastModifiedDateTime desc",
      },
    });
    return (data.value ?? []).map(toSummary);
  }

  async delete(id: string): Promise<void> {
    const detail = await this.get(id).catch(() => null);
    if (detail && (detail as any).isDraft === false) {
      throw new Error(
        `Refusing to delete message ${id}: it is not a draft. ` +
          "delete_draft only operates on items in the Drafts folder.",
      );
    }
    await this.graph.request(`/me/messages/${encodeURIComponent(id)}`, {
      method: "DELETE",
      expectJson: false,
    });
  }
}

function buildMessagePayload(input: DraftInput): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (input.subject !== undefined) out.subject = input.subject;
  if (input.to) out.toRecipients = toRecipients(input.to);
  if (input.cc) out.ccRecipients = toRecipients(input.cc);
  if (input.bcc) out.bccRecipients = toRecipients(input.bcc);
  if (input.body !== undefined) {
    out.body = {
      contentType: input.bodyType === "html" ? "html" : "text",
      content: input.body,
    };
  }
  return out;
}

function toRecipients(addrs: string[]): Array<{ emailAddress: { address: string } }> {
  return addrs.map((address) => ({ emailAddress: { address } }));
}

function recipientsToEmails(arr: any): string[] {
  if (!Array.isArray(arr)) return [];
  return arr
    .map((r) => r?.emailAddress?.address)
    .filter((s): s is string => typeof s === "string");
}

function toSummary(m: any): DraftSummary {
  return {
    id: m.id,
    subject: m.subject ?? "",
    to: recipientsToEmails(m.toRecipients),
    cc: recipientsToEmails(m.ccRecipients),
    bcc: recipientsToEmails(m.bccRecipients),
    bodyPreview: m.bodyPreview,
    lastModifiedDateTime: m.lastModifiedDateTime,
    webLink: m.webLink,
  };
}

function toDetail(m: any): DraftDetail {
  const summary = toSummary(m);
  const contentType = (m.body?.contentType ?? "text").toLowerCase() === "html" ? "html" : "text";
  return {
    ...summary,
    bodyContentType: contentType,
    body: m.body?.content ?? "",
  };
}
