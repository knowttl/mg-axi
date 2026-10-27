import { constants } from "node:fs";
import { lstat, open, unlink } from "node:fs/promises";
import { isAbsolute } from "node:path";
import { createHash } from "node:crypto";
import { AxiError } from "axi-sdk-js";
import { encodeGraphPathSegment } from "./graph-session.js";
import { validateMailbox } from "./mail-authorization.js";
import type { MutationCoordinator, DownloadDefinition } from "./mutations.js";

export const MAIL_DOWNLOAD_ALTERNATES = {
  "GET:/users/{user-id}/messages/{message-id}/$value": "GET:/me/messages/{message-id}/$value",
  "GET:/users/{user-id}/messages/{message-id}/attachments/{attachment-id}/$value": {
    me: "GET:/me/messages/{message-id}/attachments/{attachment-id}/$value",
    metadata: "GET:/users/{user-id}/messages/{message-id}/attachments/{attachment-id}",
    meMetadata: "GET:/me/messages/{message-id}/attachments/{attachment-id}",
  },
};
export const MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024;
export type BinaryTransport = (request: { method: "GET"; url: string; headers: Record<string, string>; signal: AbortSignal }) => Promise<{
  status: number; headers: Record<string, string>; body: AsyncIterable<Uint8Array>;
}>;
export const binaryFetchTransport: BinaryTransport = async request => {
  const response = await fetch(request.url, { ...request, redirect: "manual" });
  return { status: response.status, headers: Object.fromEntries(response.headers), body: response.body ?? (async function* () {})() };
};

// Called only after the coordinator records intent. Neither Graph filenames nor
// response URLs select the local path; bytes never enter structured output.
export async function saveMailBytes(args: { path: string; url: string; token: string; attachment: boolean; transport: BinaryTransport }): Promise<Record<string, unknown>> {
  let file;
  try { file = await open(args.path, constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW, 0o600); }
  catch { throw new AxiError("--out must name a new writable file", "VALIDATION_ERROR"); }
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 30000);
  const request = async (url: string, cap: number) => {
    const response = await args.transport({ method: "GET", url, headers: { Authorization: `Bearer ${args.token}`, Accept: "*/*" }, signal: controller.signal });
    if (response.status !== 200) throw new AxiError(`Graph download returned status ${response.status}; request was not retried`, "GRAPH_ERROR");
    const length = response.headers["content-length"];
    if (length !== undefined && (!/^\d+$/.test(length) || !Number.isSafeInteger(Number(length)) || Number(length) > cap)) throw new AxiError("Reported download size exceeds the allowed cap or is invalid", "DOWNLOAD_TOO_LARGE");
    return response;
  };
  try {
    await file.chmod(0o600);
    if (args.attachment) {
      const response = await request(`${args.url.slice(0, -7)}?$select=id,size`, 65536);
      const chunks: Uint8Array[] = [];
      let size = 0;
      for await (const chunk of response.body) {
        controller.signal.throwIfAborted();
        size += chunk.byteLength;
        if (size > 65536) throw new Error();
        chunks.push(chunk);
      }
      const metadata = JSON.parse(Buffer.concat(chunks).toString("utf8")) as Record<string, unknown>;
      if (!["#microsoft.graph.fileAttachment", "#microsoft.graph.itemAttachment"].includes(String(metadata["@odata.type"]))) throw new AxiError("Reference or unknown attachment types cannot be saved", "POLICY_DENIED");
      if (!Number.isSafeInteger(metadata.size) || Number(metadata.size) < 0 || Number(metadata.size) > MAX_DOWNLOAD_BYTES) throw new AxiError("Reported attachment size exceeds 50 MiB or is invalid", "DOWNLOAD_TOO_LARGE");
    }
    const response = await request(args.url, MAX_DOWNLOAD_BYTES);
    const hash = createHash("sha256");
    let size = 0;
    for await (const chunk of response.body) {
      controller.signal.throwIfAborted();
      size += chunk.byteLength;
      if (size > MAX_DOWNLOAD_BYTES) throw new AxiError("Download exceeds 50 MiB", "DOWNLOAD_TOO_LARGE");
      hash.update(chunk);
      let offset = 0;
      while (offset < chunk.byteLength) {
        const written = await file.write(chunk, offset, chunk.byteLength - offset);
        if (!written.bytesWritten) throw new Error();
        offset += written.bytesWritten;
      }
    }
    controller.signal.throwIfAborted();
    if (response.headers["content-length"] !== undefined && Number(response.headers["content-length"]) !== size) throw new AxiError("Download length does not match the reported size", "GRAPH_ERROR");
    await file.sync();
    await file.close();
    return { path: args.path, size, sha256: hash.digest("hex"), contentType: response.headers["content-type"]?.split(";")[0]?.match(/^[\w.+-]+\/[\w.+-]+$/)?.[0] ?? "application/octet-stream", kind: args.attachment ? "attachment" : "mime" };
  } catch (error) {
    controller.abort();
    await file.close().catch(() => {});
    await unlink(args.path);
    if (error instanceof AxiError) throw error;
    throw new AxiError("Download failed; partial file removed; request was not retried", "GRAPH_ERROR");
  } finally { clearTimeout(timer); controller.abort(); }
}

export async function downloadMail(coordinator: MutationCoordinator, flags: Record<string, string | boolean>, attachment: boolean, help: string): Promise<Record<string, unknown>> {
  const mailbox = String(flags.mailbox), id = String(flags.id), attachmentId = String(flags.attachment);
  validateMailbox(mailbox, help);
  for (const value of attachment ? [id, attachmentId] : [id]) {
    if (!value || /[\x00-\x20\x7f/\\?#%$]/.test(value) || value === "." || value === "..") throw new AxiError("Invalid mail resource ID", "VALIDATION_ERROR", [help]);
  }
  const path = String(flags.out);
  if (!isAbsolute(path) || /[\x00-\x1f\x7f]/.test(path)) throw new AxiError("--out requires an absolute new file path", "VALIDATION_ERROR", [help]);
  try { await lstat(path); throw new AxiError("--out already exists", "VALIDATION_ERROR", [help]); }
  catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error; }
  const target = attachment ? attachmentId : id;
  if (flags.execute === true && flags.confirm !== target) throw new AxiError("--execute requires --confirm repeating the exact resource ID", flags.confirm === undefined ? "CONFIRM_REQUIRED" : "CONFIRM_MISMATCH", [help]);
  const root = mailbox === "me" ? "me" : `users/${encodeGraphPathSegment(mailbox)}`;
  const definition: DownloadDefinition = { operation: attachment ? "mail.attachment.save" : "mail.message.export", method: "GET", version: "v1.0", effect: "disruptive", target,
    path: `/v1.0/${root}/messages/${encodeGraphPathSegment(id)}${attachment ? `/attachments/${encodeGraphPathSegment(attachmentId)}` : ""}/$value` };
  return coordinator.download(definition, { execute: flags.execute === true, confirm: flags.confirm as string | undefined, out: path });
}
