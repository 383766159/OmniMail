import { ImapConnectionError } from '../../platform/imap/imap-errors'
import {
  microsoftAttachmentContent,
  parseMicrosoftMessage,
} from './microsoft-message-parser'
import type {
  MicrosoftAuthMode,
  MicrosoftFolder,
  MicrosoftMessageMetadata,
} from './microsoft-types'

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0'
const DEFAULT_DEADLINE_MS = 150_000
const PAGE_SIZE = 200
const PREVIEW_CHARS = 180
// mailFolder 的 wellKnownName 不是可选取属性（放进 $select 会 400），文件夹请求一律不带 $select。
const MESSAGE_SELECT = [
  'id', 'internetMessageId', 'from', 'toRecipients', 'ccRecipients', 'subject',
  'bodyPreview', 'receivedDateTime', 'sentDateTime', 'isRead', 'hasAttachments',
  'flag',
].join(',')
// Message 类型没有 size 属性（$select 会 400），因此 sizeBytes 固定记 0。
const MAX_LISTED_MESSAGES = 500
const MAX_RAW_BYTES = 10 * 1024 * 1024

type GraphAddress = { emailAddress?: { name?: string; address?: string } }

export type GraphMessage = {
  id: string
  internetMessageId?: string
  from?: GraphAddress
  toRecipients?: GraphAddress[]
  ccRecipients?: GraphAddress[]
  subject?: string
  bodyPreview?: string
  receivedDateTime?: string
  sentDateTime?: string
  isRead?: boolean
  hasAttachments?: boolean
  size?: number
  flag?: { flagStatus?: string }
}

type GraphFolder = {
  id: string
  displayName?: string
  wellKnownName?: string
  childFolderCount?: number
  totalItemCount?: number
}

// FNV-1a 64 位混合后截断到 52 位：Graph 消息 id 到数字 UID 的稳定、零依赖映射。
export function stableNumericId(value: string): number {
  let hash = 0xcbf29ce484222325n
  for (let index = 0; index < value.length; index += 1) {
    hash = ((hash ^ BigInt(value.charCodeAt(index))) * 0x100000001b3n) & 0xffffffffffffffffn
  }
  hash = ((hash ^ (hash >> 33n)) * 0xff51afd7ed558ccdn) & 0xffffffffffffffffn
  hash ^= hash >> 33n
  const truncated = Number(hash & 0xfffffffffffffn)
  return truncated > 0 ? truncated : 1
}

function addresses(values: GraphAddress[] | undefined): string[] {
  return (values || [])
    .map(({ emailAddress }) => (emailAddress?.address || '').toLowerCase())
    .filter(Boolean)
}

function seconds(value: string | undefined): number | null {
  const parsed = Date.parse(value || '')
  return Number.isFinite(parsed) ? Math.floor(parsed / 1000) : null
}

function textPreview(value: string | undefined): string {
  return (value || '').replace(/\s+/g, ' ').trim().slice(0, PREVIEW_CHARS)
}

export class MicrosoftGraphClient {
  private deadline = Date.now() + DEFAULT_DEADLINE_MS
  private folderId = ''
  private listing = new Map<number, GraphMessage>()

  constructor(
    private readonly email: string,
    private readonly authMode: MicrosoftAuthMode,
    private readonly credential: string,
  ) {}

  async open(): Promise<void> {
    // 只用 Mail 域端点做连通性校验：这批账号的授权 scope 不含 User.Read，
    // 访问 /me 会返回 401，会被误判成凭据问题（并显示成 IMAP 相关文案）。
    await this.send('/me/mailFolders?$top=1')
  }

  async close(): Promise<void> {
    return undefined
  }

  async listFolders(): Promise<MicrosoftFolder[]> {
    this.deadline = Date.now() + DEFAULT_DEADLINE_MS
    const roots = await this.collect<GraphFolder>(
      '/me/mailFolders?$top=100',
    )
    const folders = [...roots]
    for (const folder of roots) {
      if ((folder.childFolderCount || 0) < 1) continue
      folders.push(...await this.collect<GraphFolder>(
        `/me/mailFolders/${encodeURIComponent(folder.id)}/childFolders?$top=100`,
      ))
    }
    if (!folders.some(({ wellKnownName }) => wellKnownName === 'inbox')) {
      throw new ImapConnectionError(502, 'Microsoft Graph 未返回 INBOX 文件夹。', true)
    }
    return folders.map((folder) => ({
      path: folder.id,
      displayName: folder.displayName || '',
      specialUse: folder.wellKnownName || '',
      flags: folder.wellKnownName === 'deleteditems' ? ['\\Trash'] : [],
      uidValidity: 1,
      lastUid: 0,
    }))
  }

  async examineFolder(path: string): Promise<{ uidValidity: number; exists: number }> {
    if (path) this.folderId = path
    if (!this.folderId) {
      throw new ImapConnectionError(400, 'Microsoft Graph 文件夹无效。', true)
    }
    const folder = await this.send<GraphFolder>(
      `/me/mailFolders/${encodeURIComponent(this.folderId)}?$select=id,totalItemCount`,
    )
    return { uidValidity: 1, exists: Number(folder?.totalItemCount) || 0 }
  }

  async searchAllUids(): Promise<number[]> {
    if (!this.folderId) {
      throw new ImapConnectionError(400, 'Microsoft Graph 文件夹无效。', true)
    }
    const messages = await this.collect<GraphMessage>(
      `/me/mailFolders/${encodeURIComponent(this.folderId)}/messages`
        + `?$top=${PAGE_SIZE}&$orderby=receivedDateTime desc&$select=${MESSAGE_SELECT}`,
      MAX_LISTED_MESSAGES,
    )
    this.listing = new Map()
    for (const message of messages) this.listing.set(stableNumericId(message.id), message)
    // Graph 按收件时间倒序返回，反转后与 IMAP UID SEARCH ALL 的升序语义一致。
    return [...this.listing.keys()].reverse()
  }

  async fetchMetadata(
    uids: number[],
    deadline = this.deadline,
  ): Promise<MicrosoftMessageMetadata[]> {
    const selected = [...new Set(uids)]
      .filter((uid) => Number.isSafeInteger(uid) && uid > 0)
      .sort((left, right) => left - right)
    const messages: MicrosoftMessageMetadata[] = []
    for (const uid of selected) {
      if (Date.now() >= deadline) {
        throw new ImapConnectionError(504, 'Microsoft 邮箱同步超过总执行时间上限。')
      }
      const message = this.listing.get(uid)
      if (!message) continue
      messages.push(this.metadata(uid, message))
    }
    return messages
  }

  async getMessage(folderPath: string, uid: number, remoteId?: string) {
    const messageId = await this.resolveRemoteId(folderPath, uid, remoteId)
    const response = await this.raw(
      `/me/messages/${encodeURIComponent(messageId)}/$value`,
    )
    const declared = Number(response.headers.get('content-length') || 0)
    if (declared > MAX_RAW_BYTES) {
      throw new ImapConnectionError(413, 'Microsoft 邮件超过 10 MiB 上限。', true)
    }
    const bytes = new Uint8Array(await response.arrayBuffer())
    if (bytes.length > MAX_RAW_BYTES) {
      throw new ImapConnectionError(413, 'Microsoft 邮件超过 10 MiB 上限。', true)
    }
    return parseMicrosoftMessage(bytes, String(uid))
  }

  async markSeen(
    folderPath: string,
    uid: number,
    expectedUidValidity: number,
    remoteId?: string,
  ): Promise<void> {
    if (expectedUidValidity !== 1) {
      throw new ImapConnectionError(
        404,
        'Microsoft 文件夹 UIDVALIDITY 已变化，请刷新邮件列表。',
        true,
      )
    }
    const messageId = await this.resolveRemoteId(folderPath, uid, remoteId)
    await this.send(`/me/messages/${encodeURIComponent(messageId)}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ isRead: true }),
    })
  }

  async getAttachment(folderPath: string, uid: number, partId: string, remoteId?: string) {
    const { parsedAttachments } = await this.getMessage(folderPath, uid, remoteId)
    return microsoftAttachmentContent(parsedAttachments, partId)
  }

  private metadata(uid: number, message: GraphMessage): MicrosoftMessageMetadata {
    const isRead = message.isRead === true
    return {
      uid,
      remoteId: message.id,
      internetMessageId: message.internetMessageId || '',
      senderName: message.from?.emailAddress?.name || '',
      senderAddress: (message.from?.emailAddress?.address || '').toLowerCase(),
      recipients: addresses(message.toRecipients),
      cc: addresses(message.ccRecipients),
      subject: message.subject || '',
      preview: textPreview(message.bodyPreview),
      receivedAt: seconds(message.receivedDateTime) ?? Math.floor(Date.now() / 1000),
      sentAt: seconds(message.sentDateTime),
      sizeBytes: 0,
      flags: isRead ? ['\\Seen'] : [],
      isRead,
      isStarred: message.flag?.flagStatus === 'flagged',
      hasAttachments: message.hasAttachments === true,
    }
  }

  private async resolveRemoteId(
    folderPath: string,
    uid: number,
    remoteId?: string,
  ): Promise<string> {
    if (!Number.isSafeInteger(uid) || uid < 1) {
      throw new ImapConnectionError(400, 'Microsoft 邮件 UID 无效。', true)
    }
    if (remoteId) return remoteId
    const known = this.listing.get(uid)
    if (known) return known.id
    await this.examineFolder(folderPath)
    await this.searchAllUids()
    const found = this.listing.get(uid)
    if (!found) throw new ImapConnectionError(404, 'Microsoft 邮件不存在。', true)
    return found.id
  }

  private async collect<T>(path: string, maximum = 0): Promise<T[]> {
    const items: T[] = []
    let next: string | undefined = path
    for (let page = 0; next && page < 20; page++) {
      const body: { value?: T[], '@odata.nextLink'?: string } =
        await this.send<{ value?: T[], '@odata.nextLink'?: string }>(next)
      if (Array.isArray(body?.value)) items.push(...body.value)
      const link: string | undefined = body ? body['@odata.nextLink'] : undefined
      next = typeof link === 'string' && link ? link : undefined
      if (maximum > 0 && items.length >= maximum) break
    }
    return maximum > 0 ? items.slice(0, maximum) : items
  }

  private async raw(path: string): Promise<Response> {
    return this.request(path, { headers: {} })
  }

  private async send<T>(path: string, init: {
    method?: string
    headers?: Record<string, string>
    body?: string
  } = {}): Promise<T> {
    const response = await this.request(path, {
      method: init.method || 'GET',
      headers: { Accept: 'application/json', ...init.headers },
      body: init.body,
    })
    if (init.method === 'PATCH' || response.status === 204) return undefined as T
    try {
      return await response.json<T>()
    } catch {
      return undefined as T
    }
  }

  private async request(path: string, init: {
    method?: string
    headers?: Record<string, string>
    body?: string
  }): Promise<Response> {
    const remaining = this.deadline - Date.now()
    if (remaining <= 0) {
      throw new ImapConnectionError(504, 'Microsoft 邮箱同步超过总执行时间上限。')
    }
    let response: Response
    try {
      response = await fetch(path.startsWith('http') ? path : `${GRAPH_BASE}${path}`, {
        method: init.method || 'GET',
        headers: { Authorization: `Bearer ${this.credential}`, ...init.headers },
        body: init.body,
        signal: AbortSignal.timeout(remaining),
      })
    } catch (error) {
      if (error instanceof Error && error.name === 'TimeoutError') {
        throw new ImapConnectionError(504, 'Microsoft 邮箱同步超过总执行时间上限。')
      }
      throw new ImapConnectionError(502, 'Microsoft Graph 暂时无法连接。', true)
    }
    if (!response.ok) throw await this.failure(response)
    return response
  }

  private async failure(response: Response): Promise<ImapConnectionError> {
    const status = response.status
    if (status === 401 || status === 403) {
      return new ImapConnectionError(401, `Microsoft Graph 拒绝访问（${status}）。`, true)
    }
    if (status === 404) return new ImapConnectionError(404, 'Microsoft 邮件不存在。', true)
    if (status === 429 || status >= 500) {
      return new ImapConnectionError(502, `Microsoft Graph 暂时不可用（${status}）。`, true)
    }
    return new ImapConnectionError(400, `Microsoft Graph 拒绝了请求（${status}）。`, true)
  }
}
