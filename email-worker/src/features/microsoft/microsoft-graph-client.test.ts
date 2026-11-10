import { afterEach, describe, expect, it, vi } from 'vitest'
import { MicrosoftGraphClient, stableNumericId } from './microsoft-graph-client'

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  })
}

function client(): MicrosoftGraphClient {
  return new MicrosoftGraphClient('user@outlook.com', 'oauth2', 'access-token')
}

describe('Microsoft Graph transport', () => {
  afterEach(() => vi.unstubAllGlobals())

  it('maps Graph folders, flags, and child folders', async () => {
    const fetchMock = vi.fn(async (url: string) => {
      if (url.includes('/childFolders')) {
        return jsonResponse({
          value: [{ id: 'child', displayName: 'Child', wellKnownName: null, childFolderCount: 0 }],
        })
      }
      return jsonResponse({
        value: [
          { id: 'inbox', displayName: 'Inbox', wellKnownName: 'inbox', childFolderCount: 1 },
          { id: 'trash', displayName: 'Deleted', wellKnownName: 'deleteditems', childFolderCount: 0 },
        ],
      })
    })
    vi.stubGlobal('fetch', fetchMock)
    const folders = await client().listFolders()
    expect(folders.map(({ path }) => path)).toEqual(['inbox', 'trash', 'child'])
    expect(folders[1]).toMatchObject({ flags: ['\\Trash'], uidValidity: 1, lastUid: 0 })
    expect(folders[0]).toMatchObject({ specialUse: 'inbox', flags: [] })
    expect(fetchMock.mock.calls[0][0]).toContain('graph.microsoft.com/v1.0/me/mailFolders')
  })

  it('requires a well-known inbox folder', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => jsonResponse({
      value: [{ id: 'archive', displayName: 'Archive', wellKnownName: 'archive', childFolderCount: 0 }],
    })))
    await expect(client().listFolders()).rejects.toMatchObject({ status: 502 })
  })

  it('lists messages and maps metadata with remote ids', async () => {
    vi.stubGlobal('fetch', vi.fn(async (url: string) => {
      if (url.includes('/messages?')) {
        return jsonResponse({ value: [{
          id: 'remote-new',
          internetMessageId: '<new@example.com>',
          from: { emailAddress: { name: 'Sender', address: 'Sender@Example.com' } },
          toRecipients: [{ emailAddress: { address: 'User@Outlook.com' } }],
          ccRecipients: [],
          subject: 'Hello',
          bodyPreview: '  many   spaces  ',
          receivedDateTime: '2026-08-25T10:00:00.000Z',
          sentDateTime: '2026-08-25T09:59:00.000Z',
          isRead: true,
          hasAttachments: true,
          size: 2048,
          flag: { flagStatus: 'flagged' },
        }, {
          id: 'remote-old',
          from: { emailAddress: { address: 'old@example.com' } },
          subject: 'Older',
          receivedDateTime: '2026-08-24T10:00:00.000Z',
          isRead: false,
        }] })
      }
      return jsonResponse({ id: 'inbox', totalItemCount: 2 })
    }))
    const graph = client()
    await graph.examineFolder('inbox')
    const uids = await graph.searchAllUids()
    expect(uids).toHaveLength(2)
    const metadata = await graph.fetchMetadata(uids)
    const newest = metadata.find(({ remoteId }) => remoteId === 'remote-new')
    expect(newest).toMatchObject({
      uid: stableNumericId('remote-new'),
      senderAddress: 'sender@example.com',
      recipients: ['user@outlook.com'],
      preview: 'many spaces',
      sentAt: 1787651940,
      receivedAt: 1787652000,
      // Graph 的 Message 没有 size 属性（$select 会 400），因此 sizeBytes 固定记 0。
      sizeBytes: 0,
      isRead: true,
      isStarred: true,
      hasAttachments: true,
      flags: ['\\Seen'],
    })
    expect(metadata.find(({ remoteId }) => remoteId === 'remote-old'))
      .toMatchObject({ isRead: false, flags: [] })
  })

  it('maps rejected tokens to the existing authorization error', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => jsonResponse({ error: { code: 'InvalidAuthenticationToken' } }, 401)))
    await expect(client().open()).rejects.toMatchObject({
      status: 401, message: 'Microsoft Graph 拒绝访问（401）。', definitive: true,
    })
  })

  it('maps Graph outages to a retryable gateway error', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => jsonResponse({ error: {} }, 503)))
    await expect(client().open()).rejects.toMatchObject({ status: 502, definitive: true })
  })

  it('patches isRead through the explicit remote id', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ id: 'remote-1', isRead: true }))
    vi.stubGlobal('fetch', fetchMock)
    const uid = stableNumericId('remote-1')
    await client().markSeen('inbox', uid, 1, 'remote-1')
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe('https://graph.microsoft.com/v1.0/me/messages/remote-1')
    expect(init.method).toBe('PATCH')
    expect(init.body).toBe('{"isRead":true}')
    expect(init.headers).toMatchObject({ 'Content-Type': 'application/json' })
  })

  it('rejects marking seen when the folder identity changed', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({}))
    vi.stubGlobal('fetch', fetchMock)
    await expect(client().markSeen('inbox', 7, 42, 'remote-1')).rejects.toMatchObject({ status: 404 })
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('derives stable 52-bit numeric ids from Graph message ids', () => {
    expect(stableNumericId('AAMkAGI2')).toBe(stableNumericId('AAMkAGI2'))
    expect(stableNumericId('AAMkAGI2')).not.toBe(stableNumericId('AAMkAGI3'))
    for (const value of ['a', 'remote-1', 'AAMkADU5ZmM2LT']) {
      const uid = stableNumericId(value)
      expect(Number.isSafeInteger(uid)).toBe(true)
      expect(uid).toBeGreaterThan(0)
      expect(uid).toBeLessThanOrEqual(2 ** 52 - 1)
    }
  })
})
