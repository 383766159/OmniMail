import type { Env } from '../../app/types'
import { ImapConnectionError } from '../../platform/imap/imap-errors'
import {
  createMicrosoftClient,
  type MicrosoftMailClient,
} from './microsoft-client'
import { microsoftAccessToken } from './microsoft-token-manager'
import { MicrosoftStoreError } from './microsoft-store'
import type { MicrosoftAccount } from './microsoft-types'

async function client(
  account: MicrosoftAccount,
  credential: string,
): Promise<MicrosoftMailClient> {
  return createMicrosoftClient(
    account.transport ?? 'imap',
    account.normalizedEmail,
    account.authMode,
    credential,
  )
}

export async function openMicrosoftClient(
  env: Env,
  account: MicrosoftAccount,
): Promise<MicrosoftMailClient> {
  if (account.authMode !== 'oauth2') {
    throw new MicrosoftStoreError(
      409, 'password_auth_removed', 'Microsoft 密码 LOGIN 已停用，请断开后使用 OAuth2 重新连接。',
    )
  }
  const credential = account.authMode === 'oauth2'
    ? await microsoftAccessToken(env, account)
    : account.password
  let remote = await client(account, credential)
  try {
    await remote.open()
    return remote
  } catch (error) {
    await remote.close()
    if (account.authMode !== 'oauth2'
      || !(error instanceof ImapConnectionError)
      || (error.status !== 400 && error.status !== 401)) {
      throw error
    }
    const refreshed = await microsoftAccessToken(env, account, { force: true })
    remote = await client(account, refreshed)
    try {
      await remote.open()
      return remote
    } catch (retryError) {
      await remote.close()
      throw retryError
    }
  }
}
