import type { MicrosoftGraphClient } from './microsoft-graph-client'
import type { MicrosoftImapClient } from './microsoft-imap'
import type { MicrosoftTransport } from './microsoft-transport'
import type { MicrosoftAuthMode } from './microsoft-types'

export type MicrosoftMailClient = MicrosoftImapClient | MicrosoftGraphClient

export async function createMicrosoftClient(
  transport: MicrosoftTransport,
  email: string,
  authMode: MicrosoftAuthMode,
  credential: string,
): Promise<MicrosoftMailClient> {
  if (transport === 'graph') {
    const { MicrosoftGraphClient } = await import('./microsoft-graph-client')
    return new MicrosoftGraphClient(email, authMode, credential)
  }
  const { MicrosoftImapClient } = await import('./microsoft-imap')
  return new MicrosoftImapClient(email, authMode, credential)
}
