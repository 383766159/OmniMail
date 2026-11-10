export type MicrosoftTransport = 'imap' | 'graph'

export const MICROSOFT_IMAP_SCOPE = 'https://outlook.office.com/IMAP.AccessAsUser.All'
export const MICROSOFT_GRAPH_SCOPE = 'https://graph.microsoft.com/.default'

export function parseMicrosoftTransport(value: unknown): MicrosoftTransport {
  return typeof value === 'string' && value.trim().toLowerCase() === 'graph'
    ? 'graph' : 'imap'
}

export function microsoftTokenScope(transport: MicrosoftTransport): string {
  return `${transport === 'graph' ? MICROSOFT_GRAPH_SCOPE : MICROSOFT_IMAP_SCOPE} offline_access`
}

export function microsoftScopeSatisfied(scope: string, transport: MicrosoftTransport): boolean {
  if (!scope) return true
  const wanted = transport === 'graph' ? MICROSOFT_GRAPH_SCOPE : MICROSOFT_IMAP_SCOPE
  return scope.split(/\s+/).some((item) => item.toLowerCase() === wanted.toLowerCase())
}
