import type { MicrosoftTransport } from '../../../shared/api'
import { t } from '../../../shared/i18n'
import '../styles/microsoft-transport.css'

export function MicrosoftTransportSelector({ value, onChange, disabled = false }: {
  value: MicrosoftTransport
  onChange: (value: MicrosoftTransport) => void
  disabled?: boolean
}) {
  const options: Array<{ value: MicrosoftTransport; label: string }> = [
    { value: 'imap', label: t('IMAP（默认）') },
    { value: 'graph', label: t('Graph（IMAP 被封时使用）') },
  ]
  return <div className="microsoft-transport-field">
    <span>{t('传输方式')}</span>
    <div className="microsoft-transport-options" role="radiogroup" aria-label={t('传输方式')}>
      {options.map((option) => <button key={option.value} type="button" role="radio"
        className={option.value === value ? 'is-selected' : ''}
        aria-checked={option.value === value} disabled={disabled}
        onClick={() => onChange(option.value)}>{option.label}</button>)}
    </div>
  </div>
}
