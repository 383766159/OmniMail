ALTER TABLE microsoft_imap_accounts ADD COLUMN transport TEXT NOT NULL DEFAULT 'imap';
ALTER TABLE microsoft_imap_messages ADD COLUMN remote_id TEXT NOT NULL DEFAULT '';
CREATE INDEX idx_microsoft_imap_messages_remote
  ON microsoft_imap_messages(account_id, folder_path, remote_id);
