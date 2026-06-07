CREATE TABLE IF NOT EXISTS user_profiles (
  user_id BIGINT PRIMARY KEY,
  preferences JSONB NOT NULL DEFAULT '{}'::jsonb,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS subscriptions (
  user_id BIGINT PRIMARY KEY,
  is_active BOOLEAN NOT NULL DEFAULT TRUE,
  last_digest_at TIMESTAMPTZ NULL,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS interactions (
  id BIGSERIAL PRIMARY KEY,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  user_id BIGINT NOT NULL,
  action_type TEXT NOT NULL,
  listing_id TEXT NULL,
  context TEXT NULL
);

CREATE INDEX IF NOT EXISTS idx_interactions_user_created
  ON interactions (user_id, created_at DESC);
