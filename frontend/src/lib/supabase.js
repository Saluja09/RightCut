import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || ''
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || ''

// Validate: URL must be https and key must be present (JWT eyJ... or new sb_publishable_...)
const isValidConfig = SUPABASE_URL.startsWith('https://') &&
  SUPABASE_ANON_KEY.length > 20

// If no valid Supabase config, export a null client — app runs in offline/guest mode
export const supabase = isValidConfig
  ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
  : null

export const supabaseEnabled = isValidConfig && !!supabase
