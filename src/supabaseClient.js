import { createClient } from '@supabase/supabase-js'

// Vite legge le variabili d'ambiente usando import.meta.env
const supabaseUrl = import.meta.env.VITE_SUPABASE_URL
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY

if (!supabaseUrl || !supabaseAnonKey) {
  console.error("Attenzione: Chiavi Supabase mancanti nel file .env.local!")
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey)