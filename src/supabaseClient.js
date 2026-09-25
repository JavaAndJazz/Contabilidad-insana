import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // eslint-disable-next-line no-console
  console.error(
    'Faltan las variables VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY. ' +
    'Revisa tu archivo .env (copia .env.example) o las variables de entorno en Vercel/Netlify.'
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

// Tabla usada: app_data (key text primary key, value jsonb, updated_at timestamptz)
// Ver supabase-schema.sql para crearla.

export async function kvGet(key) {
  const { data, error } = await supabase
    .from('app_data')
    .select('value')
    .eq('key', key)
    .maybeSingle();

  if (error) throw error;
  return data ? data.value : null;
}

export async function kvSet(key, value) {
  const { error } = await supabase
    .from('app_data')
    .upsert({ key, value, updated_at: new Date().toISOString() });

  if (error) throw error;
}
