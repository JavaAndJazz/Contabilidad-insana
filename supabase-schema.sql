-- Ejecuta esto UNA SOLA VEZ en Supabase: tu proyecto → SQL Editor → New query → pega y "Run".

create table if not exists app_data (
  key text primary key,
  value jsonb,
  updated_at timestamptz default now()
);

-- Seguridad a nivel de fila (obligatoria en Supabase para poder leer/escribir).
alter table app_data enable row level security;

-- Política simple: cualquiera con la URL y la clave pública (anon key) puede leer y escribir.
-- Suficiente para una herramienta interna de uso privado. Si más adelante quieres
-- exigir inicio de sesión (usuario/contraseña) para entrar al sistema, se puede
-- reemplazar esta política por una que exija auth.uid() y usar Supabase Auth.
create policy "Acceso interno (anon) a app_data"
  on app_data
  for all
  using (true)
  with check (true);
