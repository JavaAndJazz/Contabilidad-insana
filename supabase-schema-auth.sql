-- Ejecuta esto UNA SOLA VEZ en Supabase (SQL Editor → New query → pegar → Run).
-- Reemplaza el acceso abierto por uno que exige haber iniciado sesión.

drop policy if exists "Acceso interno (anon) a app_data" on app_data;

create policy "Solo usuarios con sesión iniciada"
  on app_data
  for all
  using (auth.role() = 'authenticated')
  with check (auth.role() = 'authenticated');
