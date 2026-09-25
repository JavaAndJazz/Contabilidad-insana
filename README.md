# Contabilidad NIIF — Grupo Procreas 929, C.A.

Sistema contable con Libro Diario, Libro Mayor, Balance de Comprobación,
Estado de Resultados, Estado de Situación Financiera e Inventario.
Exporta cada reporte a Excel real (.xlsx) y guarda todo en una base de
datos en la nube (Supabase), así que se puede usar desde cualquier
computadora y los datos son los mismos para todos.

## Qué necesitas (todo gratis)

1. Una cuenta en **[Supabase](https://supabase.com)** (la base de datos)
2. Una cuenta en **[GitHub](https://github.com)** (para guardar el código)
3. Una cuenta en **[Vercel](https://vercel.com)** o **[Netlify](https://netlify.com)** (para publicar la página web)

---

## Paso 1 — Crear la base de datos en Supabase

1. Entra a [supabase.com](https://supabase.com), crea una cuenta gratis y crea un **New Project**.
2. Ponle un nombre (ej. `procreas-contabilidad`) y una contraseña de base de datos (guárdala, no la necesitarás para esto pero por si acaso).
3. Espera a que el proyecto termine de crearse (1-2 minutos).
4. Ve a **SQL Editor** (menú izquierdo) → **New query**.
5. Abre el archivo `supabase-schema.sql` de esta carpeta, copia todo su contenido, pégalo ahí y dale **Run**. Esto crea la tabla donde se guardan tus datos.
6. Ve a **Settings** (ícono de engranaje) → **API**. Ahí vas a ver dos datos que necesitas copiar:
   - **Project URL** (algo como `https://xxxxx.supabase.co`)
   - **anon public key** (una clave larga)

Guárdalos, los usas en el Paso 3.

## Paso 2 — Subir el código a GitHub

1. Entra a [github.com](https://github.com) y crea una cuenta si no tienes.
2. Crea un repositorio nuevo (botón verde **New**), ponle nombre (ej. `procreas-contabilidad`), déjalo en **Private** si no quieres que sea público, y **Create repository**.
3. Sube todos los archivos de esta carpeta a ese repositorio. La forma más fácil sin usar la terminal:
   - En la página de tu repositorio recién creado, haz clic en **uploading an existing file**.
   - Arrastra **todos** los archivos y carpetas de este proyecto (menos `node_modules` si llegaras a tenerla, no debería existir aún).
   - Dale **Commit changes**.

(Si sabes usar git/terminal, es simplemente `git init`, `git add .`, `git commit -m "inicial"`, y conectar con tu repo de GitHub — más rápido.)

## Paso 3 — Publicar en Vercel

1. Entra a [vercel.com](https://vercel.com) y crea una cuenta (puedes entrar directo con tu cuenta de GitHub, es lo más fácil).
2. **Add New → Project**.
3. Selecciona el repositorio `procreas-contabilidad` que acabas de subir.
4. Vercel detecta automáticamente que es un proyecto Vite. Antes de darle a "Deploy", abre la sección **Environment Variables** y agrega estas dos (con los valores que copiaste en el Paso 1):

   | Name | Value |
   |---|---|
   | `VITE_SUPABASE_URL` | tu Project URL de Supabase |
   | `VITE_SUPABASE_ANON_KEY` | tu anon public key de Supabase |

5. Dale **Deploy**. En 1-2 minutos te da un link (algo como `procreas-contabilidad.vercel.app`).

Ese link es tu página web. Se lo mandas a tu contador y ya puede trabajar — desde su computadora, la tuya, o el celular, todos ven y guardan sobre los mismos datos.

*(Con Netlify es prácticamente el mismo proceso: conectas el repo de GitHub, agregas las mismas dos variables de entorno en Site settings → Environment variables, y publicas.)*

---

## Cómo actualizar la página más adelante

Si en el futuro me pides cambios al sistema, yo te entrego el archivo `App.jsx` actualizado. Solo tienes que reemplazar ese archivo en tu repositorio de GitHub (subirlo de nuevo pisando el anterior) y Vercel/Netlify vuelve a publicar la página automáticamente en un par de minutos. No hay que tocar nada más.

## Notas importantes

- **Seguridad**: cualquiera que tenga el link de la página puede entrar y ver/editar los datos (no hay usuario/contraseña todavía). Está bien para uso interno, pero si quieres restringir el acceso con inicio de sesión, se puede agregar más adelante (Supabase Auth).
- **Sincronización**: la app revisa la base de datos cada 20 segundos para reflejar cambios de otros dispositivos. Si dos personas editan exactamente al mismo tiempo, gana el último guardado (no hay fusión automática de cambios).
- **Desarrollo local** (opcional, solo si quieres probarlo en tu computadora antes de publicar): necesitas [Node.js](https://nodejs.org) instalado, luego en esta carpeta:
  ```
  npm install
  cp .env.example .env   # y pon ahí tus datos de Supabase
  npm run dev
  ```
