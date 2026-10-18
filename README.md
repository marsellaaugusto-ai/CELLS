# Cells

Hoja de cálculo online, con fórmulas en español y sincronización con otras apps.

## Qué hace hoy

- Hoja editable con selección, arrastre, copiar/cortar/pegar (compatible con Excel y Google Sheets) y deshacer/rehacer.
- Más de 250 funciones. Acepta nombres en español (`SUMA`, `SI`, `BUSCARV`…) o en inglés, y `;` o `,` como separador.
- Referencias a otras hojas (`'Hoja 2'!A1`), detección de referencias circulares y errores explicados al pasar el mouse.
- Formato: negrita, cursiva, alineación, moneda, porcentaje, decimales y color de celda.
- Varias hojas por libro (doble clic para renombrar, clic derecho para borrar).
- Gráfico de barras de la selección.
- Importar y exportar CSV, exportar JSON.
- Guardado automático: en el navegador, o en la nube con login por email si configurás Supabase.
- En la nube: sincronización en vivo entre dispositivos y una API REST para conectar otras apps.

## Correrlo en tu computadora

```bash
npm install
npm run dev        # abre http://localhost:5173
npm test           # pruebas del motor de fórmulas
npm run build      # genera la versión publicable en dist/
```

## Guardar en la nube (gratis, opcional)

1. Creá un proyecto en [supabase.com](https://supabase.com) (el plan gratuito alcanza).
2. En **SQL Editor**, pegá y ejecutá el contenido de `supabase/schema.sql`.
3. En **Authentication → URL Configuration**, agregá la dirección donde publiques Cells a **Redirect URLs**.
4. Copiá `.env.example` como `.env` y completá `VITE_SUPABASE_URL` y `VITE_SUPABASE_ANON_KEY`
   (están en **Project Settings → API**).

Sin estas variables, Cells funciona igual y guarda en el navegador.

## Publicarlo en internet (gratis)

El repositorio es privado, así que GitHub Pages no sirve en el plan gratuito. Usá Vercel, Netlify o Cloudflare Pages:

1. Entrá con tu cuenta de GitHub y elegí el repositorio `Cells`.
2. Comando de build: `npm run build`. Carpeta de salida: `dist`.
3. Si usás Supabase, cargá las dos variables `VITE_SUPABASE_…` en la configuración del proyecto.

Cada push a `main` vuelve a publicar la app.

## Conectar otras apps

Con la nube activada, cada libro tiene una dirección REST (la ves en **Conectar**):

```
GET  {SUPABASE_URL}/rest/v1/workbooks?id=eq.{ID}&select=data
```

Desde n8n, Make, Zapier o tu propio código se lee y escribe con la clave pública de Supabase
y el token de un usuario con acceso. Las apps abiertas reciben los cambios en vivo.

## Estructura

| Carpeta | Contenido |
|---|---|
| `src/engine` | Motor de fórmulas, traducción del español y formato de números |
| `src/state` | Cambios al libro y deshacer/rehacer |
| `src/storage` | Guardado en el navegador o en Supabase |
| `src/components` | Interfaz: menú lateral, paneles, cuadrícula y gráfico |
| `supabase/schema.sql` | Tabla y permisos de la base de datos |

## Licencia

© 2026 [TU NOMBRE]. Todos los derechos reservados. Ver `LICENSE` y las licencias de terceros en `THIRD_PARTY.md`.
