#!/usr/bin/env node

console.log(`
╔════════════════════════════════════════════════════════════════╗
║         GUÍA DE DEPLOYMENT - SNOWFLAKE EXCEL PLUGIN           ║
╚════════════════════════════════════════════════════════════════╝

📦 Tu aplicación tiene 2 partes que necesitas desplegar:

┌────────────────────────────────────────────────────────────────┐
│ 1. BACKEND (OAuth Server)                                     │
│    - Express.js con rutas OAuth                               │
│    - Se despliega en: Vercel o Railway                        │
└────────────────────────────────────────────────────────────────┘

┌────────────────────────────────────────────────────────────────┐
│ 2. FRONTEND (Taskpane)                                        │
│    - HTML, CSS, JavaScript                                    │
│    - Se despliega en: GitHub Pages                            │
└────────────────────────────────────────────────────────────────┘

🚀 PASOS RÁPIDOS:

══════════════════════════════════════════════════════════════════
PASO 1: Desplegar BACKEND en Vercel
══════════════════════════════════════════════════════════════════

1️⃣  Instala Vercel CLI:
    npm install -g vercel

2️⃣  Login en Vercel:
    vercel login

3️⃣  Despliega:
    vercel

4️⃣  Configura variables de entorno en Vercel:
    vercel env add ALLOWED_ORIGINS

    Valor: https://bit-quill.github.io

5️⃣  Deploy a producción:
    vercel --prod

6️⃣  Guarda la URL que te da Vercel!
    Ejemplo: https://snowflake-excel-oauth.vercel.app

══════════════════════════════════════════════════════════════════
PASO 2: Actualizar configuración del Frontend
══════════════════════════════════════════════════════════════════

7️⃣  Crea archivo src/config.js con la URL de tu backend:

    export default {
      API_BASE_URL: 'https://TU-APP.vercel.app'
    };

    Reemplaza TU-APP con tu URL de Vercel del paso 6.

══════════════════════════════════════════════════════════════════
PASO 3: Desplegar FRONTEND en GitHub Pages
══════════════════════════════════════════════════════════════════

8️⃣  Haz build del proyecto:
    npm run build

9️⃣  Haz commit de los cambios:
    git add .
    git commit -m "Deploy to production"
    git push origin main

🔟  GitHub Actions desplegará automáticamente!
    Ve el progreso en: Actions tab en GitHub

1️⃣1️⃣  Configura GitHub Pages (primera vez):
    - Ve a: Settings > Pages
    - Source: GitHub Actions
    - Guarda

══════════════════════════════════════════════════════════════════
PASO 4: Configurar Snowflake
══════════════════════════════════════════════════════════════════

1️⃣2️⃣  Crea Security Integration en Snowflake:

    CREATE SECURITY INTEGRATION excel_oauth_prod
      TYPE = OAUTH
      ENABLED = TRUE
      OAUTH_CLIENT = PUBLIC
      OAUTH_REDIRECT_URI = 'https://bit-quill.github.io/ssvp-hosted/auth/callback'
      OAUTH_ISSUE_REFRESH_TOKENS = TRUE
      OAUTH_REFRESH_TOKEN_VALIDITY = 7776000;

1️⃣3️⃣  Obtén el Client ID:

    DESCRIBE SECURITY INTEGRATION excel_oauth_prod;

    Copia el OAUTH_CLIENT_ID.

══════════════════════════════════════════════════════════════════
✅ LISTO! Ahora puedes usar tu aplicación:
══════════════════════════════════════════════════════════════════

🌐 URL del Frontend:
   https://bit-quill.github.io/ssvp-hosted/

🔧 URL del Backend:
   https://TU-APP.vercel.app

📖 Para más detalles, lee DEPLOYMENT.md

═══════════════════════════════════════════════════════════════════

🆘 ¿Necesitas ayuda?

  • Revisa DEPLOYMENT.md para la guía completa
  • Verifica los logs en Vercel Dashboard
  • Revisa GitHub Actions para el estado del deploy
  • Abre la consola del navegador (F12) para errores

═══════════════════════════════════════════════════════════════════
`);
