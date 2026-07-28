Donhang internal app — frontend + GAS proxy

Repository contents:
- code.gs (Google Apps Script backend)
- webapp.html (original GAS-served frontend)
- public/ (static frontend for Vercel)
- api/gas-proxy.js (serverless proxy for GAS)

Deployment:
- Frontend: deploy `public/` to Vercel or any static host.
- API proxy: deploy `api/gas-proxy.js` to Vercel serverless functions and set `WEB_APP_URL` in `code.gs` to your GAS exec URL.

See project docs in the repo for more details.
