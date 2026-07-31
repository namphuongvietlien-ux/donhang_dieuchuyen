Donhang internal app — frontend + GAS proxy

Repository contents:
- code.gs: Google Apps Script backend
- webapp.html: original GAS-served frontend
- public/: static frontend for Vercel
- api/gas-proxy.js: serverless proxy for GAS

Deployment:
- Frontend: deploy public/ to Vercel or any static host.
- API proxy: deploy api/gas-proxy.js to Vercel serverless functions.
- code.gs:
	WEB_APP_URL phải là URL web app giao diện người dùng.
	GAS_URL trong api/gas-proxy.js mới là URL Apps Script /exec.

Usage guide:
- Xem hướng dẫn tại HUONG_DAN_SU_DUNG.md
