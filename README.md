# ALERTAS_PROCESOS Web

Proyecto con:
- `backend`: FastAPI para leer Excel y construir alertas.
- `frontend`: Next.js + TypeScript + Tailwind para interfaz web.

## 1) Backend (FastAPI)

Desde la raiz del proyecto:

```powershell
.venv\Scripts\python.exe -m pip install -r backend\requirements.txt
.venv\Scripts\python.exe -m uvicorn backend.app:app --reload --host 0.0.0.0 --port 8000
```

API principal:
- `POST /api/alerts/preview` con `form-data`:
  - `file`: archivo Excel
  - `sharepoint_url`: opcional, link de archivo en SharePoint/OneDrive
  - `sheet_name`: opcional (por defecto `Procesos Adminis_Penal`)
- `GET /api/health`
- `POST /api/sharepoint/diagnostic` con `form-data`:
  - `sharepoint_url`: valida configuracion de Graph y prueba descarga del archivo

### SharePoint privado con Microsoft Graph

Si el link no es publico, configura estas variables en backend:

- `MS_TENANT_ID`
- `MS_CLIENT_ID`
- `MS_CLIENT_SECRET`

Permisos recomendados en Azure App Registration (Application permissions):
- `Files.Read.All`
- `Sites.Read.All`

Luego da **admin consent** para el tenant.

Puedes validar todo con:

```bash
curl -X POST http://localhost:8000/api/sharepoint/diagnostic \
  -F "sharepoint_url=https://tuempresa.sharepoint.com/.../archivo.xlsx"
```

## 2) Frontend (Next.js)

```powershell
cd frontend
copy .env.example .env.local
npm install
npm run dev
```

El frontend queda en `http://localhost:3000` y usa `NEXT_PUBLIC_API_URL` para conectarse al backend.

## 3) Flujo de uso

1. Abre la web.
2. Sube el Excel.
   o pega link de SharePoint en el campo correspondiente.
3. Indica la hoja (si no, usa `Procesos Adminis_Penal`).
4. Presiona **Leer y previsualizar**.
5. Veras:
   - resumen (hoja usada, total de registros, total de alertas)
   - preview de datos fuente
   - preview de alertas

## 4) Despliegue recomendado

- Frontend: Vercel (Next.js nativo).
- Backend: Render/Railway/Fly.io con comando:

```bash
uvicorn backend.app:app --host 0.0.0.0 --port $PORT
```

Ajusta CORS en `backend/app.py` para dominios concretos en produccion.
