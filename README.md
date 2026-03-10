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
- `GET /api/responsables`
- `POST /api/responsables/save`
- `POST /api/report/send`

Variables backend (ver `backend/.env.example`):
- `CORS_ALLOWED_ORIGINS`: `*` o lista separada por comas
- `ALERT_SMTP_*`: envio de informe general + correos personalizados por responsable
- `MS_*`: opcional para SharePoint privado por Graph

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

## 4) Despliegue recomendado (listo para usar)

### Backend en Render

Este repo incluye `render.yaml` en la raiz.

1. En Render: **New + > Blueprint**
2. Selecciona este repositorio GitHub.
3. Render detecta `render.yaml` y crea `alertas-procesos-api`.
4. En variables sensibles completa:
   - `ALERT_SMTP_USER`
   - `ALERT_SMTP_PASSWORD` (app password de Gmail)
   - `ALERT_EMAIL_FROM`
   - `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET` (si usas SharePoint privado)
5. Despliega.
6. Verifica salud: `https://TU_BACKEND.onrender.com/api/health`

### Frontend en Vercel

1. En Vercel: **Add New Project** y selecciona este repo.
2. Root Directory: `frontend`
3. Variable de entorno:
   - `NEXT_PUBLIC_API_URL=https://TU_BACKEND.onrender.com`
4. Deploy.

### CORS en produccion

En Render, define `CORS_ALLOWED_ORIGINS` con el dominio de Vercel:

`https://TU_APP.vercel.app`

Si necesitas mas de uno, separalos por coma.
