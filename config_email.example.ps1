$env:ALERT_SMTP_HOST = "smtp.office365.com"
$env:ALERT_SMTP_PORT = "587"
$env:ALERT_SMTP_USER = "correo_base@tuempresa.com"
$env:ALERT_SMTP_PASSWORD = "TU_PASSWORD_O_APP_PASSWORD"
$env:ALERT_EMAIL_FROM = "correo_base@tuempresa.com"
$env:ALERT_SMTP_USE_TLS = "1"
$env:ALERT_EMAIL_ENABLED = "1"

.venv\Scripts\python.exe main.py
