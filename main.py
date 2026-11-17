import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd

# ==============================================================================
# 1. ENTRADA DE DATOS (Simulación de la data cruda proveniente de la operación)
# ==============================================================================
# Definimos una lista de diccionarios que contiene la información de los agentes
# evaluados, incluyendo sus métricas reales y las metas establecidas.
agentes_data = [
    {
        "agent_id": "AG001",
        "agent_name": "Laura",
        "calls_handled": 420,
        "tmo_minutes": 8.2,
        "csat_percent": 94,
        "qa_percent": 96,
        "targets": {"tmo_minutes": 8.0, "csat_percent": 90, "qa_percent": 90},
        "status": "near_tmo_target",
    },
    {
        "agent_id": "AG002",
        "agent_name": "Carlos",
        "calls_handled": 380,
        "tmo_minutes": 9.5,
        "csat_percent": 85,
        "qa_percent": 82,
        "targets": {"tmo_minutes": 8.0, "csat_percent": 90, "qa_percent": 90},
        "status": "needs_improvement",
    },
    {
        "agent_id": "AG003",
        "agent_name": "Ana",
        "calls_handled": 450,
        "tmo_minutes": 7.5,
        "csat_percent": 92,
        "qa_percent": 95,
        "targets": {"tmo_minutes": 8.0, "csat_percent": 90, "qa_percent": 90},
        "status": "optimal",
    },
    {
        "agent_id": "AG004",
        "agent_name": "Mateo",
        "calls_handled": 310,
        "tmo_minutes": 9.1,
        "csat_percent": 88,
        "qa_percent": 86,
        "targets": {"tmo_minutes": 8.0, "csat_percent": 90, "qa_percent": 90},
        "status": "quality_review",
    },
    {
        "agent_id": "AG005",
        "agent_name": "Sofia",
        "calls_handled": 440,
        "tmo_minutes": 7.8,
        "csat_percent": 95,
        "qa_percent": 98,
        "targets": {"tmo_minutes": 8.0, "csat_percent": 90, "qa_percent": 90},
        "status": "optimal",
    },
]

print("--- 1. INICIANDO PROCESAMIENTO DE DATOS ---")

# ==============================================================================
# 2. PROCESAMIENTO INDIVIDUAL (Lógica de Metas, Incumplimientos y Riesgo)
# ==============================================================================
# Recorremos cada agente de forma independiente (equivalente al Split Out en n8n)
# para evaluar su desempeño en comparación con las metas operativas.
processed_agents = []

for agent in agentes_data:
  # Evaluamos si las métricas superan o caen por debajo de los umbrales permitidos
  tmo_exceeded = agent["tmo_minutes"] > agent["targets"]["tmo_minutes"]
  csat_low = agent["csat_percent"] < agent["targets"]["csat_percent"]
  qa_low = agent["qa_percent"] < agent["targets"]["qa_percent"]

  # Acumulamos las fallas específicas en una lista de problemas
  issues = []
  if tmo_exceeded:
    issues.append(f"TMO elevado ({agent['tmo_minutes']} min)")
  if csat_low:
    issues.append(f"CSAT bajo ({agent['csat_percent']}%)")
  if qa_low:
    issues.append(f"QA bajo ({agent['qa_percent']}%)")

  # Asignamos el nivel de riesgo según la cantidad de incumplimientos y el estado base
  if (
      agent["status"] in ["needs_improvement", "quality_review"]
      or len(issues) >= 2
  ):
    risk_level = "Alto Riesgo"
  elif agent["status"] == "near_tmo_target" or len(issues) == 1:
    risk_level = "Atención Requerida"
  else:
    risk_level = "Óptimo"

  # Estructuramos el diccionario limpio con el resultado de la evaluación
  processed_agents.append({
      "agent_id": agent["agent_id"],
      "agent_name": agent["agent_name"],
      "calls_handled": agent["calls_handled"],
      "tmo_minutes": agent["tmo_minutes"],
      "csat_percent": agent["csat_percent"],
      "qa_percent": agent["qa_percent"],
      "status": agent["status"],
      "risk_level": risk_level,
      "issues_count": len(issues),
      "issues": ", ".join(issues) if issues else "Sin novedades",
  })

print(f"Total agentes procesados: {len(processed_agents)}")

# ==============================================================================
# 3. ENRUTAMIENTO CONDICIONAL (Equivalente al nodo Switch)
# ==============================================================================
# Filtramos la información según el nivel de riesgo para separar los casos críticos
print("\n--- 3. APLICANDO ENRUTAMIENTO (SWITCH) ---")
alertas_tmo = [a for a in processed_agents if a["risk_level"] == "Alto Riesgo"]
seguimiento_tmo = [
    a for a in processed_agents if a["risk_level"] == "Atención Requerida"
]

print(
    f"-> Ruta Alerta TMO (Alto Riesgo): {len(alertas_tmo)} agentes detectados."
)
for a in alertas_tmo:
  print(f"   [ALERTA] {a['agent_name']} - Motivos: {a['issues']}")

# ==============================================================================
# 4. GENERACIÓN DE ARCHIVO EXCEL CONSOLIDADO (Equivalente a Convert to XLSX)
# ==============================================================================
# Transformamos la lista procesada en un DataFrame de Pandas para exportarla a Excel
print("\n--- 4. GENERANDO REPORTE EXCEL CONSOLIDADO ---")
df_report = pd.DataFrame(processed_agents)

# Renombramos las columnas para que tengan una presentación formal y ejecutiva
df_excel = df_report.rename(columns={
    "agent_id": "ID Agente",
    "agent_name": "Nombre",
    "calls_handled": "Llamadas",
    "tmo_minutes": "TMO (Min)",
    "csat_percent": "CSAT (%)",
    "qa_percent": "QA (%)",
    "status": "Estado Base",
    "risk_level": "Nivel de Riesgo",
    "issues_count": "Cant. Faltas",
    "issues": "Detalle de Incumplimientos",
})

# Guardamos el archivo en formato .xlsx en el directorio local
nombre_archivo = "reporte_general_agentes.xlsx"
df_excel.to_excel(nombre_archivo, index=False)
print(f"¡Archivo '{nombre_archivo}' generado y guardado con éxito!")

# ==============================================================================
# 5. CONFIGURACIÓN Y ENVÍO REAL DE CORREOS MEDIANTE SMTP (Equivalente a Send Email)
# ==============================================================================
print("\n--- 5. ENVIANDO CORREO REAL CON EL REPORTE ---")

# Credenciales y parámetros del servidor de correo (Gmail)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
REM_EMAIL = "josevisbal144@gmail.com"
REM_PASSWORD = "rboxqrotxnuljubj"  # Contraseña de aplicación generada en Google
DESTINATARIO = "josevisbal144@gmail.com"  # Dirección de destino del reporte


def enviar_correo_real(destinatario, asunto, cuerpo_html, archivo_adjunto=None):
  try:
    # Creamos el objeto multiparte para soportar texto HTML y archivos adjuntos
    msg = MIMEMultipart()
    msg["From"] = REM_EMAIL
    msg["To"] = destinatario
    msg["Subject"] = asunto

    # Adjuntamos el cuerpo del mensaje interpretándolo como HTML
    msg.attach(MIMEText(cuerpo_html, "html"))

    # Verificamos si existe un archivo adjunto para incluirlo en el mensaje
    if archivo_adjunto and os.path.exists(archivo_adjunto):
      with open(archivo_adjunto, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

      encoders.encode_base64(part)
      part.add_header(
          "Content-Disposition",
          f"attachment; filename= {os.path.basename(archivo_adjunto)}",
      )
      msg.attach(part)

    # Conexión al servidor SMTP utilizando seguridad TLS
    print(f"Conectando al servidor SMTP para enviar correo a {destinatario}...")
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()  # Iniciamos conexión cifrada
    server.login(REM_EMAIL, REM_PASSWORD)  # Autenticación con cuenta y token
    server.sendmail(
        REM_EMAIL, destinatario, msg.as_string()
    )  # Envío del mensaje
    server.quit()  # Cierre de la conexión

    print(f"¡Correo enviado con éxito a {destinatario}!")

  except Exception as e:
    print(f"❌ Error al enviar el correo: {e}")


# Estructura del diseño HTML para el cuerpo del correo gerencial
cuerpo_html_gerencia = f"""
<div style="font-family: Arial, sans-serif; color: #333333; max-width: 600px; line-height: 1.6;">
  <h2 style="color: #2c3e50; border-bottom: 2px solid #27ae60; padding-bottom: 8px;">
    📊 Reporte General y Seguimiento Operativo BPO
  </h2>
  <p>Hola, José,</p>
  <p>Se adjunta el reporte consolidado generado automáticamente desde Python con las métricas de desempeño de los agentes.</p>
  <div style="background-color: #f8f9fa; border-left: 4px solid #27ae60; padding: 12px; margin: 15px 0;">
    <p style="margin: 0; font-weight: bold;">📁 Archivo adjunto:</p>
    <p style="margin: 5px 0 0 0; color: #555555;"><code>{nombre_archivo}</code></p>
  </div>
  <p style="font-size: 12px; color: #888888;">Sistema Automático de Gestión BPO - Python Script.</p>
</div>
"""

# Ejecución final de la función para enviar el correo real con su archivo adjunto
enviar_correo_real(
    destinatario=DESTINATARIO,
    asunto="📊 Reporte Periódico de Métricas y Seguimiento BPO",
    cuerpo_html=cuerpo_html_gerencia,
    archivo_adjunto=nombre_archivo,
)

print("\n¡Flujo en Python ejecutado de extremo a extremo con éxito!")