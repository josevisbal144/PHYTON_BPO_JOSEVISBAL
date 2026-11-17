#  Sistema Automatizado de Métricas y Alertas BPO en Python

##  Descripción del Proyecto
Este proyecto consiste en un script desarrollado en **Python** para automatizar el procesamiento, análisis y generación de reportes operativos en una empresa de BPO (Call Center). El sistema evalúa el desempeño de los agentes comparando métricas clave contra objetivos establecidos, clasifica el nivel de riesgo del personal, genera un reporte ejecutivo en formato Excel y envía notificaciones por correo electrónico de manera autónoma mediante un servidor SMTP.


##  Características Principales
* **Procesamiento de Datos Automatizado:** Simula la lectura de una base de datos de agentes y evalúa el cumplimiento de metas de forma masiva.
* **Evaluación de Desempeño y Riesgo:** Analiza múltiples variables críticas simultáneamente:
  * **TMO (Tiempo Medio de Operación / AHT)**
  * **CSAT (Satisfacción del Cliente)**
  * **QA (Quality Assurance / Auditoría de Calidad)**
* **Enrutamiento Lógico:** Clasifica automáticamente a los agentes en categorías (*Alto Riesgo*, *Atención Requereida*, *Óptimo*) según sus desviaciones operativas.
* **Generación de Reportes Ejecutivos:** Exporta los resultados procesados a un archivo consolidado en formato `.xlsx` (`reporte_general_agentes.xlsx`) utilizando **Pandas**.
* **Automatización de Correo Electrónico:** Envía correos electrónicos reales con diseño HTML y adjunta el reporte de Excel generado de forma segura a través de un cliente SMTP con autenticación TLS (Gmail).


##  Tecnologías y Librerías Utilizadas
* **Python 3.x** - Lenguaje de programación principal.
* **Pandas** - Manipulación, limpieza y estructuración de datos tabulares.
* **OpenPyXL** - Motor de exportación y creación de libros de Excel (`.xlsx`).
* **Smtplib / Email (Nativas de Python)** - Conexión segura con servidores de correo para el envío de alertas y reportes.

---

## ⚙️ Requisitos e Instalación

Asegúrate de tener instalado Python en tu equipo. Luego, abre la terminal en la carpeta del proyecto e instala las librerías necesarias ejecutando el siguiente comando:

```bash
pip install pandas openpyxl
