import pandas as pd
import pyautogui as pg
import webbrowser as web
import subprocess
import time
import random
import pyperclip
import keyboard
import threading
import os
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, Side  

#Rutas
ruta_chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

pdf__path = {
    "FCNM": r"FCNM_BROCHURE.pdf",
    "FIPA": r"FIPA_BROCHURE.pdf",
    "FIARN": r"FIARN_BROCHURE.pdf",
    "FIQ": r"FIQ_BROCHURE.pdf",
    "FIIS": r"FIIS_BROCHURE.pdf",
    "FIME": r"FIME_BROCHURE.pdf",
    "FIEE": r"FIEE_BROCHURE.pdf",
    "FCC": r"FCC_BROCHURE.pdf",
    "FCED": r"FCED_BROCURE.pdf",
    "FCE": {
        "Doctorado": r"FCE_DOCTORADO_BROCHURE.pdf",
        "Maestría": r"FCE_MAESTRIA_BROCHURE.pdf"
        },
    "FCA": r"FCA_BROCHURE.pdf",
    "FCS": {
        "Doctorado": r"FCS_DOCTORADO_BROCHURE.pdf",
        "Maestría": r"FCS_MAESTRIA_BROCHURE.pdf"
    }
}

# Función para esperar con variación aleatoria
def esperar_dinamico(tiempo_min, tiempo_max):
    tiempo_espera = random.uniform(tiempo_min, tiempo_max)
    time.sleep(tiempo_espera)
    
#funcion para enviar mensaje por whatsapp predeterminado
def enviar_mensaje(celular, nombre, apellido, programa, name_programa, facultad):
    saludos = [
        f"""👋 Hola {nombre} {apellido}!
Soy Edgar Ramos de la *Escuela de Posgrado de la UNAC*
🚀 *Tengo una noticia clave para tu desarrollo profesional!*""",
        f"""👋 Buenas tardes {nombre} {apellido}!
Soy Edgar Ramos de la *Escuela de Posgrado de la UNAC*
📢 Quiero que seas parte de esta gran oportunidad. *¡Atento a lo siguiente!*""",
        f"""👋 Hola {nombre} {apellido}!
Te habla Edgar Ramos desde la *Escuela de Posgrado de la UNAC*
💡 No dejes pasar esta información clave para tu futuro. *¡Mira esto!*"""
    ]

    mensaje_base = (
        f"""
📢 *ÚLTIMO DÌA PARA INSCRIBIRSE:*
{name_programa}
📅 Cierre de inscripciones: 25 de marzo

💰 *Costo de inscripción:* S/ 145
📂 *Incluye:* Carpeta de Postulante y Derecho de Inscripción

📅 *Fechas clave:*
    - Entrevista virtual: 26 y 27 de marzo
    - Resultados: 1-2 días después del examen
    - Inicio de clases: Primera semana de abril

"""
    )

    if facultad in ["FCS", "FIARN"]:
        mensaje_base += (
            "📍 *Modalidad de estudios:*\n"
            "    - 100% Virtual\n"
            "    - Fines de semana, de 8:00 a.m. a 8:00 p.m.\n"
            "\n"
        )
    else:
        mensaje_base += (
            "📍 *Modalidad de estudios:*\n"
            "    - 20% presencial y 80% virtual\n"
            "    - Presencial: 1 vez al mes (fines de semana, de 8:00 a.m. a 8:00 p.m.)\n"
            "\n"
        )

    if programa == "Doctorado":
        mensaje_base += (
            "⏳ *Duración del programa:* 6 semestres académicos\n"
            "💵 *Costo por semestre:* S/ ~~2500~~ S/ 2100\n"
            "\n"
        )
    elif programa == "Maestría":
        mensaje_base += (
            "⏳ *Duración del programa:* 3 semestres académicos\n"
            "💵 *Costo por semestre:* S/ ~~2500~~ S/ 2100\n"
            "\n"
        )

    mensaje_base += "🎯 ¡Inscríbete y prepárate para el siguiente nivel en tu formación profesional! 💼📖"

    pdf_mensaje = random.choice([
        "📎 Te adjunto el brochure con toda la información necesaria.",
        "📂 Aquí tienes el documento con la información detallada.",
        "🔍 Te envío el brochure oficial con los detalles del programa.",
        "📄 En el brochure adjunto encontrarás todos los detalles."
    ])
    consulta_mensaje = random.choice([
        """
🤝 Si tienes alguna duda o necesitas ayuda con tu inscripción, estoy aquí para apoyarte.
📩 *Correo:* posgrado.admision@unac.edu.pe
📞 *WhatsApp:* 900969591\n
✨ ¡Responde este mensaje y asegura tu inscripción hoy mismo!""",
        """
📢 *¡No pierdas esta oportunidad!*
Si tienes consultas, escríbeme y te ayudaré en lo que necesites.
📩 *Correo:* posgrado.admision@unac.edu.pe
📞 *WhatsApp:* 900969591\n
✅ Responde este mensaje y da el primer paso hacia tu futuro académico.""",
        """
📌 Estoy disponible para resolver cualquier duda y acompañarte en tu proceso de inscripción.
📩 *Correo:* posgrado.admision@unac.edu.pe
📞 *WhatsApp:* 900969591\n
🚀 ¡Escríbeme ahora y asegura tu cupo en la maestría!"""
    ])
    
def normalize_phone(phone):
    """Normalize phone numbers to ensure they are 9 digits without spaces."""
    phone = ''.join(filter(str.isdigit, str(phone)))  # Remove non-digit characters
    return phone[-9:] if len(phone) >= 9 else None  # Ensure it's 9 digits

def process_excel_and_send_messages(file_path):
    """Read Excel file, normalize phone numbers, and send WhatsApp messages."""
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Normalize phone numbers
    df['Celular'] = df['Celular'].apply(normalize_phone)

    # Filter out rows with invalid phone numbers
    df = df[df['Celular'].notnull()]

    # Iterate through the DataFrame and send messages
    for _, row in df.iterrows():
        enviar_mensaje(
            celular=row['Celular'],
            nombre=row['Nombre'],
            apellido=row['Apellido_Paterno'],
            programa=row['Programa'],
            name_programa=row['Name_Programa'],
            facultad=row['Programa'][:4]  # Assuming the first 4 letters indicate the faculty
        )

# Main execution
if __name__ == "__main__":
    excel_path = r"Preinscripción.xlsx"  # Path to the Excel file
    process_excel_and_send_messages(excel_path)


