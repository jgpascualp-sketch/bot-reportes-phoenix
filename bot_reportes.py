import os
import sys
import zipfile
import logging
import threading
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO, stream=sys.stdout)
logger = logging.getLogger(__name__)

class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain")
        self.end_headers()
        self.wfile.write(b"OK - Bot Activo")
    def log_message(self, format, *args):
        pass

def iniciar_servidor_web():
    puerto = int(os.environ.get("PORT", 10000))
    HTTPServer(("0.0.0.0", puerto), HealthHandler).serve_forever()

# ==============================================================================
# BASE DE DATOS DE CLÍNICAS (Extraída de tu CRM)
# ==============================================================================
CLINICAS_BASE = {
    "CLINICA DEL RIÑON SAN CAMILLE SAC": {
        "contacto": "YURIKO MENDOZA ZAVALETA",
        "telefono": "964274611",
        "direccion": "URB BUENOS AIRES MZA. E LOTE. 2B NUEVO CHIMBOTE"
    },
    "CLÍNICA INTERDIAL NORTE SAC": {
        "contacto": "TEC SANTIAGO",
        "telefono": "930371479",
        "direccion": "C. C LOTE 38, INDEPENDENCIA 15311"
    },
    "CENTRO RENAL HABICH SAC": {
        "contacto": "MANUEL UBILLOS OLIVARES",
        "telefono": "945317263",
        "direccion": "AUXILIAR PANAMERICANA NORTE 2722"
    },
    "CLÍNICA DEL RIÑÓN SAN RENATO SAC": {
        "contacto": "JOSUE JULIO VEGA JARA",
        "telefono": "921401405",
        "direccion": "SHIMPY PALTAY S/N TARICA HUARAZ"
    },
    "HEALTH AND SAFETY OCCUPATIONAL S.A.C.": {
        "contacto": "TEC ELVIS",
        "telefono": "960138133",
        "direccion": "JR AREQUIPA 655 ANCASH"
    },
    "HEMODIAL CENTER SAC": {
        "contacto": "SR. VLADIMIR",
        "telefono": "996399989",
        "direccion": "AV. PROCERES DE LA INDEPENDENCIA 2675, SAN JUAN DE LURIGANCHO 15419"
    },
    "CENTRO DE HEMODIÁLISIS COMAS SAC": {
        "contacto": "SRTA. JANE DEYRA",
        "telefono": "940267063",
        "direccion": "PASAJE 49 MZ D LT.10 URB. EL RETABLO 2DA ETAPA, COMAS"
    },
    "C.D. NUESTRA SEÑORA DEL CARMEN EIRL.": {
        "contacto": "SR. OCHOA",
        "telefono": "902290710",
        "direccion": "JIRON O'HIGGINS N° 130, SAN VICENTE DE CAÑETE, DEPARTAMENTO DE LIMA."
    },
    "CLÍNICA DE DIÁLISIS LAJA SUCURSAL NACIMIENTO": {
        "contacto": "SR. MARCO GONZALEZ",
        "telefono": "56997225822",
        "direccion": "AV. JULIO HEMMELMANN 1751, NACIMIENTO, BIO BIO, CHILE"
    },
    "CENTRO DE DIÁLISIS RENGODIAL": {
        "contacto": "SR. BERNARDO VALENZUELA",
        "telefono": "56944281346",
        "direccion": "CALLE RENGO N° 1761, RENGO, REGION DE O'HIGGINS, CHILE."
    },
    "CENTRO DE HEMODIÁLISIS CHIMBOTE SAC": {
        "contacto": "TEC. ERICK",
        "telefono": "933559879",
        "direccion": "AV. ENRIQUE MEIGGS N° 1601 ANCASH - SANTA - CHIMBOTE"
    },
    "CENTRO DE DIÁLISIS NEFROCONTINENTE SAC": {
        "contacto": "SRTA. ELVA",
        "telefono": "94158967",
        "direccion": "AV. MEXICO N° 685 - COMAS"
    }
}

# ==============================================================================
# OPCIONES DE PREVENTIVO (Se copian en Detalles y en Solución)
# ==============================================================================
OPCIONES_PM = {
    "Primer mantenimiento preventivo de los 6 meses": "PRIMER MANTENIMIENTO PREVENTIVO DE LOS 6 MESES",
    "Segundo mantenimiento preventivo de los 12 meses": "SEGUNDO MANTENIMIENTO PREVENTIVO DE LOS 12 MESES"
}

OPCIONES_TIPO_SERVICIO = [
    ("Mantenimiento Correctivo (Repair)", "correctivo"),
    ("Mantenimiento Preventivo (PM)", "preventivo"),
    ("Diagnostico (Diagnostic / Inspection)", "diagnostico"),
    ("Entrenamiento (Operation training)", "entrenamiento"),
    ("Instalación (Installation)", "instal"),
    ("Seguimiento Tratamiento (Follow-up Trearment)", "seguimiento"),
    ("Desinstalación (Uninstallation)", "desinstal"),
    ("Otro (Other)", "otro")
]

ITEMS_CHECKLIST = [
    "Apariencia (Appearance check)", "Bateria de respaldo (Backup battery)",
    "Placa Hall o Tarjeta electronica", "Señal de sensor (Sensor signal)",
    "Pantalla táctil (Touch screen)", "Pieza hidráulica (Hydraulic parts)",
    "Parametros de Tratamiento", "Sim. de Tratamiento (Simulation treatme)",
    "Opciones (Options)", "Sensor de Cond. (Conductivity sensor)",
    "Bomba Ceramica (Ceramic pump)", "Bomba de Heparina (Syringe pump)",
    "Calibración de Conductividad (Cond. Calibration)",
    "Calibración de Temperatura (Temp. calibration)",
    "Calibración de Presión (Pressure calibration)", "Otros (Other)"
]

(
    SELECCION_CLINICA, HOSPITAL, CONTACTO, TELEFONO, DIRECCION,
    MODELO, TIPO_SERVICIO_MENU, PM_TIPO, PM_MODO_LOTE, SERIES_LOTE,
    SERIE, HOROMETRO, VERSION_SW, DETALLES, FALLA_TIPO, SOLUCION,
    CHECKLIST_MENU, REP_PARTE, REP_CANT, REP_OBS, SATISFACCION,
    CONFIRMAR_INGENIERO, SUBIR_FIRMA, FECHA,
    CUSTOMER_NAME, FIRMA_CLIENTE
) = range(26)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("👋 Usa /reporte para generar un reporte técnico rápido.")

async def iniciar_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clinicas_mem = context.bot_data.setdefault("clinicas", dict(CLINICAS_BASE))
    context.user_data.clear()
    context.user_data["checklist_seleccionados"] = []
    context.user_data["tipos_servicio_seleccionados"] = []
    context.user_data["series_lote"] = []
    context.user_data["consecutivo"] = ""

    botones = [[nombre] for nombre in clinicas_mem.keys()]
    botones.append(["➕ Nueva Clínica (Escribir manual)"])
    markup = ReplyKeyboardMarkup(botones, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("🏥 Seleccione la Clínica o pulse '➕ Nueva Clínica':", reply_markup=markup)
    return SELECCION_CLINICA

async def get_seleccion_clinica(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    clinicas_mem = context.bot_data.get("clinicas", CLINICAS_BASE)
    
    if txt in clinicas_mem:
        datos = clinicas_mem[txt]
        context.user_data["hospital"] = txt.upper()
        context.user_data["contacto"] = datos.get("contacto", "").upper()
        context.user_data["telefono"] = datos.get("telefono", "").upper()
        context.user_data["direccion"] = datos.get("direccion", "").upper()
        
        await update.message.reply_text(
            f"✅ {context.user_data['hospital']}\n"
            f"👤 Contacto: {context.user_data['contacto']}\n"
            f"📞 Teléfono: {context.user_data['telefono']}\n"
            f"📍 Dirección: {context.user_data['direccion']}",
            reply_markup=ReplyKeyboardRemove()
        )
        markup = ReplyKeyboardMarkup([["DORA-6000"], ["Otro"]], one_time_keyboard=True, resize_keyboard=True)
        await update.message.reply_text("🤖 Modelo del equipo:", reply_markup=markup)
        return MODELO

    await update.message.reply_text("🏥 Nombre de la Nueva Clínica / Hospital:", reply_markup=ReplyKeyboardRemove())
    return HOSPITAL

async def get_hospital(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["hospital"] = update.message.text.strip().upper()
    await update.message.reply_text("👤 Nombre del Contacto (Contact):")
    return CONTACTO

async def get_contacto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["contacto"] = update.message.text.strip().upper()
    await update.message.reply_text("📞 Teléfono:")
    return TELEFONO

async def get_telefono(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["telefono"] = update.message.text.strip().upper()
    await update.message.reply_text("📍 Dirección:")
    return DIRECCION

async def get_direccion(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["direccion"] = update.message.text.strip().upper()
    hosp = context.user_data["hospital"]
    context.bot_data.setdefault("clinicas", dict(CLINICAS_BASE))[hosp] = {
        "contacto": context.user_data["contacto"],
        "telefono": context.user_data["telefono"],
        "direccion": context.user_data["direccion"]
    }
    markup = ReplyKeyboardMarkup([["DORA-6000"], ["Otro"]], one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("🤖 Modelo del equipo:", reply_markup=markup)
    return MODELO

def armar_teclado_tipo_servicio(seleccionados):
    botones = []
    for nombre_largo, clave in OPCIONES_TIPO_SERVICIO:
        nombre_corto = nombre_largo.split("(")[0].strip()
        marca = "✔ " if clave in seleccionados else "⬜ "
        botones.append([InlineKeyboardButton(f"{marca}{nombre_corto}", callback_data=f"srv_{clave}")])
    botones.append([InlineKeyboardButton("✅ LISTO / CONTINUAR", callback_data="srv_DONE")])
    return InlineKeyboardMarkup(botones)

async def get_modelo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["modelo"] = update.message.text.strip().upper()
    markup = armar_teclado_tipo_servicio(context.user_data["tipos_servicio_seleccionados"])
    await update.message.reply_text("🛠️ Tipo de Servicio (marque uno o varios):", reply_markup=markup)
    return TIPO_SERVICIO_MENU

async def tipo_servicio_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    sel = context.user_data.get("tipos_servicio_seleccionados", [])
    
    if data == "srv_DONE":
        if "preventivo" in sel:
            botones_pm = [[k] for k in OPCIONES_PM.keys()]
            botones_pm.append(["Escribir manual"])
            markup = ReplyKeyboardMarkup(botones_pm, one_time_keyboard=True, resize_keyboard=True)
            await query.message.reply_text("⚙️ Seleccione el Mantenimiento Preventivo:", reply_markup=markup)
            return PM_TIPO
            
        await query.message.reply_text("🔢 Ingrese Serial No.:", reply_markup=ReplyKeyboardRemove())
        return SERIE

    clave = data.replace("srv_", "")
    if clave in sel:
        sel.remove(clave)
    else:
        sel.append(clave)
    context.user_data["tipos_servicio_seleccionados"] = sel
    await query.edit_message_reply_markup(reply_markup=armar_teclado_tipo_servicio(context.user_data["tipos_servicio_seleccionados"]))
    return TIPO_SERVICIO_MENU

async def get_pm_tipo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    if txt in OPCIONES_PM:
        texto_pm = OPCIONES_PM[txt]
        context.user_data["pm_auto"] = True
        context.user_data["
