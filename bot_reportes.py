import os
import sys
import json
import logging
import threading
import subprocess
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
from docx.shared import Inches
import gspread
from google.oauth2.service_account import Credentials

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    stream=sys.stdout
)
logger = logging.getLogger(__name__)

# Servidor HTTP para Render
class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain")
        self.end_headers()
        self.wfile.write(b"OK - Bot Reportes Activo")

    def log_message(self, format, *args):
        pass

def iniciar_servidor_web():
    puerto = int(os.environ.get("PORT", 10000))
    servidor = HTTPServer(("0.0.0.0", puerto), HealthHandler)
    logger.info(f"Servidor web Render iniciado en puerto {puerto}")
    servidor.serve_forever()

# Conexión con Google Sheets para autocompletar
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def buscar_historial(busqueda):
    try:
        creds_raw = os.environ.get("GOOGLE_CREDENTIALS_JSON")
        spreadsheet_id = os.environ.get("SPREADSHEET_ID")
        if not creds_raw or not spreadsheet_id:
            return None
        creds = Credentials.from_service_account_info(json.loads(creds_raw), scopes=SCOPES)
        gc = gspread.authorize(creds)
        sheet = gc.open_by_key(spreadsheet_id).sheet1
        registros = sheet.get_all_records()
        termino = busqueda.strip().lower()
        
        for r in reversed(registros):
            hosp = str(r.get("Hospital Name", "")).strip().lower()
            cont = str(r.get("Hospital Contact Person", "")).strip().lower()
            if (termino in hosp and hosp) or (termino in cont and cont):
                return {
                    "hospital": str(r.get("Hospital Name", "")),
                    "contacto": str(r.get("Hospital Contact Person", "")),
                    "telefono": str(r.get("Hospital Contact information", "")),
                    "direccion": str(r.get("Hospital Address", ""))
                }
    except Exception as e:
        logger.error(f"Error consultando Sheets: {e}")
    return None

def convertir_a_pdf(ruta_docx):
    """Convierte el archivo DOCX a PDF usando LibreOffice en Linux"""
    try:
        cmd = f"libreoffice --headless --convert-to pdf {ruta_docx}"
        subprocess.run(cmd, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        ruta_pdf = ruta_docx.replace(".docx", ".pdf")
        if os.path.exists(ruta_pdf):
            return ruta_pdf
    except Exception as e:
        logger.error(f"No se pudo convertir a PDF: {e}")
    return None

# Estados de conversación
(
    CONSECUTIVO,
    BUSCAR_CLIENTE,
    CONFIRMAR_AUTO,
    CONTACTO,
    TELEFONO,
    DIRECCION,
    MODELO,
    SERIE,
    HOROMETRO,
    VERSION_SW,
    TIPO_SERVICIO,
    DETALLES,
    FALLA_TIPO,
    SOLUCION,
    CHECKLIST,
    REPUESTOS,
    SATISFACCION,
    INGENIERO,
    FECHA,
    MONEDA,
) = range(20)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("👋 Usa /reporte para generar un nuevo Reporte de Servicio Técnico.")

async def iniciar_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text(
        "📄 Paso 1: Ingrese Consecutivo (ej: 0000007) o pulse el botón:",
        reply_markup=reply_markup
    )
    return CONSECUTIVO

async def get_consecutivo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["consecutivo"] = "" if txt == "Dejar vacío" else txt
    await update.message.reply_text("🏥 Ingrese Nombre de la Clínica/Hospital o Contacto para buscar datos previos:", reply_markup=ReplyKeyboardRemove())
    return BUSCAR_CLIENTE

async def get_buscar_cliente(update: Update, context: ContextTypes.DEFAULT_TYPE):
    busqueda = update.message.text
    context.user_data["hospital"] = busqueda
    previo = buscar_historial(busqueda)
    
    if previo:
        context.user_data["sug_data"] = previo
        teclado = [["✅ Sí, autocompletar"], ["✏️ No, ingresar manual"]]
        reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
        await update.message.reply_text(
            f"💡 Datos encontrados en servicios anteriores:\n"
            f"🏥 Clínica: {previo['hospital']}\n"
            f"👤 Contacto: {previo['contacto']}\n"
            f"📞 Teléfono: {previo['telefono']}\n"
            f"📍 Dirección: {previo['direccion']}\n\n"
            f"¿Deseas autocompletar estos datos?",
            reply_markup=reply_markup
        )
        return CONFIRMAR_AUTO
        
    await update.message.reply_text("👤 Ingrese Nombre del Contacto (Cliente):")
    return CONTACTO

async def get_confirmar_auto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.startswith("✅"):
        sug = context.user_data["sug_data"]
        context.user_data["hospital"] = sug["hospital"]
        context.user_data["contacto"] = sug["contacto"]
        context.user_data["telefono"] = sug["telefono"]
        context.user_data["direccion"] = sug["direccion"]
        teclado = [["DORA-6000"], ["Otro"]]
        reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
        await update.message.reply_text("🤖 Seleccione o ingrese Modelo:", reply_markup=reply_markup)
        return MODELO
        
    await update.message.reply_text("👤 Ingrese Nombre del Contacto (Cliente):", reply_markup=ReplyKeyboardRemove())
    return CONTACTO

async def get_contacto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["contacto"] = update.message.text
    await update.message.reply_text("📞 Ingrese Número de Teléfono:")
    return TELEFONO

async def get_telefono(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["telefono"] = update.message.text
    await update.message.reply_text("📍 Ingrese Dirección de la Clínica:")
    return DIRECCION

async def get_direccion(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["direccion"] = update.message.text
    teclado = [["DORA-6000"], ["Otro"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("🤖 Seleccione o ingrese Modelo:", reply_markup=reply_markup)
    return MODELO

async def get_modelo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["modelo"] = update.message.text
    await update.message.reply_text("🔢 Ingrese Serial No.:", reply_markup=ReplyKeyboardRemove())
    return SERIE

async def get_serie(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["serie"] = update.message.text
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⏱️ Ingrese Horómetro o pulse el botón:", reply_markup=reply_markup)
    return HOROMETRO

async def get_horometro(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["horometro"] = "" if txt == "Dejar vacío" else txt
    teclado = [["020305"], ["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("💻 Ingrese Versión de Software:", reply_markup=reply_markup)
    return VERSION_SW

async def get_version_sw(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["version_sw"] = "" if txt == "Dejar vacío" else txt
    teclado = [
        ["Instalación (Installation)"],
        ["Mantenimiento Preventivo (PM)"],
        ["Mantenimiento Correctivo (Repair)"],
        ["Inspección / Diagnóstico"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("🛠️ Seleccione Tipo de Servicio (se marcará con X):", reply_markup=reply_markup)
    return TIPO_SERVICIO

async def get_tipo_servicio(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["tipo_servicio"] = update.message.text
    await update.message.reply_text("📝 Ingrese Detalles del Trabajo Realizado:", reply_markup=ReplyKeyboardRemove())
    return DETALLES

async def get_detalles(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["detalles"] = update.message.text
    teclado = [
        ["Ninguno / Normal"],
        ["Fallo Hidráulico", "Fallo Circuito"],
        ["Fallo Mecánico", "Fallo Software"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⚠️ Clasificación de Fallas (se marcará con X):", reply_markup=reply_markup)
    return FALLA_TIPO

async def get_falla_tipo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["falla_tipo"] = update.message.text
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("💡 Ingrese Motivo del Fallo y Solución:", reply_markup=reply_markup)
    return SOLUCION

async def get_solucion(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["solucion"] = "" if txt == "Dejar vacío" else txt
    teclado = [["Todos conformes / Todo OK"], ["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("📋 Lista de verificación / Checklist de pruebas:", reply_markup=reply_markup)
    return CHECKLIST

async def get_checklist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["checklist"] = update.message.text
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⚙️ Registro de Reemplazo de Componentes (o pulse Dejar vacío):", reply_markup=reply_markup)
    return REPUESTOS

async def get_repuestos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["repuestos"] = "" if txt == "Dejar vacío" else txt
    teclado = [
        ["Satisfecho (Very Satisfied)"],
        ["Relativamente satisfecho"],
        ["Normal"],
        ["Insatisfecho"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⭐ Encuesta de Satisfacción del Usuario (se marcará con X):", reply_markup=reply_markup)
    return SATISFACCION

async def get_satisfaccion(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["satisfaccion"] = update.message.text
    teclado = [["Jesus Guillermo Pascual chalan"], ["Ingresar otro nombre"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("👷 Ingeniero a cargo:", reply_markup=reply_markup)
    return INGENIERO

async def get_ingeniero(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    if txt == "Ingresar otro nombre":
        await update.message.reply_text("Escriba el nombre completo del ingeniero:", reply_markup=ReplyKeyboardRemove())
        return INGENIERO
    context.user_data["ingeniero"] = txt
    
    fecha_hoy = datetime.now().strftime("%d/%m/%Y")
    teclado = [[fecha_hoy], ["Ingresar otra fecha"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text(f"📅 Fecha del reporte (presione para confirmar):", reply_markup=reply_markup)
    return FECHA

async def get_fecha(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    if txt == "Ingresar otra fecha":
        await update.message.reply_text("Ingrese la fecha (ej: DD/MM/AAAA):", reply_markup=ReplyKeyboardRemove())
        return FECHA
    context.user_data["fecha"] = txt
    
    teclado = [["Dejar vacío"], ["Soles (S/.)"], ["USD ($)"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("💵 Moneda y cobro:", reply_markup=reply_markup)
    return MONEDA

async def get_moneda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["moneda"] = "" if txt == "Dejar vacío" else txt
    
    await update.message.reply_text("⏳ Generando el reporte técnico y compilando documento...")
    
    plantilla = "1-TECHNICAL SERVICE REPORT.docx"
    if not os.path.exists(plantilla):
        await update.message.reply_text("⚠️ No se encontró la plantilla .docx en el repositorio.")
        return ConversationHandler.END

    doc = docx.Document(plantilla)
    t = doc.tables[0]
    
    # 1. Cabecera y datos de la clínica
    t.rows[0].cells[1].text = context.user_data.get("hospital", "")
    t.rows[1].cells[1].text = context.user_data.get("contacto", "")
    t.rows[1].cells[3].text = context.user_data.get("telefono", "")
    t.rows[2].cells[1].text = context.user_data.get("direccion", "")
    t.rows[3].cells[1].text = context.user_data.get("modelo", "")
    t.rows[3].cells[3].text = context.user_data.get("version_sw", "")
    t.rows[4].cells[1].text = context.user_data.get("serie", "")
    t.rows[5].cells[1].text = context.user_data.get("horometro", "")
    
    # 2. Tipo de servicio (Marca con [X])
    ts = context.user_data.get("tipo_servicio", "")
    if "Instalación" in ts:
        t.rows[6].cells[1].text = "[X] Instalación"
    elif "Preventivo" in ts:
        t.rows[6].cells[1].text = "[X] Preventivo (PM)"
    elif "Correctivo" in ts:
        t.rows[6].cells[1].text = "[X] Correctivo (Repair)"
    else:
        t.rows[6].cells[1].text = f"[X] {ts}"

    # 3. Detalles y soluciones
    t.rows[7].cells[1].text = context.user_data.get("detalles", "")
    
    # Clasificación de fallas
    tf = context.user_data.get("falla_tipo", "")
    t.rows[9].cells[1].text = f"[X] {tf}" if tf and tf != "Ninguno / Normal" else "Ninguno / Normal"
    t.rows[10].cells[1].text = context.user_data.get("solucion", "")
    
    # 4. Tabla de firmas y validaciones
    t2 = doc.tables[1]
    contacto = context.user_data.get("contacto", "")
    ingeniero = context.user_data.get("ingeniero", "Jesus Guillermo Pascual chalan")
    fecha_reporte = context.user_data.get("fecha", datetime.now().strftime("%d/%m/%Y"))
    
    t2.rows[0].cells[1].text = contacto
    t2.rows[0].cells[3].text = ingeniero
    
    # Firma digital en imagen
    if os.path.exists("firma_transparente.png"):
        cell_sig = t2.rows[1].cells[3]
        cell_sig.text = ""
        p = cell_sig.paragraphs[0]
        p.add_run().add_picture("firma_transparente.png", width=Inches(1.5))
        
    t2.rows[2].cells[1].text = fecha_reporte
    t2.rows[2].cells[3].text = fecha_reporte
    
    # Guardar Word
    base_nombre = f"Reporte_{context.user_data.get('serie', 'servicio')}_{datetime.now().strftime('%Y%m%d_%H%M')}"
    ruta_docx = f"{base_nombre}.docx"
    doc.save(ruta_docx)
    
    # Intentar generar PDF
    ruta_pdf = convertir_a_pdf(ruta_docx)
    archivo_a_enviar = ruta_pdf if (ruta_pdf and os.path.exists(ruta_pdf)) else ruta_docx
    
    with open(archivo_a_enviar, "rb") as f:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=f,
            caption="✅ ¡Reporte generado con éxito!"
        )
        
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ Reporte cancelado.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

def main():
    t = threading.Thread(target=iniciar_servidor_web, daemon=True)
    t.start()

    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        logger.error("TELEGRAM_BOT_TOKEN no configurado")
        sys.exit(1)

    app = ApplicationBuilder().token(token).build()
    
    conv = ConversationHandler(
        entry_points=[CommandHandler("reporte", iniciar_reporte)],
        states={
            CONSECUTIVO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_consecutivo)],
            BUSCAR_CLIENTE: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_buscar_cliente)],
            CONFIRMAR_AUTO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_confirmar_auto)],
            CONTACTO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_contacto)],
            TELEFONO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_telefono)],
            DIRECCION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_direccion)],
            MODELO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_modelo)],
            SERIE: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_serie)],
            HOROMETRO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_horometro)],
            VERSION_SW: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_version_sw)],
            TIPO_SERVICIO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_tipo_servicio)],
            DETALLES: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_detalles)],
            FALLA_TIPO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_falla_tipo)],
            SOLUCION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_solucion)],
            CHECKLIST: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_checklist)],
            REPUESTOS: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_repuestos)],
            SATISFACCION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_satisfaccion)],
            INGENIERO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_ingeniero)],
            FECHA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_fecha)],
            MONEDA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_moneda)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv)
    
    logger.info("Bot de Reportes en ejecución...")
    app.run_polling()

if __name__ == "__main__":
    main()
