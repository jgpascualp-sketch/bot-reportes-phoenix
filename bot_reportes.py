import os
import json
import logging
import threading
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)
import docx
from docx.shared import Inches
import gspread
from google.oauth2.service_account import Credentials

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# Servidor web ligero para evitar que Render apague el proceso
class SimpleHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Bot de Reportes Activo")

    def log_message(self, format, *args):
        pass

def run_http():
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), SimpleHandler)
    server.serve_forever()

# Búsqueda en Google Sheets si existen credenciales
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def buscar_datos_hospital(nombre_hosp):
    try:
        creds_raw = os.environ.get("GOOGLE_CREDENTIALS_JSON")
        spreadsheet_id = os.environ.get("SPREADSHEET_ID")
        if not creds_raw or not spreadsheet_id:
            return None
        creds = Credentials.from_service_account_info(json.loads(creds_raw), scopes=SCOPES)
        gc = gspread.authorize(creds)
        sheet = gc.open_by_key(spreadsheet_id).sheet1
        registros = sheet.get_all_records()
        termino = nombre_hosp.strip().lower()
        for r in reversed(registros):
            h = str(r.get("Hospital Name", "")).strip().lower()
            if termino in h and h:
                return {
                    "Hospital Name": str(r.get("Hospital Name", "")),
                    "Contact": str(r.get("Hospital Contact Person", "")),
                    "Phone": str(r.get("Hospital Contact information", "")),
                    "Address": str(r.get("Hospital Address", ""))
                }
    except Exception as e:
        logger.error(f"Error Sheets: {e}")
    return None

(
    CONSECUTIVO,
    HOSPITAL,
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
    SATISFACCION,
    MONEDA,
) = range(16)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("👋 ¡Hola Jesús! Usa /reporte para generar tu reporte técnico.")

async def iniciar_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text(
        "📄 Paso 1/14: Ingrese Consecutivo (ej: 0000007) o pulse el botón:",
        reply_markup=reply_markup
    )
    return CONSECUTIVO

async def get_consecutivo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["consecutivo"] = "" if txt == "Dejar vacío" else txt
    await update.message.reply_text("🏥 Paso 2/14: Ingrese Nombre del Cliente / Hospital:", reply_markup=ReplyKeyboardRemove())
    return HOSPITAL

async def get_hospital(update: Update, context: ContextTypes.DEFAULT_TYPE):
    hosp = update.message.text
    context.user_data["hospital"] = hosp
    sug = buscar_datos_hospital(hosp)
    if sug:
        context.user_data["sug_data"] = sug
        teclado = [["✅ Sí, autocompletar"], ["✏️ No, ingresar manual"]]
        reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
        await update.message.reply_text(
            f"💡 Coincidencia encontrada:\n"
            f"👤 Contacto: {sug['Contact']}\n"
            f"📞 Teléfono: {sug['Phone']}\n"
            f"📍 Dirección: {sug['Address']}\n\n"
            f"¿Deseas autocompletar estos datos?",
            reply_markup=reply_markup
        )
        return CONFIRMAR_AUTO
    await update.message.reply_text("👤 Paso 3/14: Ingrese Nombre de Contacto:")
    return CONTACTO

async def get_confirmar_auto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.startswith("✅"):
        sug = context.user_data["sug_data"]
        context.user_data["contacto"] = sug["Contact"]
        context.user_data["telefono"] = sug["Phone"]
        context.user_data["direccion"] = sug["Address"]
        teclado = [["DORA-6000"], ["Otro"]]
        reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
        await update.message.reply_text("🤖 Paso 6/14: Seleccione o ingrese Modelo:", reply_markup=reply_markup)
        return MODELO
    await update.message.reply_text("👤 Paso 3/14: Ingrese Nombre de Contacto:", reply_markup=ReplyKeyboardRemove())
    return CONTACTO

async def get_contacto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["contacto"] = update.message.text
    await update.message.reply_text("📞 Paso 4/14: Ingrese Teléfono de contacto:")
    return TELEFONO

async def get_telefono(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["telefono"] = update.message.text
    await update.message.reply_text("📍 Paso 5/14: Ingrese Dirección:")
    return DIRECCION

async def get_direccion(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["direccion"] = update.message.text
    teclado = [["DORA-6000"], ["Otro"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("🤖 Paso 6/14: Seleccione o ingrese Modelo:", reply_markup=reply_markup)
    return MODELO

async def get_modelo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["modelo"] = update.message.text
    await update.message.reply_text("🔢 Paso 7/14: Ingrese Número de Serial (Serial No.):", reply_markup=ReplyKeyboardRemove())
    return SERIE

async def get_serie(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["serie"] = update.message.text
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⏱️ Paso 8/14: Ingrese Horómetro o pulse el botón:", reply_markup=reply_markup)
    return HOROMETRO

async def get_horometro(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["horometro"] = "" if txt == "Dejar vacío" else txt
    teclado = [["020305"], ["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("💻 Paso 9/14: Ingrese Versión de Software:", reply_markup=reply_markup)
    return VERSION_SW

async def get_version_sw(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["version_sw"] = "" if txt == "Dejar vacío" else txt
    teclado = [
        ["Instalación (Installation)"],
        ["Mantenimiento Preventivo (PM)"],
        ["Mantenimiento Correctivo (Repair)"],
        ["Diagnostico (Diagnostic / Inspection)"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("🛠️ Paso 10/14: Seleccione Tipo de Servicio:", reply_markup=reply_markup)
    return TIPO_SERVICIO

async def get_tipo_servicio(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["tipo_servicio"] = update.message.text
    await update.message.reply_text("📝 Paso 11/14: Ingrese Detalles / Feedback Details:", reply_markup=ReplyKeyboardRemove())
    return DETALLES

async def get_detalles(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["detalles"] = update.message.text
    teclado = [
        ["Ninguno / Normal"],
        ["Fallo Hidráulico", "Fallo Circuito"],
        ["Fallo Mecánico", "Fallo Software"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⚠️ Paso 12/14: Clasificación de Fallas:", reply_markup=reply_markup)
    return FALLA_TIPO

async def get_falla_tipo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["falla_tipo"] = update.message.text
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("💡 Paso 13/14: Ingrese Motivo del Fallo y Solución:", reply_markup=reply_markup)
    return SOLUCION

async def get_solucion(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["solucion"] = "" if txt == "Dejar vacío" else txt
    teclado = [
        ["Satisfecho", "Relativamente satisfecho"],
        ["Normal", "Insatisfecho"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⭐ Paso 14/14: Encuesta de satisfacción:", reply_markup=reply_markup)
    return SATISFACCION

async def get_satisfaccion(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["satisfaccion"] = update.message.text
    teclado = [["Dejar vacío"], ["Soles (S/.)"], ["USD ($)"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("💵 Moneda y cobro:", reply_markup=reply_markup)
    return MONEDA

async def get_moneda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["moneda"] = "" if txt == "Dejar vacío" else txt
    
    await update.message.reply_text("⏳ Generando reporte Word con firma...")
    
    doc = docx.Document("1-TECHNICAL SERVICE REPORT.docx")
    t = doc.tables[0]
    
    t.rows[0].cells[1].text = context.user_data.get("hospital", "")
    t.rows[1].cells[1].text = context.user_data.get("contacto", "")
    t.rows[1].cells[3].text = context.user_data.get("telefono", "")
    t.rows[2].cells[1].text = context.user_data.get("direccion", "")
    t.rows[3].cells[1].text = context.user_data.get("modelo", "")
    t.rows[3].cells[3].text = context.user_data.get("version_sw", "")
    t.rows[4].cells[1].text = context.user_data.get("serie", "")
    t.rows[5].cells[1].text = context.user_data.get("horometro", "")
    t.rows[7].cells[1].text = context.user_data.get("detalles", "")
    t.rows[10].cells[1].text = context.user_data.get("solucion", "")
    
    t2 = doc.tables[1]
    fecha_actual = datetime.now().strftime("%d/%m/%Y")
    t2.rows[0].cells[1].text = context.user_data.get("contacto", "")
    t2.rows[0].cells[3].text = "Jesus Guillermo Pascual chalan"
    
    if os.path.exists("firma_transparente.png"):
        cell_sig = t2.rows[1].cells[3]
        cell_sig.text = ""
        p = cell_sig.paragraphs[0]
        p.add_run().add_picture("firma_transparente.png", width=Inches(1.5))
        
    t2.rows[2].cells[1].text = fecha_actual
    t2.rows[2].cells[3].text = fecha_actual
    
    nombre_archivo = f"Reporte_{context.user_data.get('serie', 'servicio')}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    doc.save(nombre_archivo)
    
    with open(nombre_archivo, "rb") as f:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=f,
            caption="✅ ¡Reporte generado exitosamente con tu firma!"
        )
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ Cancelado.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

def main():
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise ValueError("TELEGRAM_BOT_TOKEN no configurado")

    threading.Thread(target=run_http, daemon=True).start()

    app = ApplicationBuilder().token(token).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("reporte", iniciar_reporte)],
        states={
            CONSECUTIVO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_consecutivo)],
            HOSPITAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_hospital)],
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
            SATISFACCION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_satisfaccion)],
            MONEDA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_moneda)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv)
    logger.info("Bot de Reportes en ejecucion y escuchando...")
    app.run_polling()

if __name__ == "__main__":
    main()
