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
from docx.shared import Inches, Pt
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
        self.wfile.write(b"OK - Bot Activo")

    def log_message(self, format, *args):
        pass

def iniciar_servidor_web():
    puerto = int(os.environ.get("PORT", 10000))
    servidor = HTTPServer(("0.0.0.0", puerto), HealthHandler)
    servidor.serve_forever()

# Búsqueda en Google Sheets
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
        logger.error(f"Error Sheets: {e}")
    return None

def convertir_a_pdf(ruta_docx):
    """Convierte Word a PDF usando LibreOffice o soffice en Linux"""
    try:
        cmd = f"libreoffice --headless --convert-to pdf '{ruta_docx}'"
        res = subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=40)
        ruta_pdf = ruta_docx.replace(".docx", ".pdf")
        if os.path.exists(ruta_pdf):
            return ruta_pdf
    except Exception as e:
        logger.error(f"Fallo conversion LibreOffice: {e}")
    return None

# Estados
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
    await update.message.reply_text("👋 Usa /reporte para iniciar el formulario de servicio técnico.")

async def iniciar_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("📄 Ingrese Consecutivo (ej: 0000008) o pulse Dejar vacío:", reply_markup=reply_markup)
    return CONSECUTIVO

async def get_consecutivo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["consecutivo"] = "" if txt == "Dejar vacío" else txt
    await update.message.reply_text("🏥 Ingrese Nombre de la Clínica/Hospital o Contacto para buscar historial:", reply_markup=ReplyKeyboardRemove())
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
            f"💡 Datos encontrados en historial:\n"
            f"🏥 Clínica: {previo['hospital']}\n"
            f"👤 Contacto: {previo['contacto']}\n"
            f"📞 Teléfono: {previo['telefono']}\n"
            f"📍 Dirección: {previo['direccion']}\n\n"
            f"¿Deseas autocompletar?",
            reply_markup=reply_markup
        )
        return CONFIRMAR_AUTO
        
    await update.message.reply_text("👤 Ingrese Nombre del Contacto:")
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
        
    await update.message.reply_text("👤 Ingrese Nombre del Contacto:", reply_markup=ReplyKeyboardRemove())
    return CONTACTO

async def get_contacto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["contacto"] = update.message.text
    await update.message.reply_text("📞 Ingrese Número de Teléfono:")
    return TELEFONO

async def get_telefono(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["telefono"] = update.message.text
    await update.message.reply_text("📍 Ingrese Dirección:")
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
    await update.message.reply_text("⏱️ Ingrese Horómetro o pulse Dejar vacío:", reply_markup=reply_markup)
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
        ["Entrenamiento (Operation training)"],
        ["Seguimiento Tratamiento"],
        ["Otro (Other)"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("🛠️ Seleccione Tipo de Servicio:", reply_markup=reply_markup)
    return TIPO_SERVICIO

async def get_tipo_servicio(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["tipo_servicio"] = update.message.text
    await update.message.reply_text("📝 Ingrese Detalles / Feedback Details:", reply_markup=ReplyKeyboardRemove())
    return DETALLES

async def get_detalles(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["detalles"] = update.message.text
    teclado = [
        ["Fallo Mecánico"],
        ["Fallo Hidráulico"],
        ["Fallo en el Circuito"],
        ["Fallo en Software"],
        ["Fallo en parte de sangre"],
        ["Fallo de montaje de pieza"],
        ["Fallo de desgaste rápido de pieza"],
        ["Otros Fallos"],
        ["Ninguno / Normal"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⚠️ Seleccione Clasificación de Fallas:", reply_markup=reply_markup)
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
    
    teclado = [
        ["✅ Marcar TODO Conforme (✔)"],
        ["Apariencia, Batería, Placa, Sensores"],
        ["Pantalla táctil, Hidráulica, Parámetros"],
        ["Bombas (Cerámica/Heparina), Calibraciones"],
        ["Dejar vacío"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("📋 Lista de verificación de servicio técnico:", reply_markup=reply_markup)
    return CHECKLIST

async def get_checklist(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["checklist"] = update.message.text
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⚙️ Registro de reemplazo de componentes (Nombre, Cantidad, Obs) o pulse Dejar vacío:", reply_markup=reply_markup)
    return REPUESTOS

async def get_repuestos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["repuestos"] = "" if txt == "Dejar vacío" else txt
    teclado = [
        ["Satisfecho (Satisfied)"],
        ["Relativamente satisfecho"],
        ["Normal"],
        ["Insatisfecho (Dissatisfied)"],
        ["Muy Insatisfecho"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⭐ Opinión de usuario / Encuesta de satisfacción:", reply_markup=reply_markup)
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
    await update.message.reply_text("📅 Fecha del reporte:", reply_markup=reply_markup)
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

def marcar_en_parrafo(parrafo, texto_clave):
    """Busca el texto y coloca la marca [X] al lado sin descuadrar la celda"""
    if texto_clave.lower() in parrafo.text.lower():
        parrafo.text = parrafo.text.replace("□", "[X]").replace("☐", "[X]")
        return True
    return False

def marcar_en_celda(celda, texto_clave):
    for p in celda.paragraphs:
        if texto_clave.lower() in p.text.lower():
            p.text = p.text.replace("□", "[X]").replace("☐", "[X]")
            return True
    return False

async def get_moneda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["moneda"] = "" if txt == "Dejar vacío" else txt
    
    await update.message.reply_text("⏳ Procesando reporte y compilando PDF...")
    
    plantilla = "1-TECHNICAL SERVICE REPORT.docx"
    if not os.path.exists(plantilla):
        await update.message.reply_text("⚠️ No se encontró la plantilla .docx.")
        return ConversationHandler.END

    doc = docx.Document(plantilla)
    
    # 1. Llenar Consecutivo
    consecutivo = context.user_data.get("consecutivo", "")
    if consecutivo:
        for p in doc.paragraphs:
            if "consecutivo" in p.text.lower():
                p.add_run(f" {consecutivo}")
                break

    # Recorrer las tablas del documento
    for t in doc.tables:
        for r in t.rows:
            # Identificar celdas por su etiqueta exacta para no errar por índices
            textos_fila = [c.text.strip().lower() for c in r.cells]
            
            # Hospital Name
            if any("hospital name" in x for x in textos_fila):
                for idx, c in enumerate(r.cells):
                    if "hospital name" in c.text.lower() and idx + 1 < len(r.cells):
                        r.cells[idx + 1].text = context.user_data.get("hospital", "")
                        break
                        
            # Contacto y Telefono
            if any("contact" in x for x in textos_fila) and any("phone" in x for x in textos_fila):
                for idx, c in enumerate(r.cells):
                    if "contact" in c.text.lower() and "phone" not in c.text.lower():
                        if idx + 1 < len(r.cells):
                            r.cells[idx + 1].text = context.user_data.get("contacto", "")
                    elif "phone" in c.text.lower():
                        if idx + 1 < len(r.cells):
                            r.cells[idx + 1].text = context.user_data.get("telefono", "")
                            
            # Dirección
            if any("adress" in x or "dirección" in x for x in textos_fila):
                for idx, c in enumerate(r.cells):
                    if ("adress" in c.text.lower() or "dirección" in c.text.lower()) and idx + 1 < len(r.cells):
                        r.cells[idx + 1].text = context.user_data.get("direccion", "")
                        break

            # Modelo y Software Version
            if any("model" in x for x in textos_fila):
                for idx, c in enumerate(r.cells):
                    if "model" in c.text.lower() and "version" not in c.text.lower():
                        if idx + 1 < len(r.cells):
                            r.cells[idx + 1].text = context.user_data.get("modelo", "")
                    elif "version" in c.text.lower() or "operating version" in c.text.lower():
                        if idx + 1 < len(r.cells):
                            r.cells[idx + 1].text = context.user_data.get("version_sw", "")

            # Serial No.
            if any("serial no" in x for x in textos_fila):
                for idx, c in enumerate(r.cells):
                    if "serial no" in c.text.lower() and idx + 1 < len(r.cells):
                        r.cells[idx + 1].text = context.user_data.get("serie", "")
                        break

            # Horómetro
            if any("running hours" in x or "horometro" in x for x in textos_fila):
                for idx, c in enumerate(r.cells):
                    if ("running hours" in c.text.lower() or "horometro" in c.text.lower()) and idx + 1 < len(r.cells):
                        r.cells[idx + 1].text = context.user_data.get("horometro", "")
                        break

            # Detalles / Feedback Details
            if any("feedback details" in x or "detalles" in x for x in textos_fila):
                for idx, c in enumerate(r.cells):
                    if ("feedback details" in c.text.lower() or "detalles" in c.text.lower()) and idx + 1 < len(r.cells):
                        r.cells[idx + 1].text = context.user_data.get("detalles", "")
                        break

            # Motivo del fallo y solución
            if any("fault reason and solution" in x or "motivo del fallo" in x for x in textos_fila):
                for idx, c in enumerate(r.cells):
                    if ("fault reason" in c.text.lower() or "motivo" in c.text.lower()) and idx + 1 < len(r.cells):
                        r.cells[idx + 1].text = context.user_data.get("solucion", "")
                        break

            # Marcar Tipo de Servicio sin descuadrar
            ts = context.user_data.get("tipo_servicio", "")
            if any("installation" in x or "mantenimiento preventivo" in x for x in textos_fila):
                for c in r.cells:
                    if "instal" in ts.lower() and "installation" in c.text.lower():
                        c.text = c.text.replace("□", "[X]").replace("☐", "[X]")
                    elif "preventivo" in ts.lower() and "preventivo" in c.text.lower():
                        c.text = c.text.replace("□", "[X]").replace("☐", "[X]")
                    elif "correctivo" in ts.lower() and "correctivo" in c.text.lower():
                        c.text = c.text.replace("□", "[X]").replace("☐", "[X]")

            # Marcar Clasificación de Fallas sin descuadrar
            tf = context.user_data.get("falla_tipo", "")
            if tf and tf != "Ninguno / Normal":
                for c in r.cells:
                    if tf.lower() in c.text.lower():
                        c.text = c.text.replace("□", "[X]").replace("☐", "[X]")

            # Marcar Checklist
            chk = context.user_data.get("checklist", "")
            if "marcar todo" in chk.lower() or "todo ok" in chk.lower():
                if any("apariencia" in x or "bateria" in x or "pantalla" in x for x in textos_fila):
                    for c in r.cells:
                        c.text = c.text.replace("□", "✔").replace("☐", "✔")

            # Encuesta de satisfacción / Opinión de usuario
            sat = context.user_data.get("satisfaccion", "")
            if any("very satisfied" in x or "satisfecho" in x for x in textos_fila):
                for c in r.cells:
                    if "muy insatisfecho" in sat.lower() and "muy insatisfecho" in c.text.lower():
                        c.text = c.text.replace("□", "[X]").replace("☐", "[X]")
                    elif "insatisfecho" in sat.lower() and "insatisfecho" in c.text.lower():
                        c.text = c.text.replace("□", "[X]").replace("☐", "[X]")
                    elif "relativamente" in sat.lower() and "relativamente" in c.text.lower():
                        c.text = c.text.replace("□", "[X]").replace("☐", "[X]")
                    elif "satisfecho" in sat.lower() and "satisfecho" in c.text.lower():
                        c.text = c.text.replace("□", "[X]").replace("☐", "[X]")

            # Firmas, Ingeniero y Fechas
            if any("engineer name" in x or "customer name" in x for x in textos_fila):
                for idx, c in enumerate(r.cells):
                    if "customer name" in c.text.lower() and idx + 1 < len(r.cells):
                        r.cells[idx + 1].text = context.user_data.get("contacto", "")
                    elif "engineer name" in c.text.lower() and idx + 1 < len(r.cells):
                        r.cells[idx + 1].text = context.user_data.get("ingeniero", "Jesus Guillermo Pascual chalan")

            # Fecha de firma
            if any("signature date" in x or "fecha de firma" in x for x in textos_fila):
                fecha_val = context.user_data.get("fecha", datetime.now().strftime("%d/%m/%Y"))
                for idx, c in enumerate(r.cells):
                    if ("signature date" in c.text.lower() or "fecha de firma" in c.text.lower()) and idx + 1 < len(r.cells):
                        r.cells[idx + 1].text = fecha_val

    # Insertar la firma en la celda correspondiente
    for t in doc.tables:
        for r_idx, r in enumerate(t.rows):
            for c_idx, c in enumerate(r.cells):
                if "engineer signature" in c.text.lower() or "firma del ingeniero" in c.text.lower():
                    # La celda de abajo o contigua
                    target_cell = None
                    if r_idx + 1 < len(t.rows):
                        target_cell = t.rows[r_idx + 1].cells[c_idx]
                    else:
                        target_cell = c
                    
                    if target_cell and os.path.exists("firma_transparente.png"):
                        target_cell.text = ""
                        p = target_cell.paragraphs[0]
                        p.add_run().add_picture("firma_transparente.png", width=Inches(1.4))
                        break

    base_nombre = f"Reporte_{context.user_data.get('serie', 'servicio')}_{datetime.now().strftime('%Y%m%d_%H%M')}"
    ruta_docx = f"{base_nombre}.docx"
    doc.save(ruta_docx)
    
    # Conversión a PDF
    ruta_pdf = convertir_a_pdf(ruta_docx)
    archivo_final = ruta_pdf if (ruta_pdf and os.path.exists(ruta_pdf)) else ruta_docx
    
    with open(archivo_final, "rb") as f:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=f,
            caption="✅ ¡Reporte generado con éxito!"
        )
        
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ Cancelado.", reply_markup=ReplyKeyboardRemove())
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
