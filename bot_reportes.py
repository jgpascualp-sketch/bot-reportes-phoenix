import os
import sys
import json
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
from docx.shared import Inches
import gspread
from google.oauth2.service_account import Credentials

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    stream=sys.stdout
)
logger = logging.getLogger(__name__)

# Servidor web Render
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

# Google Sheets Autocompletado
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

# Lista fija de verificación
ITEMS_CHECKLIST = [
    "Apariencia (Appearance check)",
    "Bateria de respaldo (Backup battery)",
    "Placa Hall o Tarjeta electronica",
    "Señal de sensor (Sensor signal)",
    "Pantalla táctil (Touch screen)",
    "Pieza hidráulica (Hydraulic parts)",
    "Parametros de Tratamiento",
    "Sim. de Tratamiento (Simulation treatment)",
    "Opciones (Options)",
    "Sensor de Cond (Conductivity sensor)",
    "Bomba Ceramica (Ceramic pump)",
    "Bomba de Heparina (Syringe pump)",
    "Calibración de Conductividad",
    "Calibración de Temperatura",
    "Calibración de Presión",
    "Otros (Other)"
]

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
    CHECKLIST_MENU,
    REPUESTOS_OPCION,
    REPUESTOS_TEXTO,
    SATISFACCION,
    INGENIERO,
    FECHA,
    MONEDA,
) = range(21)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("👋 Usa /reporte para iniciar el formulario técnico.")

async def iniciar_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data["checklist_seleccionados"] = []
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("📄 Ingrese Consecutivo (ej: 0000008) o pulse Dejar vacío:", reply_markup=reply_markup)
    return CONSECUTIVO

async def get_consecutivo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["consecutivo"] = "" if txt == "Dejar vacío" else txt
    await update.message.reply_text("🏥 Ingrese Nombre de la Clínica o Contacto para autocompletar:", reply_markup=ReplyKeyboardRemove())
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
            f"💡 Historial encontrado:\n"
            f"🏥 Clínica: {previo['hospital']}\n"
            f"👤 Contacto: {previo['contacto']}\n"
            f"📞 Teléfono: {previo['telefono']}\n"
            f"📍 Dirección: {previo['direccion']}\n\n"
            f"¿Deseas autocompletar estos datos?",
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
        ["Mantenimiento Correctivo (Repair)"],
        ["Mantenimiento Preventivo (PM)"],
        ["Instalación (Installation)"],
        ["Diagnostico (Diagnostic / Inspection)"],
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
        ["Fallo Mecánico", "Fallo Hidráulico"],
        ["Fallo en el Circuito", "Fallo en Software"],
        ["Fallo en parte de sangre", "Fallo de montaje de pieza"],
        ["Fallo de desgaste rápido de pieza", "Otros Fallos"],
        ["Ninguno / Normal"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⚠️ Clasificación de Fallas:", reply_markup=reply_markup)
    return FALLA_TIPO

async def get_falla_tipo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["falla_tipo"] = update.message.text
    teclado = [["Dejar vacío"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("💡 Ingrese Motivo del Fallo y Solución:", reply_markup=reply_markup)
    return SOLUCION

def armar_teclado_checklist(seleccionados):
    botones = []
    for item in ITEMS_CHECKLIST:
        marca = "✔ " if item in seleccionados else "⬜ "
        botones.append([InlineKeyboardButton(f"{marca}{item[:30]}", callback_data=f"chk_{item}")])
    botones.append([InlineKeyboardButton("✨ SELECCIONAR TODO", callback_data="chk_ALL")])
    botones.append([InlineKeyboardButton("✅ LISTO / CONTINUAR", callback_data="chk_DONE")])
    return InlineKeyboardMarkup(botones)

async def get_solucion(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["solucion"] = "" if txt == "Dejar vacío" else txt
    
    markup = armar_teclado_checklist(context.user_data["checklist_seleccionados"])
    await update.message.reply_text(
        "📋 Lista de verificación:\nToca los ítems que deseas marcar con ✔ y al terminar pulsa 'LISTO / CONTINUAR':",
        reply_markup=markup
    )
    return CHECKLIST_MENU

async def checklist_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    sel = context.user_data.get("checklist_seleccionados", [])
    
    if data == "chk_DONE":
        await query.message.reply_text(f"✅ Se seleccionaron {len(sel)} ítems de la lista de verificación.")
        teclado = [["Dejar vacío"], ["Ingresar componentes reemplazados"]]
        reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
        await query.message.reply_text("⚙️ Registro de reemplazo de componentes:", reply_markup=reply_markup)
        return REPUESTOS_OPCION
        
    elif data == "chk_ALL":
        context.user_data["checklist_seleccionados"] = list(ITEMS_CHECKLIST)
    else:
        item = data.replace("chk_", "")
        if item in sel:
            sel.remove(item)
        else:
            sel.append(item)
        context.user_data["checklist_seleccionados"] = sel

    await query.edit_message_reply_markup(reply_markup=armar_teclado_checklist(context.user_data["checklist_seleccionados"]))
    return CHECKLIST_MENU

async def get_repuestos_opcion(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text == "Dejar vacío":
        context.user_data["repuestos"] = []
        teclado = [
            ["Satisfecho (Satisfied)"],
            ["Relativamente satisfecho"],
            ["Normal (Normal)"],
            ["Insatisfecho (Dissatisfied)"],
            ["Muy Insatisfecho (Very Dissatisfied)"]
        ]
        reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
        await update.message.reply_text("⭐ Encuesta de satisfacción / Opinión de usuario:", reply_markup=reply_markup)
        return SATISFACCION
    else:
        await update.message.reply_text(
            "Escriba los componentes en este formato:\nNombre de la parte | Cantidad | Observación\n(ej: Válvula solenoide | 2 | Nueva)",
            reply_markup=ReplyKeyboardRemove()
        )
        return REPUESTOS_TEXTO

async def get_repuestos_texto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["repuestos_txt"] = update.message.text
    teclado = [
        ["Satisfecho (Satisfied)"],
        ["Relativamente satisfecho"],
        ["Normal (Normal)"],
        ["Insatisfecho (Dissatisfied)"],
        ["Muy Insatisfecho (Very Dissatisfied)"]
    ]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("⭐ Encuesta de satisfacción / Opinión de usuario:", reply_markup=reply_markup)
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
        await update.message.reply_text("Escriba el nombre del ingeniero:", reply_markup=ReplyKeyboardRemove())
        return INGENIERO
    context.user_data["ingeniero"] = txt
    
    fecha_hoy = datetime.now().strftime("%d/%m/%Y")
    teclado = [[fecha_hoy], ["Ingresar otra fecha"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("📅 Fecha del reporte (seleccione o edite):", reply_markup=reply_markup)
    return FECHA

async def get_fecha(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    if txt == "Ingresar otra fecha":
        await update.message.reply_text("Escriba la fecha (ej: 24/09/2026):", reply_markup=ReplyKeyboardRemove())
        return FECHA
    context.user_data["fecha"] = txt
    
    teclado = [["Dejar vacío"], ["Soles (S/.)"], ["USD ($)"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("💵 Moneda y cobro:", reply_markup=reply_markup)
    return MONEDA

def reemplazar_en_parrafos(contenedor, texto_buscar, texto_reemplazo):
    for p in contenedor.paragraphs:
        if texto_buscar.lower() in p.text.lower():
            for run in p.runs:
                if texto_buscar.lower() in run.text.lower():
                    run.text = run.text.replace(texto_buscar, texto_reemplazo)
                    return True
            p.text = p.text.replace(texto_buscar, texto_reemplazo)
            return True
    return False

async def get_moneda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["moneda"] = "" if txt == "Dejar vacío" else txt
    
    await update.message.reply_text("⏳ Procesando reporte en Word y manteniendo diseño intacto...")
    
    plantilla = "1-TECHNICAL SERVICE REPORT.docx"
    if not os.path.exists(plantilla):
        await update.message.reply_text("⚠️ No se encontró la plantilla .docx.")
        return ConversationHandler.END

    doc = docx.Document(plantilla)
    
    consecutivo = context.user_data.get("consecutivo", "")
    hosp = context.user_data.get("hospital", "")
    contacto = context.user_data.get("contacto", "")
    telefono = context.user_data.get("telefono", "")
    direccion = context.user_data.get("direccion", "")
    modelo = context.user_data.get("modelo", "")
    serie = context.user_data.get("serie", "")
    horometro = context.user_data.get("horometro", "")
    version_sw = context.user_data.get("version_sw", "")
    detalles = context.user_data.get("detalles", "")
    solucion = context.user_data.get("solucion", "")
    tipo_serv = context.user_data.get("tipo_servicio", "")
    falla_tipo = context.user_data.get("falla_tipo", "")
    satisfaccion = context.user_data.get("satisfaccion", "")
    checklist_sel = context.user_data.get("checklist_seleccionados", [])
    repuestos_txt = context.user_data.get("repuestos_txt", "")
    ingeniero = context.user_data.get("ingeniero", "Jesus Guillermo Pascual chalan")
    fecha_reporte = context.user_data.get("fecha", datetime.now().strftime("%d/%m/%Y"))

    # 1. Consecutivo
    if consecutivo:
        for p in doc.paragraphs:
            if "consecutivo" in p.text.lower():
                p.text = f"Consecutivo: {consecutivo}"
                break

    # 2. Recorrer celdas de las tablas con mapeo inteligente
    for t in doc.tables:
        for r in t.rows:
            # Identificar fila de cliente
            for i, c in enumerate(r.cells):
                txt_c = c.text.strip().lower()
                
                # Nombre del cliente
                if "nombre del cliente" in txt_c and i + 1 < len(r.cells):
                    r.cells[i + 1].text = hosp
                
                # Nombre del contacto (sin borrar etiqueta de teléfono)
                if "nombre de contacto" in txt_c and "telefono" not in txt_c:
                    if i + 1 < len(r.cells) and "telefono" not in r.cells[i + 1].text.lower():
                        r.cells[i + 1].text = contacto
                
                # Teléfono: buscar la celda específica de valor
                if "telefono" in txt_c or "number phone" in txt_c:
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = telefono

                # Dirección
                if "dirección" in txt_c or "adress" in txt_c:
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = direccion

                # Modelo (sin tocar celda de versión de software)
                if "modelo" in txt_c and "versión" not in txt_c:
                    if i + 1 < len(r.cells) and "versión" not in r.cells[i + 1].text.lower():
                        r.cells[i + 1].text = modelo

                # Versión de software: escribir al lado derecho de su título
                if "versión de software" in txt_c or "operating version" in txt_c:
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = version_sw

                # Serial No.
                if "número de serial" in txt_c or "serial no" in txt_c:
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = serie

                # Horómetro
                if "horometro" in txt_c or "running hours" in txt_c:
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = horometro

                # Detalles / Feedback Details
                if "detalles" in txt_c or "feedback details" in txt_c:
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = detalles

                # Motivo del Fallo y Solución
                if "motivo del fallo" in txt_c or "fault reason" in txt_c:
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = solucion

                # Tipo de Servicio: insertar [X] exacto sin descuadrar
                if "tipo de servicio" in txt_c or "service type" in txt_c or "correctivo" in txt_c:
                    for opcion in ["Mantenimiento Correctivo (Repair)", "Mantenimiento Preventivo (PM)", "Instalación (Installation)", "Diagnostico", "Entrenamiento", "Seguimiento", "Otro (Other)"]:
                        if opcion.split()[0].lower() in tipo_serv.lower() and opcion.split()[0].lower() in c.text.lower():
                            c.text = c.text.replace(opcion, f"[X] {opcion}")

                # Clasificación de fallas: insertar [X]
                if falla_tipo and falla_tipo != "Ninguno / Normal":
                    for p in c.paragraphs:
                        if falla_tipo.lower() in p.text.lower():
                            p.text = p.text.replace(falla_tipo, f"[X] {falla_tipo}")

                # Lista de verificación: insertar check ✔ en los elegidos
                for item_chk in checklist_sel:
                    nombre_corto = item_chk.split("(")[0].strip()
                    if nombre_corto.lower() in c.text.lower():
                        for p in c.paragraphs:
                            if nombre_corto.lower() in p.text.lower():
                                p.text = f"✔ {p.text}"

                # Satisfacción del usuario: insertar [X]
                if satisfaccion:
                    sat_clave = satisfaccion.split("(")[0].strip()
                    if sat_clave.lower() in c.text.lower():
                        for p in c.paragraphs:
                            if sat_clave.lower() in p.text.lower():
                                p.text = p.text.replace(sat_clave, f"[X] {sat_clave}")

                # Reemplazo de componentes
                if repuestos_txt and ("registro de reemplazo" in txt_c or "component replacement" in txt_c):
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = repuestos_txt

                # Clientes / Ingeniero y Fechas al pie
                if "customer name" in txt_c:
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = contacto
                if "engineer name" in txt_c:
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = ingeniero
                if "signature date" in txt_c or "fecha de firma" in txt_c:
                    if i + 1 < len(r.cells):
                        r.cells[i + 1].text = fecha_reporte

    # 3. Insertar firma digital (busca cualquier nombre de archivo que exista)
    archivo_firma = None
    for posible in ["Code_Generated_Image.png", "firma_transparente.png", "firma.png"]:
        if os.path.exists(posible):
            archivo_firma = posible
            break

    if archivo_firma:
        for t in doc.tables:
            for r_idx, r in enumerate(t.rows):
                for c_idx, c in enumerate(r.cells):
                    # Celda que dice Firma (Engineer Signature)
                    if "engineer signature" in c.text.lower() or "firma\n(engineer signature)" in c.text.lower():
                        # Si está vacía o es la celda de la firma del ingeniero (columna derecha)
                        if c_idx >= len(r.cells) // 2:
                            c.text = ""
                            p = c.paragraphs[0]
                            p.add_run().add_picture(archivo_firma, width=Inches(1.3))
                            break

    # Guardar documento final
    nombre_docx = f"Reporte_{serie if serie else 'Servicio'}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    doc.save(nombre_docx)

    with open(nombre_docx, "rb") as f:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=f,
            caption="✅ ¡Reporte generado con formato respetado, datos organizados y firma insertada!"
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
            CHECKLIST_MENU: [CallbackQueryHandler(checklist_callback)],
            REPUESTOS_OPCION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_repuestos_opcion)],
            REPUESTOS_TEXTO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_repuestos_texto)],
            SATISFACCION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_satisfaccion)],
            INGENIERO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_ingeniero)],
            FECHA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_fecha)],
            MONEDA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_moneda)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv)
    
    logger.info("Bot de Reportes listo y escuchando...")
    app.run_polling()

if __name__ == "__main__":
    main()
