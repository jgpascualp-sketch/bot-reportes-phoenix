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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
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

ITEMS_CHECKLIST = [
    "Apariencia",
    "Bateria de respaldo",
    "Placa Hall",
    "Señal de sensor",
    "Pantalla táctil",
    "Pieza hidráulica",
    "Parametros de Tratamiento",
    "Sim. de Tratamiento",
    "Opciones",
    "Sensor de Cond",
    "Bomba Ceramica",
    "Bomba de Heparina",
    "Calibración de Conductividad",
    "Calibración de Temperatura",
    "Calibración de Presión",
    "Otros"
]

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
    SATISFACCION,
    INGENIERO,
    OPCION_FIRMA,
    SUBIR_FIRMA,
    FECHA,
    MONEDA,
) = range(21)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("👋 Usa /reporte para generar un nuevo reporte técnico.")

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
    await update.message.reply_text("🏥 Ingrese Nombre de la Clínica o Contacto para buscar historial:", reply_markup=ReplyKeyboardRemove())
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
        ["Fallo Hidráulico", "Fallo en el Circuito"],
        ["Fallo Mecánico", "Fallo en Software"],
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
    fila = []
    for item in ITEMS_CHECKLIST:
        marca = "✔ " if item in seleccionados else "⬜ "
        fila.append(InlineKeyboardButton(f"{marca}{item}", callback_data=f"chk_{item}"))
        if len(fila) == 2:
            botones.append(fila)
            fila = []
    if fila:
        botones.append(fila)
    botones.append([InlineKeyboardButton("✨ SELECCIONAR TODO", callback_data="chk_ALL")])
    botones.append([InlineKeyboardButton("✅ LISTO / CONTINUAR", callback_data="chk_DONE")])
    return InlineKeyboardMarkup(botones)

async def get_solucion(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["solucion"] = "" if txt == "Dejar vacío" else txt
    
    markup = armar_teclado_checklist(context.user_data["checklist_seleccionados"])
    await update.message.reply_text(
        "📋 Lista de verificación de pruebas:\nToque los ítems que desea marcar con ✔ y presione 'LISTO / CONTINUAR':",
        reply_markup=markup
    )
    return CHECKLIST_MENU

async def checklist_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    sel = context.user_data.get("checklist_seleccionados", [])
    
    if data == "chk_DONE":
        await query.message.reply_text(f"✅ Se seleccionaron {len(sel)} comprobaciones.")
        teclado = [
            ["Satisfecho (Satisfied)"],
            ["Relativamente satisfecho"],
            ["Normal (Normal)"],
            ["Insatisfecho (Dissatisfied)"],
            ["Muy Insatisfecho (Very Dissatisfied)"]
        ]
        reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
        await query.message.reply_text("⭐ Encuesta de satisfacción / Opinión de usuario:", reply_markup=reply_markup)
        return SATISFACCION
        
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
    
    teclado = [["🖊️ Firma Automática (Jesús Pascual)"], ["📷 Subir Foto de Firma"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("🖋️ ¿Cómo desea estampar la firma?", reply_markup=reply_markup)
    return OPCION_FIRMA

async def get_opcion_firma(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    if "Subir Foto" in txt:
        await update.message.reply_text("📸 Envíe la FOTO de la firma:", reply_markup=ReplyKeyboardRemove())
        return SUBIR_FIRMA
    
    context.user_data["firma_custom"] = None
    fecha_hoy = datetime.now().strftime("%d/%m/%Y")
    teclado = [[fecha_hoy], ["Ingresar otra fecha"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("📅 Fecha del reporte (seleccione o edite):", reply_markup=reply_markup)
    return FECHA

async def get_foto_firma(update: Update, context: ContextTypes.DEFAULT_TYPE):
    foto = await update.message.photo[-1].get_file()
    ruta_temp = "firma_subida.png"
    await foto.download_to_drive(ruta_temp)
    context.user_data["firma_custom"] = ruta_temp
    
    fecha_hoy = datetime.now().strftime("%d/%m/%Y")
    teclado = [[fecha_hoy], ["Ingresar otra fecha"]]
    reply_markup = ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("✅ Firma recibida. 📅 Confirme la fecha del reporte:", reply_markup=reply_markup)
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

def marcar_cuadro_exacto(celda, palabra_clave, simbolo="☒"):
    """
    Busca la palabra clave en la celda y transforma el cuadrito asociado
    en el símbolo deseado (☒ para X, ☑ para check) sin reescribir ni mover la celda.
    """
    clave = palabra_clave.lower().strip()
    cajas = ["□", "☐", "⬜", "[ ]", "[]", "\u25a1", "\u25fb", "\u2610", "\uf06f", "o"]
    
    for p in celda.paragraphs:
        if clave in p.text.lower():
            # 1. Controles w:checkBox / w14:checkbox XML
            for cb in p._element.xpath('.//w:checkBox | .//w14:checkbox'):
                chk = cb.find(qn('w:checked')) or cb.find(qn('w14:checked'))
                if chk is None:
                    chk = OxmlElement('w:checked')
                    cb.append(chk)
                chk.set(qn('w:val'), "1")
                return True
                
            # 2. Reemplazo en los runs del párrafo para preservar fuente y espaciado
            for r in p.runs:
                for cj in cajas:
                    if cj in r.text:
                        r.text = r.text.replace(cj, simbolo, 1)
                        return True
                        
            # 3. Si el carácter de caja estaba en el párrafo general
            for cj in cajas:
                if cj in p.text:
                    p.text = p.text.replace(cj, simbolo, 1)
                    return True
                    
            # 4. Respaldo: agregar el símbolo al inicio de la línea sin descuadrar
            p.text = f"{simbolo} {p.text.strip()}"
            return True
            
    return False

async def get_moneda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    context.user_data["moneda"] = "" if txt == "Dejar vacío" else txt
    
    await update.message.reply_text("⏳ Procesando reporte...")
    
    plantilla = "1-TECHNICAL SERVICE REPORT.docx"
    if not os.path.exists(plantilla):
        await update.message.reply_text("⚠️ No se encontró la plantilla .docx.")
        return ConversationHandler.END

    # Cargar plantilla limpia para preservar diseño original intacto
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
    ingeniero = context.user_data.get("ingeniero", "Jesus Guillermo Pascual chalan")
    fecha_reporte = context.user_data.get("fecha", datetime.now().strftime("%d/%m/%Y"))

    # Consecutivo
    if consecutivo:
        for p in doc.paragraphs:
            if "consecutivo" in p.text.lower():
                p.text = f"Consecutivo: {consecutivo}"
                break

    t = doc.tables[0]

    # Fila 0: Hospital
    for c in t.rows[0].cells:
        if "hospital name" in c.text.lower():
            continue
        c.text = hosp
        break

    # Fila 1: Contacto y Teléfono
    r1 = t.rows[1].cells
    r1[1].text = contacto
    r1[-1].text = telefono

    # Fila 2: Dirección
    for c in t.rows[2].cells:
        if "dirección" in c.text.lower() or "adress" in c.text.lower():
            continue
        c.text = direccion
        break

    # Fila 3: Modelo y Versión de Software
    r3 = t.rows[3].cells
    r3[1].text = modelo
    r3[-1].text = version_sw

    # Fila 4: Serie
    t.rows[4].cells[1].text = serie

    # Fila 5: Horómetro
    t.rows[5].cells[1].text = horometro

    # Fila 6: Tipo de Servicio (Preservar la estructura original y solo marcar el recuadro con ☒)
    c_serv = t.rows[6].cells[-1] if len(t.rows[6].cells) > 1 else t.rows[6].cells[0]
    palabra_ts = tipo_serv.split()[0]
    marcar_cuadro_exacto(c_serv, palabra_ts, "☒")

    # Fila 7: Detalles
    c_det = t.rows[7].cells[-1] if len(t.rows[7].cells) > 1 else t.rows[7].cells[0]
    c_det.text = detalles

    # Fila 8: Clasificación de Fallas (Marcar ☒ en el recuadro de la opción elegida)
    if falla_tipo and falla_tipo != "Ninguno / Normal":
        for c in t.rows[8].cells:
            marcar_cuadro_exacto(c, falla_tipo, "☒")

    # Fila 9: Motivo del Fallo y Solución (Escribir en el cuadro blanco grande de solución)
    c_sol = t.rows[9].cells[-1] if len(t.rows[9].cells) > 1 else t.rows[9].cells[0]
    c_sol.text = solucion

    # Fila 10: Lista de verificación (Marcar con ☑ check en los ítems seleccionados)
    for c in t.rows[10].cells:
        for chk in checklist_sel:
            marcar_cuadro_exacto(c, chk, "☑")

    # Fila 12: Encuesta de satisfacción (Marcar ☒ en la opción elegida)
    if satisfaccion:
        c_sat = t.rows[12].cells[-1] if len(t.rows[12].cells) > 1 else t.rows[12].cells[0]
        palabra_sat = satisfaccion.split()[0]
        marcar_cuadro_exacto(c_sat, palabra_sat, "☒")

    # Tabla 1: Firmas y Nombres
    t2 = doc.tables[1]
    t2.rows[0].cells[1].text = contacto
    t2.rows[0].cells[3].text = ingeniero
    
    # Fechas centradas simétricamente
    cell_fec_cli = t2.rows[2].cells[1]
    cell_fec_cli.text = fecha_reporte
    cell_fec_cli.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell_fec_cli.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    cell_fec_ing = t2.rows[2].cells[3]
    cell_fec_ing.text = fecha_reporte
    cell_fec_ing.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell_fec_ing.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Firma: automática o imagen subida
    archivo_firma = context.user_data.get("firma_custom")
    if not archivo_firma:
        for f_nom in ["Code_Generated_Image.png", "firma_transparente.png", "firma.png"]:
            if os.path.exists(f_nom):
                archivo_firma = f_nom
                break

    if archivo_firma and os.path.exists(archivo_firma):
        cell_sig = t2.rows[1].cells[3]
        cell_sig.text = ""
        cell_sig.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell_sig.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(archivo_firma, width=Inches(1.2))

    nombre_docx = f"Reporte_{serie if serie else 'Servicio'}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    doc.save(nombre_docx)

    with open(nombre_docx, "rb") as f:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=f,
            caption="✅ Reporte generado: formato original intacto, solución en su celda, marcas ☒ y ☑ exactas y fechas centradas."
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
            SATISFACCION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_satisfaccion)],
            INGENIERO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_ingeniero)],
            OPCION_FIRMA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_opcion_firma)],
            SUBIR_FIRMA: [MessageHandler(filters.PHOTO, get_foto_firma)],
            FECHA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_fecha)],
            MONEDA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_moneda)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv)
    
    logger.info("Bot de Reportes listo.")
    app.run_polling()

if __name__ == "__main__":
    main()
