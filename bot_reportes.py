import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

class SimpleHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Bot activo")

def run_http():
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), SimpleHandler)
    server.serve_forever()

def main():
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise ValueError("TELEGRAM_BOT_TOKEN no configurado")

    # Iniciar servidor web falso para mantener vivo el Web Service en Render
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
