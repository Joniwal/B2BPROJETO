# -*- coding: utf-8 -*-
"""
app.py
======
Ponto de entrada da aplicação REDEB2B. Inicializa o Flask, carrega variáveis
de ambiente (.env), registra o blueprint da API e as rotas de página, além de
handlers de erro globais.
"""

import os
import logging
from flask import Flask, render_template, jsonify
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("redeb2b.app")


def create_app():
    app = Flask(__name__)
    app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-key-troque-em-producao")

    from api import api_bp
    app.register_blueprint(api_bp)

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.errorhandler(404)
    def not_found(_error):
        return jsonify({"ok": False, "error": "Rota não encontrada."}), 404

    @app.errorhandler(500)
    def internal_error(error):
        logger.exception("Erro interno não tratado: %s", error)
        return jsonify({"ok": False, "error": "Erro interno do servidor."}), 500

    return app


app = create_app()

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    debug = os.getenv("FLASK_DEBUG", "true").strip().lower() == "true"
    logger.info("Iniciando REDEB2B Flask app na porta %s (debug=%s)", port, debug)
    app.run(host="0.0.0.0", port=port, debug=debug)
