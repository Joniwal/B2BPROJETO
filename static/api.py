# -*- coding: utf-8 -*-
"""
api.py
======
Blueprint REST da aplicação REDEB2B. Todas as respostas são JSON e seguem o
padrão:
  sucesso: { "ok": true, "data": ... }
  erro:    { "ok": false, "error": "mensagem amigável" }
"""

import io
import logging
from datetime import datetime

import pandas as pd
from flask import Blueprint, request, jsonify, send_file

from excel_client import DataClient, DataClientError, FIELDS

logger = logging.getLogger("redeb2b.api")
api_bp = Blueprint("api", __name__, url_prefix="/api")

data_client = DataClient()


def _error_response(exc: DataClientError):
    return jsonify({"ok": False, "error": exc.message}), exc.status_code


def _parse_pagination_and_filters():
    filters = {
        "cliente": request.args.get("cliente"),
        "id": request.args.get("id"),
        "cidade": request.args.get("cidade"),
        "executadopor": request.args.get("executadopor"),
        "status": request.args.get("status"),
        "data_inicio": request.args.get("dataInicio"),
        "data_fim": request.args.get("dataFim"),
        "q": request.args.get("q"),
    }
    filters = {k: v for k, v in filters.items() if v}
    page = int(request.args.get("page", 1))
    page_size = int(request.args.get("page_size", 20))
    sort = request.args.get("sort")
    return filters, page, page_size, sort


@api_bp.route("/items", methods=["GET"])
def get_items():
    try:
        filters, page, page_size, sort = _parse_pagination_and_filters()
        result = data_client.list_items(filters=filters, page=page, page_size=page_size, sort=sort)
        return jsonify({"ok": True, "data": result})
    except DataClientError as exc:
        return _error_response(exc)
    except Exception:  # noqa: BLE001
        logger.exception("Erro inesperado em GET /api/items")
        return jsonify({"ok": False, "error": "Erro interno ao listar registros."}), 500


@api_bp.route("/items/<item_id>", methods=["GET"])
def get_item(item_id):
    try:
        item = data_client.get_item(item_id)
        return jsonify({"ok": True, "data": item})
    except DataClientError as exc:
        return _error_response(exc)
    except Exception:  # noqa: BLE001
        logger.exception("Erro inesperado em GET /api/items/%s", item_id)
        return jsonify({"ok": False, "error": "Erro interno ao buscar registro."}), 500


@api_bp.route("/items", methods=["POST"])
def create_item():
    try:
        payload = request.get_json(force=True, silent=True) or {}
        item = data_client.create_item(payload)
        return jsonify({"ok": True, "data": item}), 201
    except DataClientError as exc:
        return _error_response(exc)
    except Exception:  # noqa: BLE001
        logger.exception("Erro inesperado em POST /api/items")
        return jsonify({"ok": False, "error": "Erro interno ao criar registro."}), 500


@api_bp.route("/items/<item_id>", methods=["PATCH"])
def update_item(item_id):
    try:
        payload = request.get_json(force=True, silent=True) or {}
        item = data_client.update_item(item_id, payload)
        return jsonify({"ok": True, "data": item})
    except DataClientError as exc:
        return _error_response(exc)
    except Exception:  # noqa: BLE001
        logger.exception("Erro inesperado em PATCH /api/items/%s", item_id)
        return jsonify({"ok": False, "error": "Erro interno ao atualizar registro."}), 500


@api_bp.route("/items/<item_id>", methods=["DELETE"])
def delete_item(item_id):
    try:
        result = data_client.delete_item(item_id)
        return jsonify({"ok": True, "data": result})
    except DataClientError as exc:
        return _error_response(exc)
    except Exception:  # noqa: BLE001
        logger.exception("Erro inesperado em DELETE /api/items/%s", item_id)
        return jsonify({"ok": False, "error": "Erro interno ao excluir registro."}), 500


@api_bp.route("/dashboard", methods=["GET"])
def get_dashboard():
    try:
        filters, _, _, _ = _parse_pagination_and_filters()
        result = data_client.dashboard_aggregates(filters=filters)
        return jsonify({"ok": True, "data": result})
    except DataClientError as exc:
        return _error_response(exc)
    except Exception:  # noqa: BLE001
        logger.exception("Erro inesperado em GET /api/dashboard")
        return jsonify({"ok": False, "error": "Erro interno ao gerar dashboard."}), 500


@api_bp.route("/items-by-date", methods=["GET"])
def get_items_by_date():
    """Usado pelo modal 'ver todos os itens de uma data' — funciona tanto para
    data de agendamento quanto para data de conclusão, via o parâmetro 'campo'."""
    try:
        date_str = request.args.get("data")
        date_field = request.args.get("campo", "DATAAGENDAMENTO")
        status = request.args.get("status")
        if not date_str:
            raise DataClientError("Parâmetro 'data' (YYYY-MM-DD) é obrigatório.")
        items = data_client.items_by_date(date_str, date_field=date_field, status=status)
        return jsonify({"ok": True, "data": items})
    except DataClientError as exc:
        return _error_response(exc)
    except Exception:  # noqa: BLE001
        logger.exception("Erro inesperado em GET /api/items-by-date")
        return jsonify({"ok": False, "error": "Erro interno ao buscar registros da data."}), 500


@api_bp.route("/export", methods=["GET"])
def export_items():
    """Gera e envia um arquivo .xlsx com exatamente os registros que
    atendem aos filtros/ordenação passados (os mesmos parâmetros aceitos
    por /api/items), sem paginação — ou seja, o que está filtrado na tela."""
    try:
        filters, _page, _page_size, sort = _parse_pagination_and_filters()
        rows = data_client.export_items(filters=filters, sort=sort)

        df = pd.DataFrame(rows, columns=FIELDS)
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False, sheet_name="REDEB2B")
        buffer.seek(0)

        filename = f"REDEB2B_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        return send_file(
            buffer,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except DataClientError as exc:
        return _error_response(exc)
    except Exception:  # noqa: BLE001
        logger.exception("Erro inesperado em GET /api/export")
        return jsonify({"ok": False, "error": "Erro interno ao gerar a exportação para Excel."}), 500
