# -*- coding: utf-8 -*-
"""
excel_client.py
================
Camada de acesso a dados da aplicação REDEB2B — versão simplificada, que
opera exclusivamente sobre um arquivo Excel local (.xlsx), usando
pandas + openpyxl.

Expõe a classe `DataClient` com uma interface de CRUD:
  list_items, get_item, create_item, update_item, delete_item,
  dashboard_aggregates, items_by_date
"""

import os
import logging
import re
import unicodedata
from datetime import datetime

import pandas as pd

logger = logging.getLogger("redeb2b.excel_client")

# Colunas oficiais da planilha REDEB2B, na ordem definida no escopo.
FIELDS = [
    "IDCLIENTE", "CLIENTE", "ENDERECO", "CIDADE", "PRODUTO", "ATIVIDADE",
    "TECNOLOGIA", "VT", "DATADISPARO", "RETORNOPCC", "DATAAGENDAMENTO",
    "DATACONCLUSAO", "OBSERVACAO", "STATUS", "EXECUTADOPOR", "TIPOCABO",
    "METRAGEM", "OBSERVACAOCONCLUSAO", "NUMDRAFT", "ROTA", "USUARIO",
]

# Campos de data que precisam de tratamento especial (ISO yyyy-mm-dd).
DATE_FIELDS = {"DATADISPARO", "DATAAGENDAMENTO", "DATACONCLUSAO"}

# Padrões aceitos ao LER datas do Excel: ISO (yyyy-mm-dd, com ou sem hora) e
# pt-BR (dd/mm/yyyy) — este último é comum quando alguém edita a data
# diretamente na planilha, com o Excel configurado em português.
_ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})")
_BR_DATE_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})")


class DataClientError(Exception):
    """Erro de negócio/infra do DataClient — sempre com uma mensagem amigável."""

    def __init__(self, message, status_code=400):
        super().__init__(message)
        self.message = message
        self.status_code = status_code


def _strip_accents(value: str) -> str:
    """Remove acentos para permitir buscas 'case/acento-insensitive'."""
    if value is None:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(value))
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower()


def _parse_date(value):
    """Converte valores variados (str ISO, str dd/mm/aaaa, datetime, Timestamp
    do pandas) em uma data ISO 'yyyy-mm-dd' (string) para permitir comparação
    e ordenação consistentes, independentemente de como a data foi digitada
    na planilha. Retorna None se o valor estiver vazio ou não reconhecível."""
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d")

    s = str(value).strip()
    if not s:
        return None

    m = _ISO_DATE_RE.match(s)
    if m:
        year, month, day = m.groups()
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"

    m = _BR_DATE_RE.match(s)
    if m:
        day, month, year = m.groups()
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"

    # Formato não reconhecido — retorna como veio, truncado, para não quebrar
    # o restante do fluxo (mas não vai comparar corretamente em filtros de data).
    return s[:10]


class DataClient:
    """Cliente de dados baseado exclusivamente em um arquivo Excel local."""

    def __init__(self):
        self.excel_path = os.getenv("EXCEL_PATH", "")
        self.excel_sheet = os.getenv("EXCEL_SHEET_NAME", "TbRelatorio")

    # ------------------------------------------------------------------
    # Métodos públicos
    # ------------------------------------------------------------------
    def list_items(self, filters=None, page=1, page_size=20, sort=None):
        """Lista itens com filtros, paginação e ordenação.

        filters: dict com chaves opcionais: cliente, id, cidade,
                 executadopor, status, data_inicio, data_fim, q
        sort: string no formato "campo:asc" ou "campo:desc"
        Retorna: dict {items: [...], total: int, page: int, page_size: int}
        """
        filters = filters or {}
        rows = self._excel_read_all()
        rows = self._apply_filters(rows, filters)
        rows = self._apply_sort(rows, sort)

        total = len(rows)
        start = (page - 1) * page_size
        end = start + page_size
        page_items = rows[start:end]

        return {
            "items": page_items,
            "total": total,
            "page": page,
            "page_size": page_size,
        }

    def get_item(self, item_id):
        rows = self._excel_read_all()
        for row in rows:
            if str(row.get("IDCLIENTE")) == str(item_id):
                return row
        raise DataClientError(f"Registro com IDCLIENTE={item_id} não encontrado.", status_code=404)

    def create_item(self, data):
        data = self._sanitize_and_validate(data, is_new=True)
        rows = self._excel_read_all()
        if any(str(r.get("IDCLIENTE")) == str(data.get("IDCLIENTE")) for r in rows):
            raise DataClientError(f"Já existe um registro com IDCLIENTE={data.get('IDCLIENTE')}.", status_code=409)
        new_row = {f: data.get(f, "") for f in FIELDS}
        rows.append(new_row)
        self._excel_write_all(rows)
        logger.info("Registro criado: IDCLIENTE=%s", data.get("IDCLIENTE"))
        return new_row

    def update_item(self, item_id, data):
        data = self._sanitize_and_validate(data, is_new=False)
        rows = self._excel_read_all()
        found = False
        for row in rows:
            if str(row.get("IDCLIENTE")) == str(item_id):
                row.update(data)
                found = True
                break
        if not found:
            raise DataClientError(f"Registro com IDCLIENTE={item_id} não encontrado.", status_code=404)
        self._excel_write_all(rows)
        logger.info("Registro atualizado: IDCLIENTE=%s", item_id)
        return self.get_item(item_id)

    def delete_item(self, item_id):
        rows = self._excel_read_all()
        new_rows = [r for r in rows if str(r.get("IDCLIENTE")) != str(item_id)]
        if len(new_rows) == len(rows):
            raise DataClientError(f"Registro com IDCLIENTE={item_id} não encontrado.", status_code=404)
        self._excel_write_all(new_rows)
        logger.info("Registro excluído: IDCLIENTE=%s", item_id)
        return {"deleted": item_id}

    def dashboard_aggregates(self, filters=None):
        """Retorna agregações prontas para Chart.js + KPIs, respeitando filtros opcionais."""
        filters = filters or {}
        rows = self._excel_read_all()
        rows = self._apply_filters(rows, filters)

        def group_count(field):
            counts = {}
            for row in rows:
                key = row.get(field) or "Não informado"
                counts[key] = counts.get(key, 0) + 1
            sorted_items = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:12]
            return {
                "labels": [k for k, _ in sorted_items],
                "data": [v for _, v in sorted_items],
            }

        status_counts = group_count("STATUS")
        kpis = {
            "total": len(rows),
            "por_status": {k: v for k, v in zip(status_counts["labels"], status_counts["data"])},
        }

        concluidos = [r for r in rows if _strip_accents(r.get("STATUS")) == _strip_accents("Concluído")]

        if filters:
            # Há filtro(s) ativo(s): mostra os concluídos conforme o filtro aplicado
            # (rows já está filtrado por _apply_filters acima).
            concluidos_periodo = concluidos
            kpis["concluidos_label"] = "Concluídos (filtro)"
        else:
            # Sem filtro: por padrão, mostra apenas os concluídos do mês corrente.
            hoje = datetime.today()
            ano_atual, mes_atual = hoje.year, hoje.month

            def _no_mes_atual(row):
                d = _parse_date(row.get("DATACONCLUSAO"))
                if not d:
                    return False
                ano, mes, _dia = d.split("-")
                return int(ano) == ano_atual and int(mes) == mes_atual

            concluidos_periodo = [r for r in concluidos if _no_mes_atual(r)]
            kpis["concluidos_label"] = "Concluídos no mês"

        kpis["concluidos_total"] = len(concluidos_periodo)
        kpis["novos_total"] = sum(
            1 for r in rows if _strip_accents(r.get("STATUS")) == _strip_accents("Novo")
        )

        return {
            "kpis": kpis,
            "por_cliente": group_count("CLIENTE"),
            "por_cidade": group_count("CIDADE"),
            "por_executadopor": group_count("EXECUTADOPOR"),
            "por_status": status_counts,
            "por_data_conclusao": self._group_by_date(concluidos_periodo, "DATACONCLUSAO"),
        }

    def items_by_date(self, date_str, date_field="DATAAGENDAMENTO", status=None):
        """Retorna todos os itens cujo campo de data indicado (por padrão
        DATAAGENDAMENTO, mas também aceita DATACONCLUSAO) seja igual a
        date_str (YYYY-MM-DD), opcionalmente filtrando por STATUS.
        Já projetados apenas com as colunas usadas no modal de detalhe."""
        if date_field not in DATE_FIELDS:
            raise DataClientError(f"Campo de data inválido: {date_field}")
        rows = self._excel_read_all()
        cols = ["IDCLIENTE", "CLIENTE", "ENDERECO", "CIDADE", "TECNOLOGIA", "VT",
                "DATAAGENDAMENTO", "DATACONCLUSAO", "STATUS", "TIPOCABO",
                "METRAGEM", "NUMDRAFT", "ROTA"]
        result = []
        for row in rows:
            if _parse_date(row.get(date_field)) != date_str:
                continue
            if status and _strip_accents(row.get("STATUS")) != _strip_accents(status):
                continue
            result.append({c: row.get(c) for c in cols})
        return result

    # ------------------------------------------------------------------
    # Filtros / ordenação
    # ------------------------------------------------------------------
    def _apply_filters(self, rows, filters):
        cliente = filters.get("cliente")
        idcliente = filters.get("id")
        cidade = filters.get("cidade")
        executadopor = filters.get("executadopor")
        status = filters.get("status")
        data_inicio = filters.get("data_inicio")
        data_fim = filters.get("data_fim")
        busca = filters.get("q")

        def match(row):
            if cliente and _strip_accents(cliente) not in _strip_accents(row.get("CLIENTE")):
                return False
            if idcliente and str(idcliente) not in str(row.get("IDCLIENTE", "")):
                return False
            if cidade and _strip_accents(cidade) not in _strip_accents(row.get("CIDADE")):
                return False
            if executadopor and _strip_accents(executadopor) not in _strip_accents(row.get("EXECUTADOPOR")):
                return False
            if status and _strip_accents(status) != _strip_accents(row.get("STATUS")):
                return False
            if data_inicio or data_fim:
                d = _parse_date(row.get("DATAAGENDAMENTO"))
                if not d:
                    return False
                if data_inicio and d < data_inicio:
                    return False
                if data_fim and d > data_fim:
                    return False
            if busca:
                haystack = " ".join(str(row.get(f, "")) for f in FIELDS)
                if _strip_accents(busca) not in _strip_accents(haystack):
                    return False
            return True

        return [r for r in rows if match(r)]

    def _apply_sort(self, rows, sort):
        if not sort:
            return rows
        try:
            field, direction = sort.split(":")
        except ValueError:
            field, direction = sort, "asc"
        field = field.upper()
        if field not in FIELDS:
            return rows
        reverse = direction.lower() == "desc"

        def sort_key(row):
            value = row.get(field)
            if value is None:
                return ""
            return str(value).lower()

        return sorted(rows, key=sort_key, reverse=reverse)

    def _group_by_date(self, rows, field):
        counts = {}
        for row in rows:
            d = _parse_date(row.get(field))
            if not d:
                continue
            counts[d] = counts.get(d, 0) + 1
        ordered = sorted(counts.items(), key=lambda kv: kv[0])
        return {"labels": [k for k, _ in ordered], "data": [v for _, v in ordered]}

    def _sanitize_and_validate(self, data, is_new):
        """Sanitiza strings e valida formato de datas. Lança DataClientError se inválido."""
        clean = {}
        for field in FIELDS:
            if field not in data:
                continue
            value = data[field]
            if isinstance(value, str):
                value = value.strip()
            if field in DATE_FIELDS and value:
                try:
                    datetime.strptime(str(value)[:10], "%Y-%m-%d")
                except ValueError:
                    raise DataClientError(
                        f"Campo {field} deve estar no formato ISO yyyy-mm-dd. Valor recebido: {value}"
                    )
                value = str(value)[:10]
            clean[field] = value

        if is_new and not clean.get("IDCLIENTE"):
            raise DataClientError("IDCLIENTE é obrigatório para criar um registro.")
        if is_new and not clean.get("CLIENTE"):
            raise DataClientError("CLIENTE é obrigatório para criar um registro.")
        return clean

    # ------------------------------------------------------------------
    # Leitura / escrita do Excel
    # ------------------------------------------------------------------
    def _excel_read_all(self):
        if not self.excel_path or not os.path.exists(self.excel_path):
            raise DataClientError(
                f"Arquivo Excel não encontrado em '{self.excel_path}'. "
                "Verifique a variável EXCEL_PATH no .env.",
                status_code=500,
            )
        try:
            df = pd.read_excel(self.excel_path, sheet_name=self.excel_sheet, dtype=str)
        except Exception as exc:  # noqa: BLE001 - queremos capturar qualquer erro de leitura
            logger.exception("Erro lendo Excel")
            raise DataClientError(f"Erro ao ler o arquivo Excel: {exc}", status_code=500)

        for field in FIELDS:
            if field not in df.columns:
                df[field] = ""
        df = df.fillna("")
        rows = df.to_dict(orient="records")

        # Normaliza datas para ISO (yyyy-mm-dd), aceitando tanto o formato
        # ISO quanto dd/mm/aaaa (comum quando editado manualmente no Excel
        # em pt-BR). Isso garante que filtros, ordenação e exibição no
        # frontend funcionem de forma consistente, seja qual for o formato
        # em que a data foi digitada na planilha.
        for row in rows:
            for field in DATE_FIELDS:
                normalized = _parse_date(row.get(field))
                row[field] = normalized or ""

        return rows

    def _excel_write_all(self, rows):
        df = pd.DataFrame(rows, columns=FIELDS)
        try:
            df.to_excel(self.excel_path, sheet_name=self.excel_sheet, index=False)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Erro escrevendo Excel")
            raise DataClientError(f"Erro ao gravar o arquivo Excel: {exc}", status_code=500)
