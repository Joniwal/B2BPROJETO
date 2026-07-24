# -*- coding: utf-8 -*-
"""
sharepoint_client.py
=====================
Camada de acesso a dados da aplicação REDEB2B.

Este módulo expõe uma única classe, `DataClient`, com uma interface unificada
de CRUD (list_items, get_item, create_item, update_item, delete_item) que
pode operar em dois modos, escolhidos pela variável de ambiente
USE_EXCEL_FALLBACK:

  * MODO EXCEL (USE_EXCEL_FALLBACK=true):
      Lê/escreve diretamente em um arquivo .xlsx local usando pandas +
      openpyxl. Útil para desenvolvimento local sem depender de credenciais
      do Azure AD, ou como contingência caso o Graph esteja indisponível.

  * MODO SHAREPOINT/GRAPH (USE_EXCEL_FALLBACK=false):
      Autentica via MSAL (fluxo Client Credentials — aplicação-a-aplicação)
      e faz chamadas REST ao Microsoft Graph para operar sobre a lista do
      SharePoint (GET/POST/PATCH/DELETE em
      /sites/{site-id}/lists/{list-id}/items).

Por que Client Credentials e não Authorization Code?
-----------------------------------------------------
O backend Flask atua como um serviço "daemon": ele mesmo acessa o SharePoint
usando a identidade do aplicativo (não a de um usuário logado interativamente
em um navegador). Não há tela de login/consentimento por usuário — em vez
disso, um administrador do tenant concede consentimento de administrador uma
única vez para as permissões de aplicativo (Application permissions), como
Sites.ReadWrite.All. Esse é o fluxo recomendado pela Microsoft para
automações/serviços que precisam acessar recursos do Graph sem um usuário
presente. Caso a aplicação precise futuramente atuar "em nome do usuário"
(delegated permissions), seria necessário migrar para o fluxo Authorization
Code, o que exigiria REDIRECT_URI e uma tela de login — não implementado
aqui por não fazer parte do escopo solicitado.
"""

import os
import logging
import unicodedata
from datetime import datetime

import requests
import msal
import pandas as pd

logger = logging.getLogger("redeb2b.sharepoint_client")

# Colunas oficiais da lista/planilha REDEB2B, na ordem definida no escopo.
FIELDS = [
    "IDCLIENTE", "CLIENTE", "ENDERECO", "CIDADE", "PRODUTO", "ATIVIDADE",
    "TECNOLOGIA", "VT", "DATADISPARO", "RETORNOPCC", "DATAAGENDAMENTO",
    "DATACONCLUSAO", "OBSERVACAO", "STATUS", "EXECUTADOPOR", "TIPOCABO",
    "METRAGEM", "OBSERVACAOCONCLUSAO", "NUMDRAFT", "ROTA", "USUARIO",
]

# Campos de data que precisam de tratamento especial (ISO yyyy-mm-dd).
DATE_FIELDS = {"DATADISPARO", "DATAAGENDAMENTO", "DATACONCLUSAO"}

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


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
    """Converte valores variados (str, datetime, Timestamp do pandas) em date ISO ou None."""
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, str):
        # Já pode vir como 'YYYY-MM-DD' ou 'YYYY-MM-DDTHH:MM:SSZ'
        return value[:10]
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d")
    return str(value)[:10]


class DataClient:
    """Interface unificada de dados: SharePoint/Graph OU Excel local."""

    def __init__(self):
        self.use_excel = os.getenv("USE_EXCEL_FALLBACK", "true").strip().lower() == "true"
        self.excel_path = os.getenv("EXCEL_PATH", "")
        self.excel_sheet = os.getenv("EXCEL_SHEET_NAME", "TbRelatorio")

        # Config Graph (só é usada quando use_excel=False)
        self.tenant_id = os.getenv("TENANT_ID")
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.site_id = os.getenv("SHAREPOINT_SITE_ID")
        self.list_id = os.getenv("SHAREPOINT_LIST_ID")

        self._token_cache = None
        self._msal_app = None

        if not self.use_excel:
            self._init_msal_app()

    # ------------------------------------------------------------------
    # Autenticação (modo Graph)
    # ------------------------------------------------------------------
    def _init_msal_app(self):
        authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self._msal_app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=authority,
        )

    def authenticate(self):
        """Obtém (ou reaproveita do cache) um token de aplicativo via MSAL.

        Retorna o access_token (string). Lança DataClientError em caso de falha
        de autenticação/permissão, com mensagem apropriada para o frontend.
        """
        if self.use_excel:
            raise DataClientError("authenticate() chamado em modo Excel — não aplicável.")

        scopes = ["https://graph.microsoft.com/.default"]
        result = self._msal_app.acquire_token_silent(scopes, account=None)
        if not result:
            result = self._msal_app.acquire_token_for_client(scopes=scopes)

        if "access_token" not in result:
            logger.error("Falha na autenticação MSAL: %s", result.get("error_description"))
            raise DataClientError(
                "Falha ao autenticar no Azure AD. Verifique TENANT_ID, CLIENT_ID, "
                "CLIENT_SECRET e se o consentimento de administrador foi concedido "
                "para as permissões da aplicação (ex.: Sites.ReadWrite.All).",
                status_code=401,
            )
        return result["access_token"]

    def _graph_headers(self):
        token = self.authenticate()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    # ------------------------------------------------------------------
    # Métodos públicos (delegam para Excel ou Graph conforme configuração)
    # ------------------------------------------------------------------
    def list_items(self, filters=None, page=1, page_size=20, sort=None):
        """Lista itens com filtros, paginação e ordenação.

        filters: dict com chaves opcionais: cliente, id, cidade,
                 executadopor, status, data_inicio, data_fim
        sort: string no formato "campo:asc" ou "campo:desc"
        Retorna: dict {items: [...], total: int, page: int, page_size: int}
        """
        filters = filters or {}
        if self.use_excel:
            rows = self._excel_read_all()
        else:
            rows = self._graph_list_all()

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
        if self.use_excel:
            rows = self._excel_read_all()
        else:
            rows = self._graph_list_all()
        for row in rows:
            if str(row.get("IDCLIENTE")) == str(item_id):
                return row
        raise DataClientError(f"Registro com IDCLIENTE={item_id} não encontrado.", status_code=404)

    def create_item(self, data):
        data = self._sanitize_and_validate(data, is_new=True)
        if self.use_excel:
            return self._excel_create(data)
        return self._graph_create(data)

    def update_item(self, item_id, data):
        data = self._sanitize_and_validate(data, is_new=False)
        if self.use_excel:
            return self._excel_update(item_id, data)
        return self._graph_update(item_id, data)

    def delete_item(self, item_id):
        if self.use_excel:
            return self._excel_delete(item_id)
        return self._graph_delete(item_id)

    def dashboard_aggregates(self, filters=None):
        """Retorna agregações prontas para Chart.js + KPIs, respeitando filtros opcionais."""
        filters = filters or {}
        if self.use_excel:
            rows = self._excel_read_all()
        else:
            rows = self._graph_list_all()
        rows = self._apply_filters(rows, filters)

        def group_count(field):
            counts = {}
            for row in rows:
                key = row.get(field) or "Não informado"
                counts[key] = counts.get(key, 0) + 1
            # Ordena por contagem desc e limita a 12 categorias para legibilidade
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

        return {
            "kpis": kpis,
            "por_cliente": group_count("CLIENTE"),
            "por_cidade": group_count("CIDADE"),
            "por_executadopor": group_count("EXECUTADOPOR"),
            "por_status": status_counts,
            "por_data_agendamento": self._group_by_date(rows, "DATAAGENDAMENTO"),
        }

    def items_by_date(self, date_str):
        """Retorna todos os itens cuja DATAAGENDAMENTO == date_str (YYYY-MM-DD),
        já projetados apenas com as colunas usadas no modal de detalhe por data."""
        if self.use_excel:
            rows = self._excel_read_all()
        else:
            rows = self._graph_list_all()
        cols = ["IDCLIENTE", "CLIENTE", "ENDERECO", "CIDADE", "TECNOLOGIA", "VT",
                "DATAAGENDAMENTO", "TIPOCABO", "METRAGEM", "NUMDRAFT", "ROTA"]
        result = []
        for row in rows:
            if _parse_date(row.get("DATAAGENDAMENTO")) == date_str:
                result.append({c: row.get(c) for c in cols})
        return result

    # ------------------------------------------------------------------
    # Filtros / ordenação (compartilhados entre os dois modos)
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
    # Implementação: Excel local (pandas/openpyxl)
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

        # Garante que todas as colunas esperadas existam, mesmo que vazias.
        for field in FIELDS:
            if field not in df.columns:
                df[field] = ""
        df = df.fillna("")
        return df.to_dict(orient="records")

    def _excel_write_all(self, rows):
        df = pd.DataFrame(rows, columns=FIELDS)
        try:
            df.to_excel(self.excel_path, sheet_name=self.excel_sheet, index=False)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Erro escrevendo Excel")
            raise DataClientError(f"Erro ao gravar o arquivo Excel: {exc}", status_code=500)

    def _excel_create(self, data):
        rows = self._excel_read_all()
        if any(str(r.get("IDCLIENTE")) == str(data.get("IDCLIENTE")) for r in rows):
            raise DataClientError(f"Já existe um registro com IDCLIENTE={data.get('IDCLIENTE')}.", status_code=409)
        new_row = {f: data.get(f, "") for f in FIELDS}
        rows.append(new_row)
        self._excel_write_all(rows)
        logger.info("Registro criado (Excel): IDCLIENTE=%s", data.get("IDCLIENTE"))
        return new_row

    def _excel_update(self, item_id, data):
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
        logger.info("Registro atualizado (Excel): IDCLIENTE=%s", item_id)
        return self.get_item(item_id)

    def _excel_delete(self, item_id):
        rows = self._excel_read_all()
        new_rows = [r for r in rows if str(r.get("IDCLIENTE")) != str(item_id)]
        if len(new_rows) == len(rows):
            raise DataClientError(f"Registro com IDCLIENTE={item_id} não encontrado.", status_code=404)
        self._excel_write_all(new_rows)
        logger.info("Registro excluído (Excel): IDCLIENTE=%s", item_id)
        return {"deleted": item_id}

    # ------------------------------------------------------------------
    # Implementação: Microsoft Graph / SharePoint List
    # ------------------------------------------------------------------
    def _graph_list_all(self):
        """Lista TODOS os itens da lista (com fieldValueSet expandido), paginando
        internamente pelo @odata.nextLink do Graph até esgotar os resultados."""
        headers = self._graph_headers()
        url = (
            f"{GRAPH_BASE_URL}/sites/{self.site_id}/lists/{self.list_id}/items"
            f"?expand=fields&$top=200"
        )
        rows = []
        while url:
            resp = requests.get(url, headers=headers, timeout=30)
            self._raise_for_graph_error(resp)
            payload = resp.json()
            for item in payload.get("value", []):
                fields = item.get("fields", {})
                row = {f: fields.get(f, "") for f in FIELDS}
                row["_graph_item_id"] = item.get("id")  # ID interno do Graph, usado internamente
                rows.append(row)
            url = payload.get("@odata.nextLink")
        return rows

    def _graph_create(self, data):
        headers = self._graph_headers()
        url = f"{GRAPH_BASE_URL}/sites/{self.site_id}/lists/{self.list_id}/items"
        body = {"fields": data}
        resp = requests.post(url, headers=headers, json=body, timeout=30)
        self._raise_for_graph_error(resp)
        logger.info("Registro criado (Graph): IDCLIENTE=%s", data.get("IDCLIENTE"))
        return data

    def _graph_update(self, item_id, data):
        graph_id = self._resolve_graph_item_id(item_id)
        headers = self._graph_headers()
        url = f"{GRAPH_BASE_URL}/sites/{self.site_id}/lists/{self.list_id}/items/{graph_id}/fields"
        resp = requests.patch(url, headers=headers, json=data, timeout=30)
        self._raise_for_graph_error(resp)
        logger.info("Registro atualizado (Graph): IDCLIENTE=%s", item_id)
        return self.get_item(item_id)

    def _graph_delete(self, item_id):
        graph_id = self._resolve_graph_item_id(item_id)
        headers = self._graph_headers()
        url = f"{GRAPH_BASE_URL}/sites/{self.site_id}/lists/{self.list_id}/items/{graph_id}"
        resp = requests.delete(url, headers=headers, timeout=30)
        self._raise_for_graph_error(resp)
        logger.info("Registro excluído (Graph): IDCLIENTE=%s", item_id)
        return {"deleted": item_id}

    def _resolve_graph_item_id(self, idcliente):
        """O Graph identifica itens pelo seu próprio 'id' interno da lista, que é
        diferente de IDCLIENTE (chave de negócio). Esta função faz a ponte entre
        os dois, buscando o item pelo IDCLIENTE e retornando o id interno."""
        rows = self._graph_list_all()
        for row in rows:
            if str(row.get("IDCLIENTE")) == str(idcliente):
                return row["_graph_item_id"]
        raise DataClientError(f"Registro com IDCLIENTE={idcliente} não encontrado no SharePoint.", status_code=404)

    @staticmethod
    def _raise_for_graph_error(resp):
        if resp.status_code >= 400:
            try:
                payload = resp.json()
                message = payload.get("error", {}).get("message", resp.text)
            except ValueError:
                message = resp.text
            logger.error("Erro Graph [%s]: %s", resp.status_code, message)
            if resp.status_code in (401, 403):
                raise DataClientError(
                    "Acesso negado pelo Microsoft Graph. Verifique se a aplicação tem a "
                    "permissão Sites.ReadWrite.All (Application permission) e se o "
                    "consentimento de administrador foi concedido no Azure AD.",
                    status_code=resp.status_code,
                )
            raise DataClientError(f"Erro na chamada ao Microsoft Graph: {message}", status_code=resp.status_code)
