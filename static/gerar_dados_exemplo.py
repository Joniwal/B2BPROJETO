# -*- coding: utf-8 -*-
"""
gerar_dados_exemplo.py
=======================
Script utilitário para gerar um arquivo Excel de exemplo (REDE_B2B.xlsx) com
dados fictícios, compatível com a estrutura de colunas usada pela aplicação
REDEB2B. Útil para testar a aplicação rapidamente sem precisar preencher o Excel
manualmente.

Uso:
    python gerar_dados_exemplo.py [caminho_opcional_do_arquivo.xlsx]

Se nenhum caminho for informado, usa o valor de EXCEL_PATH do .env (ou
"./REDE_B2B_exemplo.xlsx" como padrão).
"""

import os
import sys
import random
from datetime import date, timedelta

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

FIELDS = [
    "IDCLIENTE", "CLIENTE", "ENDERECO", "CIDADE", "PRODUTO", "ATIVIDADE",
    "TECNOLOGIA", "VT", "DATADISPARO", "RETORNOPCC", "DATAAGENDAMENTO",
    "DATACONCLUSAO", "OBSERVACAO", "STATUS", "EXECUTADOPOR", "TIPOCABO",
    "METRAGEM", "OBSERVACAOCONCLUSAO", "NUMDRAFT", "ROTA", "USUARIO",
]

CIDADES = ["São Paulo", "Campinas", "Sorocaba", "Ribeirão Preto", "Santos"]
STATUS = [
    "Novo", "CONCLUIDO", "PENDENTE AGENDAMENTO", "PCC", "SEM ACAO OSP",
    "CANCELADO", "VISTORIA", "INICIADO NAO FINALIZADO", "EM EXECUÇÃO",
]
TECNOLOGIAS = ["FTTH", "FTTC", "Rádio", "Metálico"]
EXECUTORES = ["João Silva", "Maria Souza", "Carlos Lima", "Ana Pereira"]
TIPOS_CABO = ["Drop 1FO", "Drop 2FO", "Backbone 12FO"]


def gerar_linha(i):
    hoje = date.today()
    data_agendamento = hoje + timedelta(days=random.randint(-10, 20))
    return {
        "IDCLIENTE": f"CLI{1000 + i}",
        "CLIENTE": f"Empresa Exemplo {i}",
        "ENDERECO": f"Rua das Flores, {100 + i}",
        "CIDADE": random.choice(CIDADES),
        "PRODUTO": "Internet Dedicada",
        "ATIVIDADE": "Instalação",
        "TECNOLOGIA": random.choice(TECNOLOGIAS),
        "VT": f"VT{i:04d}",
        "DATADISPARO": (hoje - timedelta(days=random.randint(1, 30))).isoformat(),
        "RETORNOPCC": random.choice(["OK", "Pendente"]),
        "DATAAGENDAMENTO": data_agendamento.isoformat(),
        "DATACONCLUSAO": "" if random.random() < 0.5 else data_agendamento.isoformat(),
        "OBSERVACAO": "Registro gerado automaticamente para testes.",
        "STATUS": random.choice(STATUS),
        "EXECUTADOPOR": random.choice(EXECUTORES),
        "TIPOCABO": random.choice(TIPOS_CABO),
        "METRAGEM": str(random.randint(50, 500)),
        "OBSERVACAOCONCLUSAO": "",
        "NUMDRAFT": f"DR{i:05d}",
        "ROTA": f"Rota-{random.randint(1, 5)}",
        "USUARIO": "admin",
    }


def main():
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        path = os.getenv("EXCEL_PATH", "./REDE_B2B_exemplo.xlsx")

    sheet_name = os.getenv("EXCEL_SHEET_NAME", "TbRelatorio")
    rows = [gerar_linha(i) for i in range(1, 41)]
    df = pd.DataFrame(rows, columns=FIELDS)

    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    df.to_excel(path, sheet_name=sheet_name, index=False)
    print(f"Arquivo de exemplo gerado em: {path} (aba '{sheet_name}') com {len(rows)} registros.")


if __name__ == "__main__":
    main()
