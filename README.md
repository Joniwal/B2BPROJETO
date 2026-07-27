# REDEB2B — Painel Web (Flask + Excel local)

Aplicação web para gerenciar os registros de agendamento de clientes e
atividades de rede (REDEB2B), com CRUD completo, busca, filtros, paginação,
ordenação e um dashboard analítico com gráficos. Os dados são lidos e
gravados diretamente em um arquivo Excel local (`.xlsx`), sem nenhuma
dependência de SharePoint, Azure AD ou Microsoft Graph.

---

## 1. Arquitetura

```
redeb2b_app/
├── app.py                  # Inicialização do Flask, registro de blueprint, error handlers
├── api.py                  # Blueprint com os endpoints REST (/api/...)
├── excel_client.py         # Camada de dados: leitura/escrita no Excel via pandas/openpyxl
├── gerar_dados_exemplo.py  # Script para popular um Excel de teste
├── requirements.txt
├── .env.example
├── templates/
│   ├── index.html          # Página única: topbar, filtros, KPIs, gráficos, tabela
│   └── modals.html         # Modais de criar/editar, excluir e "itens por data"
└── static/
    ├── css/styles.css
    └── js/main.js           # fetch/async-await, paginação, ordenação, Chart.js
```

**`excel_client.py`** expõe a classe `DataClient`, com os métodos
`list_items`, `get_item`, `create_item`, `update_item`, `delete_item`,
`dashboard_aggregates` e `items_by_date`. O arquivo Excel usado é decidido
assim:
1. Se `EXCEL_PATH` estiver preenchido no `.env`, esse caminho é usado
   diretamente (modo manual).
2. Se `EXCEL_PATH` estiver em branco, o sistema procura automaticamente um
   arquivo chamado `EXCEL_FILENAME` (padrão: `REDE_B2B.xlsx`) dentro das
   pastas do OneDrive sincronizadas nesta máquina (pessoal e/ou
   corporativo), inclusive dentro de subpastas.

A cada escrita (criar/atualizar/excluir), o arquivo inteiro é relido e
regravado (de forma atômica, num arquivo temporário substituído no final),
então evite editar o Excel manualmente enquanto a aplicação estiver com o
navegador aberto e em uso simultâneo.

**Endpoints da API** (`api.py`, todas as respostas em JSON):

| Método | Rota | Descrição |
|---|---|---|
| GET | `/api/items` | Lista paginada, com filtros e ordenação |
| GET | `/api/items/<id>` | Detalhe de um registro (por IDCLIENTE) |
| POST | `/api/items` | Cria um registro |
| PATCH | `/api/items/<id>` | Atualiza um registro |
| DELETE | `/api/items/<id>` | Exclui um registro |
| GET | `/api/dashboard` | Agregações para KPIs e gráficos (Chart.js) |
| GET | `/api/items-by-date?data=YYYY-MM-DD` | Registros agendados em uma data (usado ao clicar no gráfico de linha) |
| GET | `/api/export` | Baixa em `.xlsx` os registros que atendem aos mesmos filtros/ordenação de `/api/items`, sem paginação (todos os que passam pelo filtro) |
| GET | `/api/excel-status` | Diagnóstico: mostra se o Excel foi localizado e em qual caminho (útil para depurar problemas de OneDrive) |

---

## 2. Instalação e configuração

### 2.1 Pré-requisitos
- Python 3.10+
- pip

### 2.2 Passos

```bash
# 1. Entre na pasta do projeto
cd redeb2b_app

# 2. Crie e ative um ambiente virtual (recomendado)
python -m venv venv
source venv/bin/activate        # Linux/Mac
venv\Scripts\activate           # Windows

# 3. Instale as dependências
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

# 4. Copie o arquivo de exemplo de variáveis de ambiente
cp .env.example .env             # Linux/Mac
copy .env.example .env           # Windows

# 5. Edite o .env se necessário (veja abaixo)
```

No `.env`, você tem duas opções:

**Opção 1 — automática (recomendada):** deixe `EXCEL_PATH` em branco. O
sistema vai procurar sozinho o arquivo `REDE_B2B.xlsx` (ou o nome que você
colocar em `EXCEL_FILENAME`) dentro das pastas do OneDrive sincronizadas
nesta máquina — não precisa saber o caminho completo.
```
EXCEL_PATH=
EXCEL_FILENAME=REDE_B2B.xlsx
EXCEL_SHEET_NAME=REDEB2B
```

**Opção 2 — manual:** preencha `EXCEL_PATH` com o caminho completo. Quando
preenchido, ele tem prioridade sobre a busca automática.
```
EXCEL_PATH=C:\Users\PEGGY\OneDrive - Telefonica\PROJETO\REDE_B2B.xlsx
EXCEL_SHEET_NAME=REDEB2B
```

Para conferir se o arquivo foi encontrado (e onde), acesse no navegador,
com a aplicação rodando:
```
http://localhost:5000/api/excel-status
```

Se você ainda não tem o arquivo, gere um de exemplo com dados fictícios:
```bash
python gerar_dados_exemplo.py
# ou, para um caminho específico:
python gerar_dados_exemplo.py "C:\Users\PEGGY\Desktop\PROJETO\REDE_B2B.xlsx"
```

### 2.3 Executando a aplicação

```bash
python app.py
```
Acesse **http://localhost:5000** no navegador.

---

## 3. Segurança e observações

- Como não há mais integração externa (Graph/Azure AD), não existe segredo
  algum a proteger — apenas o caminho do arquivo Excel, que fica só no
  backend (nunca é exposto ao navegador).
- Todos os inputs recebidos pela API são sanitizados (strip de espaços) e
  validados (ex.: datas devem estar em `yyyy-mm-dd`) antes de serem
  gravados no Excel.
- Evite ter o arquivo Excel aberto no próprio Excel/LibreOffice enquanto a
  aplicação estiver rodando — alguns sistemas bloqueiam o arquivo para
  escrita externa enquanto ele está aberto, o que pode gerar erro ao
  salvar. Feche o arquivo no Excel antes de criar/editar/excluir pelo
  painel.
- **Backup:** como a aplicação regrava o arquivo inteiro a cada alteração,
  é recomendável manter uma rotina simples de backup do `.xlsx` (cópia
  periódica), especialmente antes de operações em lote.

---

## 4. Exemplos de uso da API (cURL)

### Listar itens com filtros, paginação e ordenação
```bash
curl "http://localhost:5000/api/items?cliente=Empresa&status=Agendado&page=1&page_size=20&sort=DATAAGENDAMENTO:desc"
```

### Buscar um item pelo IDCLIENTE
```bash
curl "http://localhost:5000/api/items/CLI1001"
```

### Criar um item
```bash
curl -X POST "http://localhost:5000/api/items" \
  -H "Content-Type: application/json" \
  -d '{
    "IDCLIENTE": "CLI2001",
    "CLIENTE": "Nova Empresa LTDA",
    "ENDERECO": "Av. Central, 500",
    "CIDADE": "Campinas",
    "PRODUTO": "Internet Dedicada",
    "ATIVIDADE": "Instalação",
    "TECNOLOGIA": "FTTH",
    "VT": "VT9001",
    "DATAAGENDAMENTO": "2026-08-05",
    "STATUS": "Agendado",
    "EXECUTADOPOR": "João Silva",
    "TIPOCABO": "Drop 1FO",
    "METRAGEM": "120",
    "NUMDRAFT": "DR00099",
    "ROTA": "Rota-2",
    "USUARIO": "admin"
  }'
```

### Atualizar um item (PATCH)
```bash
curl -X PATCH "http://localhost:5000/api/items/CLI2001" \
  -H "Content-Type: application/json" \
  -d '{"STATUS": "Concluído", "DATACONCLUSAO": "2026-08-06"}'
```

### Excluir um item
```bash
curl -X DELETE "http://localhost:5000/api/items/CLI2001"
```

### Dashboard (agregações para os gráficos)
```bash
curl "http://localhost:5000/api/dashboard?cidade=Campinas"
```

### Itens agendados em uma data específica
```bash
curl "http://localhost:5000/api/items-by-date?data=2026-08-05"
```

---

## 5. Testes rápidos

```bash
# 1. Gere dados de exemplo
python gerar_dados_exemplo.py

# 2. Rode a aplicação
python app.py

# 3. Abra http://localhost:5000 e confirme que:
#    - os KPIs e os 4 gráficos aparecem preenchidos
#    - a tabela lista os registros com paginação e permite ordenar por coluna
#      (por padrão, ordenado por Data de Agendamento, mais recente primeiro)
#    - a busca rápida e os filtros do painel esquerdo filtram a tabela e os gráficos
#    - clicar no ícone de "olho" abre o modal de edição com os dados carregados
#    - "Novo registro" abre o modal em branco e salva um registro novo (POST)
#    - o ícone de lixeira pede confirmação e exclui o registro (DELETE)
#    - clicar em um ponto do gráfico "Por Data de Agendamento" abre o modal
#      com a lista de registros daquela data
```
