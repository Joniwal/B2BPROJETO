/* ==========================================================================
   REDEB2B — main.js
   Lógica do frontend: busca de dados via fetch/async-await, renderização de
   tabela, paginação, ordenação, filtros, dashboard (Chart.js) e modais.
   ========================================================================== */

const state = {
  page: 1,
  pageSize: 20,
  sort: "DATAAGENDAMENTO:desc", // mais recente primeiro, por padrão
  filters: {},
  charts: {},
  deleteTargetId: null,
};

const STATUS_OPTIONS = ["Novo", "Pendente", "Agendado", "Em Execução", "Concluído", "Cancelado"];

// Nomes disponíveis no combobox "Quem está registrando?". A data e hora são
// adicionadas automaticamente ao valor salvo no campo USUARIO — não edite
// isso manualmente, só a lista de nomes abaixo.
const USUARIO_OPTIONS = ["Cintia", "Maria Cristina", "Isabella", "Joni Wilson", "Marcos Neves", "Ana Vitória"];

// --------------------------------------------------------------------------
// Estas 4 listas ficam vazias por padrão — adicione os valores que você
// quiser disponibilizar em cada combobox, um por linha, como no exemplo
// comentado abaixo. A ordem da lista é a ordem que aparece no dropdown.
// Um valor já existente num registro antigo que não esteja na lista continua
// aparecendo normalmente ao editar aquele registro (é adicionado ao combobox
// automaticamente), só não fica disponível para novos registros até você
// incluí-lo aqui.
// Exemplo: const CIDADE_OPTIONS = ["Campinas", "Sorocaba", "São Paulo"];
// --------------------------------------------------------------------------
const CIDADE_OPTIONS = [];
const ATIVIDADE_OPTIONS = [];
const TECNOLOGIA_OPTIONS = [];
const EXECUTADOPOR_OPTIONS = [];

// Campos do formulário que são comboboxes com fallback (aceitam um valor
// antigo fora da lista, sem perder o dado ao editar).
const SELECT_FIELDS_WITH_FALLBACK = ["CIDADE", "ATIVIDADE", "TECNOLOGIA", "EXECUTADOPOR"];

const FORM_FIELDS = [
  "IDCLIENTE", "CLIENTE", "ENDERECO", "CIDADE", "PRODUTO", "ATIVIDADE",
  "TECNOLOGIA", "VT", "DATADISPARO", "RETORNOPCC", "DATAAGENDAMENTO",
  "DATACONCLUSAO", "OBSERVACAO", "STATUS", "EXECUTADOPOR", "TIPOCABO",
  "METRAGEM", "OBSERVACAOCONCLUSAO", "NUMDRAFT", "ROTA", "USUARIO",
];

document.addEventListener("DOMContentLoaded", () => {
  populateStatusSelects();
  populateFixedFormSelects();
  bindEvents();
  loadDashboard();
  loadItems();
});

/* -------------------------------------------------------------------- */
/* Setup                                                                 */
/* -------------------------------------------------------------------- */
function populateStatusSelects() {
  const quick = document.getElementById("quickStatusFilter");
  const filterSelect = document.getElementById("fStatus");
  const formSelect = document.getElementById("f_STATUS");
  STATUS_OPTIONS.forEach((status) => {
    quick.appendChild(new Option(status, status));
    filterSelect.appendChild(new Option(status, status));
    formSelect.appendChild(new Option(status, status));
  });
}

function populateSelect(selectId, options, placeholder) {
  const select = document.getElementById(selectId);
  select.innerHTML = "";
  select.appendChild(new Option(placeholder, ""));
  options.forEach((opt) => select.appendChild(new Option(opt, opt)));
}

function populateFixedFormSelects() {
  populateSelect("f_USUARIO", USUARIO_OPTIONS, "Selecione...");
  populateSelect("f_CIDADE", CIDADE_OPTIONS, "Selecione a cidade...");
  populateSelect("f_ATIVIDADE", ATIVIDADE_OPTIONS, "Selecione a atividade...");
  populateSelect("f_TECNOLOGIA", TECNOLOGIA_OPTIONS, "Selecione a tecnologia...");
  populateSelect("f_EXECUTADOPOR", EXECUTADOPOR_OPTIONS, "Selecione o executor...");
}

/** Define o valor de um combobox com fallback: se o valor não existir entre
 * as opções (ex.: dado antigo digitado livremente antes deste combobox
 * existir), adiciona uma opção extra marcada para não perder a informação. */
function setSelectValueWithFallback(selectId, value) {
  const select = document.getElementById(selectId);
  if (!value) {
    select.value = "";
    return;
  }
  const exists = Array.from(select.options).some((o) => o.value === value);
  if (!exists) {
    const opt = new Option(value, value);
    opt.dataset.fallback = "true";
    select.appendChild(opt);
  }
  select.value = value;
}

/** Remove as opções de fallback adicionadas temporariamente, para não
 * acumular duplicatas entre uma edição e outra. */
function clearSelectFallbackOptions(selectId) {
  const select = document.getElementById(selectId);
  select.querySelectorAll('option[data-fallback="true"]').forEach((o) => o.remove());
}

function bindEvents() {
  document.getElementById("btnAplicarFiltros").addEventListener("click", () => {
    collectFilters();
    state.page = 1;
    loadItems();
    loadDashboard();
  });

  document.getElementById("btnLimparFiltros").addEventListener("click", () => {
    ["fCliente", "fId", "fCidade", "fExecutadoPor", "fStatus", "fDataInicio", "fDataFim"].forEach((id) => {
      document.getElementById(id).value = "";
    });
    document.getElementById("quickSearch").value = "";
    document.getElementById("quickStatusFilter").value = "";
    state.filters = {};
    state.page = 1;
    loadItems();
    loadDashboard();
  });

  document.getElementById("quickStatusFilter").addEventListener("change", (e) => {
    document.getElementById("fStatus").value = e.target.value;
    collectFilters();
    state.page = 1;
    loadItems();
    loadDashboard();
  });

  let searchDebounce;
  document.getElementById("quickSearch").addEventListener("input", (e) => {
    clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => {
      state.filters.q = e.target.value || undefined;
      state.page = 1;
      loadItems();
    }, 400);
  });

  document.getElementById("pageSizeSelect").addEventListener("change", (e) => {
    state.pageSize = parseInt(e.target.value, 10);
    state.page = 1;
    loadItems();
  });

  document.querySelectorAll("#itemsTable thead th[data-sort]").forEach((th) => {
    th.addEventListener("click", () => {
      const field = th.dataset.sort;
      if (state.sort && state.sort.startsWith(field)) {
        const dir = state.sort.endsWith("asc") ? "desc" : "asc";
        state.sort = `${field}:${dir}`;
      } else {
        state.sort = `${field}:asc`;
      }
      loadItems();
    });
  });

  document.getElementById("btnNovoRegistro").addEventListener("click", () => openCreateModal());
  document.getElementById("btnSalvarItem").addEventListener("click", () => submitItemForm());
  document.getElementById("btnConfirmDelete").addEventListener("click", () => confirmDelete());

  document.getElementById("btnAtualizar").addEventListener("click", () => {
    loadItems();
    loadDashboard();
    showAlert("Dados atualizados.", "success");
  });

  document.getElementById("btnVerAgendamentosData").addEventListener("click", () => {
    const dateStr = document.getElementById("fDataInicio").value;
    if (!dateStr) {
      showAlert('Selecione uma data em "Data agendamento — de" antes de visualizar.', "warning");
      return;
    }
    openDateItemsModal(dateStr);
  });
}

function collectFilters() {
  state.filters = {
    cliente: document.getElementById("fCliente").value || undefined,
    id: document.getElementById("fId").value || undefined,
    cidade: document.getElementById("fCidade").value || undefined,
    executadopor: document.getElementById("fExecutadoPor").value || undefined,
    status: document.getElementById("fStatus").value || undefined,
    data_inicio: document.getElementById("fDataInicio").value || undefined,
    data_fim: document.getElementById("fDataFim").value || undefined,
    q: document.getElementById("quickSearch").value || undefined,
  };
}

/* -------------------------------------------------------------------- */
/* Helpers de rede                                                       */
/* -------------------------------------------------------------------- */
function buildQueryParams(extra = {}) {
  const params = new URLSearchParams();
  const map = {
    cliente: state.filters.cliente,
    id: state.filters.id,
    cidade: state.filters.cidade,
    executadopor: state.filters.executadopor,
    status: state.filters.status,
    dataInicio: state.filters.data_inicio,
    dataFim: state.filters.data_fim,
    q: state.filters.q,
    ...extra,
  };
  Object.entries(map).forEach(([k, v]) => {
    if (v !== undefined && v !== null && v !== "") params.append(k, v);
  });
  return params;
}

async function apiFetch(url, options = {}) {
  try {
    const resp = await fetch(url, {
      headers: { "Content-Type": "application/json" },
      ...options,
    });
    const payload = await resp.json();
    if (!resp.ok || !payload.ok) {
      throw new Error(payload.error || `Erro HTTP ${resp.status}`);
    }
    return payload.data;
  } catch (err) {
    showAlert(err.message || "Erro de comunicação com o servidor.", "danger");
    throw err;
  }
}

function showAlert(message, type = "success") {
  const container = document.getElementById("alertContainer");
  const el = document.createElement("div");
  el.className = `alert alert-${type} alert-dismissible fade show`;
  el.innerHTML = `${escapeHtml(message)}<button type="button" class="btn-close" data-bs-dismiss="alert"></button>`;
  container.appendChild(el);
  setTimeout(() => el.remove(), 5000);
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str ?? "";
  return div.innerHTML;
}

function formatDateBR(dateStr) {
  if (!dateStr) return "";
  const [year, month, day] = String(dateStr).slice(0, 10).split("-");
  if (!year || !month || !day) return dateStr;
  return `${day}/${month}/${year}`;
}

function formatDateDiaMes(dateStr) {
  if (!dateStr) return "";
  const [, month, day] = String(dateStr).slice(0, 10).split("-");
  if (!month || !day) return dateStr;
  return `${day}/${month}`;
}

/* -------------------------------------------------------------------- */
/* Tabela / paginação / ordenação                                        */
/* -------------------------------------------------------------------- */
async function loadItems() {
  const params = buildQueryParams({
    page: state.page,
    page_size: state.pageSize,
    sort: state.sort || undefined,
  });
  try {
    const data = await apiFetch(`/api/items?${params.toString()}`);
    renderTable(data.items);
    renderPagination(data.total, data.page, data.page_size);
  } catch (err) {
    document.getElementById("itemsTableBody").innerHTML =
      `<tr><td colspan="7" class="text-center text-danger py-4">Erro ao carregar registros.</td></tr>`;
  }
}

function statusClass(status) {
  const key = (status || "").toLowerCase().replace(/\s+/g, "-").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const known = ["novo", "pendente", "agendado", "em-execucao", "concluido", "cancelado"];
  return known.includes(key) ? `status-${key}` : "status-outro";
}

function renderTable(items) {
  const tbody = document.getElementById("itemsTableBody");
  if (!items || items.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">Nenhum registro encontrado.</td></tr>`;
    return;
  }
  tbody.innerHTML = items.map((item) => `
    <tr>
      <td>${escapeHtml(item.IDCLIENTE)}</td>
      <td>${escapeHtml(item.CLIENTE)}</td>
      <td>${escapeHtml(item.CIDADE)}</td>
      <td>${escapeHtml(item.EXECUTADOPOR)}</td>
      <td><span class="status-badge ${statusClass(item.STATUS)}">${escapeHtml(item.STATUS || "—")}</span></td>
      <td>${formatDateBR(item.DATAAGENDAMENTO)}</td>
      <td class="text-end">
        <i class="bi bi-eye action-icon" title="Ver / editar" data-action="view" data-id="${escapeHtml(item.IDCLIENTE)}"></i>
        <i class="bi bi-trash action-icon text-danger" title="Excluir" data-action="delete" data-id="${escapeHtml(item.IDCLIENTE)}"></i>
      </td>
    </tr>
  `).join("");

  tbody.querySelectorAll('[data-action="view"]').forEach((el) => {
    el.addEventListener("click", () => openEditModal(el.dataset.id));
  });
  tbody.querySelectorAll('[data-action="delete"]').forEach((el) => {
    el.addEventListener("click", () => openDeleteModal(el.dataset.id));
  });
}

function renderPagination(total, page, pageSize) {
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  const info = document.getElementById("paginationInfo");
  const start = total === 0 ? 0 : (page - 1) * pageSize + 1;
  const end = Math.min(page * pageSize, total);
  info.textContent = `Mostrando ${start}–${end} de ${total} registros`;

  const controls = document.getElementById("paginationControls");
  const pages = [];
  const addPage = (p, label, disabled = false, active = false) => {
    pages.push(`
      <li class="page-item ${disabled ? "disabled" : ""} ${active ? "active" : ""}">
        <a class="page-link" href="#" data-page="${p}">${label}</a>
      </li>
    `);
  };

  addPage(page - 1, "«", page <= 1);
  const windowStart = Math.max(1, page - 2);
  const windowEnd = Math.min(totalPages, page + 2);
  for (let p = windowStart; p <= windowEnd; p++) {
    addPage(p, p, false, p === page);
  }
  addPage(page + 1, "»", page >= totalPages);

  controls.innerHTML = pages.join("");
  controls.querySelectorAll("a[data-page]").forEach((a) => {
    a.addEventListener("click", (e) => {
      e.preventDefault();
      const p = parseInt(a.dataset.page, 10);
      if (p >= 1 && p <= totalPages && p !== state.page) {
        state.page = p;
        loadItems();
      }
    });
  });
}

/* -------------------------------------------------------------------- */
/* Dashboard / KPIs / Charts                                             */
/* -------------------------------------------------------------------- */
async function loadDashboard() {
  const params = buildQueryParams();
  try {
    const data = await apiFetch(`/api/dashboard?${params.toString()}`);
    renderKpis(data.kpis);
    try {
      renderCharts(data);
    } catch (chartErr) {
      console.error("Erro ao renderizar gráficos:", chartErr);
      showAlert("Não foi possível renderizar os gráficos (verifique o console do navegador).", "warning");
    }
  } catch (err) {
    // erro já exibido via apiFetch
  }
}

function renderKpis(kpis) {
  const row = document.getElementById("kpiRow");
  const normalize = (s) => (s || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  const statusEntries = Object.entries(kpis.por_status || {}).filter(([status]) => {
    const n = normalize(status);
    return n !== "concluido" && n !== "novo";
  });

  const cards = [
    { label: "Total de Registros", value: kpis.total ?? 0, icon: "bi-collection", color: "kpi-grey", statusFilter: "" },
    ...statusEntries.slice(0, 1).map(([status, count]) => ({
      label: status,
      value: count,
      icon: "bi-flag",
      color: "kpi-orange",
      statusFilter: status,
    })),
    {
      label: "Novos",
      value: kpis.novos_total ?? 0,
      icon: "bi-stars",
      color: "kpi-blue",
      statusFilter: "Novo",
    },
    {
      label: kpis.concluidos_label || "Concluídos",
      value: kpis.concluidos_total ?? 0,
      icon: "bi-check-circle",
      color: "kpi-green",
      statusFilter: "Concluído",
    },
  ];

  row.innerHTML = cards.map((c) => `
    <div class="col-lg-3 col-md-6">
      <div class="kpi-card kpi-clickable ${c.color}" data-status-filter="${escapeHtml(c.statusFilter)}" title="Clique para filtrar por este status">
        <i class="bi ${c.icon} kpi-icon"></i>
        <div>
          <div class="kpi-label">${escapeHtml(c.label)}</div>
          <div class="kpi-value">${c.value}</div>
        </div>
      </div>
    </div>
  `).join("");

  row.querySelectorAll(".kpi-card").forEach((el) => {
    el.addEventListener("click", () => {
      const statusFilter = el.dataset.statusFilter || "";
      document.getElementById("fStatus").value = statusFilter;
      document.getElementById("quickStatusFilter").value = statusFilter;
      collectFilters();
      state.page = 1;
      loadItems();
      loadDashboard();
    });
  });
}

function destroyChart(key) {
  if (state.charts[key]) {
    state.charts[key].destroy();
    delete state.charts[key];
  }
}

function renderCharts(data) {
  renderBarChart("chartCliente", data.por_cliente, "#3d95c4");
  renderBarChart("chartCidade", data.por_cidade, "#59a869");
  renderBarChart("chartExecutadoPor", data.por_executadopor, "#f0913e");
  renderLineChart("chartDataConclusao", data.por_data_conclusao);
}

function renderBarChart(canvasId, dataset, color) {
  destroyChart(canvasId);
  const ctx = document.getElementById(canvasId).getContext("2d");
  state.charts[canvasId] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: dataset.labels,
      datasets: [{ data: dataset.data, backgroundColor: color, borderRadius: 3 }],
    },
    options: {
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0, font: { size: 8 } } },
        x: { ticks: { font: { size: 8 }, maxRotation: 60, minRotation: 45 } },
      },
      responsive: true,
      maintainAspectRatio: false,
    },
  });
}

function renderLineChart(canvasId, dataset) {
  destroyChart(canvasId);
  const ctx = document.getElementById(canvasId).getContext("2d");
  state.charts[canvasId] = new Chart(ctx, {
    type: "line",
    data: {
      labels: dataset.labels,
      datasets: [{
        data: dataset.data,
        borderColor: "#d9483a",
        backgroundColor: "rgba(217,72,58,0.15)",
        tension: 0.3,
        fill: true,
        pointRadius: 3,
        pointHoverRadius: 5,
      }],
    },
    options: {
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: (items) => formatDateBR(dataset.labels[items[0].dataIndex]),
          },
        },
      },
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0, font: { size: 8 } } },
        x: {
          ticks: {
            font: { size: 8 },
            maxRotation: 60,
            minRotation: 45,
            callback: function (value, index) {
              return formatDateDiaMes(dataset.labels[index]);
            },
          },
        },
      },
      responsive: true,
      maintainAspectRatio: false,
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const idx = elements[0].index;
        const dateStr = dataset.labels[idx];
        if (dateStr) openDateItemsModal(dateStr, "DATACONCLUSAO", "Concluído");
      },
    },
  });
}

/* -------------------------------------------------------------------- */
/* Modal de itens agendados por data (clique no gráfico de linha)        */
/* -------------------------------------------------------------------- */
async function openDateItemsModal(dateStr, dateField = "DATAAGENDAMENTO", status = null) {
  const isConclusao = dateField === "DATACONCLUSAO";
  document.getElementById("dateItemsTitlePrefix").textContent = isConclusao
    ? "Clientes concluídos em"
    : "Agendamentos em";
  document.getElementById("dateItemsLabel").textContent = formatDateBR(dateStr);

  const tbody = document.getElementById("dateItemsBody");
  tbody.innerHTML = `<tr><td colspan="13" class="text-center text-muted py-3">Carregando...</td></tr>`;
  const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById("dateItemsModal"));
  modal.show();

  try {
    const params = new URLSearchParams({ data: dateStr, campo: dateField });
    if (status) params.append("status", status);
    const items = await apiFetch(`/api/items-by-date?${params.toString()}`);
    if (!items.length) {
      tbody.innerHTML = `<tr><td colspan="13" class="text-center text-muted py-3">Nenhum registro para esta data.</td></tr>`;
      return;
    }
    tbody.innerHTML = items.map((it) => `
      <tr>
        <td>${escapeHtml(it.IDCLIENTE)}</td>
        <td>${escapeHtml(it.CLIENTE)}</td>
        <td>${escapeHtml(it.ENDERECO)}</td>
        <td>${escapeHtml(it.CIDADE)}</td>
        <td>${escapeHtml(it.TECNOLOGIA)}</td>
        <td>${escapeHtml(it.VT)}</td>
        <td><span class="status-badge ${statusClass(it.STATUS)}">${escapeHtml(it.STATUS || "—")}</span></td>
        <td>${formatDateBR(it.DATAAGENDAMENTO)}</td>
        <td>${formatDateBR(it.DATACONCLUSAO)}</td>
        <td>${escapeHtml(it.TIPOCABO)}</td>
        <td>${escapeHtml(it.METRAGEM)}</td>
        <td>${escapeHtml(it.NUMDRAFT)}</td>
        <td>${escapeHtml(it.ROTA)}</td>
      </tr>
    `).join("");
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="13" class="text-center text-danger py-3">Erro ao carregar registros da data.</td></tr>`;
  }
}

/* -------------------------------------------------------------------- */
/* Modal de criação / edição                                             */
/* -------------------------------------------------------------------- */
function clearItemForm() {
  FORM_FIELDS.forEach((f) => {
    const el = document.getElementById(`f_${f}`);
    if (el) el.value = "";
  });
  SELECT_FIELDS_WITH_FALLBACK.forEach((f) => clearSelectFallbackOptions(`f_${f}`));
  document.getElementById("f_originalId").value = "";
  document.getElementById("lastUsuarioInfo").textContent = "";
  document.getElementById("itemModalAlert").innerHTML = "";
}

function openCreateModal() {
  clearItemForm();
  document.getElementById("itemModalTitle").innerHTML = `<i class="bi bi-plus-circle"></i> Novo registro`;
  document.getElementById("f_IDCLIENTE").disabled = false;
  const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById("itemModal"));
  modal.show();
  setTimeout(() => document.getElementById("f_IDCLIENTE").focus(), 300);
}

async function openEditModal(id) {
  clearItemForm();
  document.getElementById("itemModalTitle").innerHTML = `<i class="bi bi-pencil-square"></i> Editar registro ${escapeHtml(id)}`;
  const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById("itemModal"));
  modal.show();

  try {
    const item = await apiFetch(`/api/items/${encodeURIComponent(id)}`);
    FORM_FIELDS.forEach((f) => {
      if (f === "USUARIO") return; // tratado à parte: combobox sempre inicia em branco
      const el = document.getElementById(`f_${f}`);
      if (!el) return;
      if (SELECT_FIELDS_WITH_FALLBACK.includes(f)) {
        setSelectValueWithFallback(`f_${f}`, item[f] ?? "");
      } else {
        el.value = item[f] ?? "";
      }
    });
    document.getElementById("lastUsuarioInfo").textContent = item.USUARIO
      ? `Último registro por: ${item.USUARIO}`
      : "";
    document.getElementById("f_originalId").value = id;
    document.getElementById("f_IDCLIENTE").disabled = true; // chave de negócio não é editável
    setTimeout(() => document.getElementById("f_CLIENTE").focus(), 300);
  } catch (err) {
    modal.hide();
  }
}

async function submitItemForm() {
  const originalId = document.getElementById("f_originalId").value;
  const isEdit = Boolean(originalId);

  const form = document.getElementById("itemForm");
  if (!form.checkValidity()) {
    form.reportValidity();
    return;
  }

  const payload = {};
  FORM_FIELDS.forEach((f) => {
    const el = document.getElementById(`f_${f}`);
    if (el) payload[f] = el.value;
  });

  // Usuário + data/hora automáticos: o combobox só guarda o nome escolhido;
  // aqui montamos o valor final gravado no campo USUARIO.
  const usuarioSelecionado = document.getElementById("f_USUARIO").value;
  const agora = new Date();
  const dataHora = agora.toLocaleString("pt-BR", {
    day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit",
  });
  payload.USUARIO = `${usuarioSelecionado} - ${dataHora}`;

  try {
    if (isEdit) {
      await apiFetch(`/api/items/${encodeURIComponent(originalId)}`, {
        method: "PATCH",
        body: JSON.stringify(payload),
      });
      showAlert("Registro atualizado com sucesso.", "success");
    } else {
      await apiFetch(`/api/items`, {
        method: "POST",
        body: JSON.stringify(payload),
      });
      showAlert("Registro criado com sucesso.", "success");
    }
    bootstrap.Modal.getOrCreateInstance(document.getElementById("itemModal")).hide();
    loadItems();
    loadDashboard();
  } catch (err) {
    document.getElementById("itemModalAlert").innerHTML =
      `<div class="alert alert-danger py-2">${escapeHtml(err.message)}</div>`;
  }
}

/* -------------------------------------------------------------------- */
/* Exclusão                                                              */
/* -------------------------------------------------------------------- */
function openDeleteModal(id) {
  state.deleteTargetId = id;
  document.getElementById("deleteItemLabel").textContent = id;
  bootstrap.Modal.getOrCreateInstance(document.getElementById("deleteModal")).show();
}

async function confirmDelete() {
  if (!state.deleteTargetId) return;
  try {
    await apiFetch(`/api/items/${encodeURIComponent(state.deleteTargetId)}`, { method: "DELETE" });
    showAlert("Registro excluído com sucesso.", "success");
    bootstrap.Modal.getOrCreateInstance(document.getElementById("deleteModal")).hide();
    loadItems();
    loadDashboard();
  } catch (err) {
    // erro já exibido via apiFetch
  } finally {
    state.deleteTargetId = null;
  }
}
