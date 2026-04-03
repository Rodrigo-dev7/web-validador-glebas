const state = {
  arquivoCarregado: null,
  errosCache: [],
  gruposCache: {},
};

const COLUMN_NAMES = {
  gleba: ["gleba", "num_gleba", "nr_gleba", "sequencial_gleba", "gleba_seq", "sq_glb"],
  ponto: ["ponto", "seq_ponto", "ordem_ponto", "nr_ponto", "sequencial_ponto", "sq_cgl"],
  latitude: ["latitude", "lat", "nr_lat"],
  longitude: ["longitude", "long", "lon", "lng", "nr_lon"],
};

const ERROR_BADGES = {
  "POLIGONO NAO FECHADO": "badge-red",
  "PONTOS INSUFICIENTES": "badge-amber",
  "PONTO DUPLICADO EM EXCESSO": "badge-yellow",
  "COORDENADA INVALIDA": "badge-red",
};

const ERROR_ICONS = {
  "POLIGONO NAO FECHADO": "[FECHAMENTO]",
  "PONTOS INSUFICIENTES": "[VERTICES]",
  "PONTO DUPLICADO EM EXCESSO": "[DUPLICADO]",
  "COORDENADA INVALIDA": "[COORDENADA]",
};

const UI_ICONS = {
  success: "✓",
  danger: "!",
  warning: "▲",
  info: "i",
};

const TOLERANCIA = 1e-8;
let elements = {};

document.addEventListener("DOMContentLoaded", () => {
  mapElements();
  bindEvents();
  renderizarRelatorioInicial();
});

function mapElements() {
  elements = {
    dropZone: document.getElementById("drop-zone"),
    dropKicker: document.getElementById("drop-kicker"),
    dropTitle: document.getElementById("drop-title"),
    dropSub: document.getElementById("drop-sub"),
    reportOverviewList: document.getElementById("report-overview-list"),
    reportGroups: document.getElementById("report-groups"),
    reportTotalGlebas: document.getElementById("report-total-glebas"),
    reportTotalErros: document.getElementById("report-total-erros"),
    reportTotalOk: document.getElementById("report-total-ok"),
    reportTotalPendentes: document.getElementById("report-total-pendentes"),
    glebasList: document.getElementById("glebas-list"),
    status: document.getElementById("status-txt"),
    progressFill: document.getElementById("progress-fill"),
    btnBuscar: document.getElementById("btn-buscar"),
    btnValidar: document.getElementById("btn-validar"),
    btnExportar: document.getElementById("btn-exportar"),
    btnLimpar: document.getElementById("btn-limpar"),
    fileInput: document.getElementById("file-input"),
    fileInputSide: document.getElementById("file-input-side"),
    statGlebas: document.getElementById("stat-glebas"),
    statErros: document.getElementById("stat-erros"),
    statOk: document.getElementById("stat-ok"),
    tabButtons: [...document.querySelectorAll(".tab-btn")],
    tabContents: [...document.querySelectorAll(".tab-content")],
  };
}

function bindEvents() {
  elements.btnBuscar.addEventListener("click", () => elements.fileInputSide.click());
  elements.btnValidar.addEventListener("click", iniciarValidacao);
  elements.btnExportar.addEventListener("click", exportarRelatorio);
  elements.btnLimpar.addEventListener("click", limpar);
  elements.fileInput.addEventListener("change", (event) => handleFileInput(event.target.files[0]));
  elements.fileInputSide.addEventListener("change", (event) => handleFileInput(event.target.files[0]));
  elements.dropZone.addEventListener("click", () => elements.fileInput.click());
  elements.dropZone.addEventListener("keydown", onDropZoneKeydown);
  elements.dropZone.addEventListener("dragover", onDragOver);
  elements.dropZone.addEventListener("dragleave", onDragLeave);
  elements.dropZone.addEventListener("drop", onDrop);
  elements.tabButtons.forEach((button) => {
    button.addEventListener("click", () => showTab(button.dataset.tab));
    button.addEventListener("keydown", onTabKeydown);
  });
}

function onDragOver(event) {
  event.preventDefault();
  elements.dropZone.classList.add("is-hover");
}

function onDragLeave() {
  elements.dropZone.classList.remove("is-hover");
}

function onDrop(event) {
  event.preventDefault();
  elements.dropZone.classList.remove("is-hover");
  const [file] = event.dataTransfer.files;
  handleFileInput(file);
}

function handleFileInput(file) {
  if (!file) return;
  processarArquivo(file);
}

function processarArquivo(file) {
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (!["xls", "xlsx"].includes(ext)) {
    setStatus("Formato invalido. Use um arquivo .xls ou .xlsx.");
    setProgress("error");
    return;
  }

  state.arquivoCarregado = file;
  elements.dropZone.classList.add("is-loaded");
  elements.dropKicker.textContent = "Arquivo selecionado";
  elements.dropTitle.textContent = file.name;
  elements.dropSub.textContent = "Clique para trocar o arquivo antes de validar.";
  elements.btnValidar.disabled = false;
  elements.btnExportar.disabled = true;
  setProgress("reset");
  setStatus(`Arquivo pronto para validacao: ${file.name}`);
}

function onDropZoneKeydown(event) {
  if (event.key !== "Enter" && event.key !== " ") return;
  event.preventDefault();
  elements.fileInput.click();
}

function lerPlanilha(file, callback) {
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      callback(null, rows);
    } catch (error) {
      callback(error, null);
    }
  };
  reader.readAsArrayBuffer(file);
}

function normalizarCabecalho(valor) {
  return String(valor).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim().replace(/\s+/g, "_");
}

function buscarIndice(header, aliases) {
  for (const alias of aliases) {
    const index = header.indexOf(alias);
    if (index >= 0) return index;
  }
  return -1;
}

function detectarColunas(rows) {
  if (!rows.length) return null;

  const header = (rows[0] ?? []).map(normalizarCabecalho);
  const ig = buscarIndice(header, COLUMN_NAMES.gleba);
  const ip = buscarIndice(header, COLUMN_NAMES.ponto);
  const il = buscarIndice(header, COLUMN_NAMES.latitude);
  const io = buscarIndice(header, COLUMN_NAMES.longitude);
  const recognized = [ig, ip, il, io].filter((index) => index >= 0).length;

  if (recognized >= 2) {
    if ([ig, ip, il, io].some((index) => index < 0)) {
      return { error: "Cabecalho encontrado, mas faltam colunas obrigatorias: Gleba, Ponto, Latitude e Longitude." };
    }
    return { ig, ip, il, io, isHeader: true };
  }

  const maxColumns = Math.max(...rows.map((row) => row.length), 0);
  if (maxColumns < 4) {
    return { error: "A planilha precisa ter pelo menos quatro colunas de dados: Gleba, Ponto, Latitude e Longitude." };
  }

  return { ig: 0, ip: 1, il: 2, io: 3, isHeader: false };
}

function ptIguais(lat1, lon1, lat2, lon2) {
  return Math.abs(lat1 - lat2) < TOLERANCIA && Math.abs(lon1 - lon2) < TOLERANCIA;
}

function validar(rows, cols) {
  const { ig, ip, il, io, isHeader } = cols;
  const grupos = {};
  const erros = [];
  const inicio = isHeader ? 1 : 0;

  for (let i = inicio; i < rows.length; i += 1) {
    const row = rows[i];
    const gleba = String(row?.[ig] ?? "").trim();
    if (!gleba || gleba.toLowerCase() === "nan") continue;

    if (!grupos[gleba]) grupos[gleba] = [];
    grupos[gleba].push({
      linha: i + 1,
      seq: row?.[ip] ?? "",
      lat: row?.[il] ?? "",
      lon: row?.[io] ?? "",
    });
  }

  for (const [num, pontos] of Object.entries(grupos)) {
    const coords = [];

    for (const ponto of pontos) {
      const lat = parseFloat(String(ponto.lat).replace(",", "."));
      const lon = parseFloat(String(ponto.lon).replace(",", "."));

      if (Number.isNaN(lat) || Number.isNaN(lon)) {
        erros.push({
          gleba: num,
          linha: ponto.linha,
          seq: ponto.seq,
          tipo: "COORDENADA INVALIDA",
          msg: `Latitude "${ponto.lat}" ou longitude "${ponto.lon}" nao e numerica.`,
        });
        continue;
      }

      coords.push({ lat, lon, linha: ponto.linha });
    }

    if (!coords.length) continue;

    const primeiro = coords[0];
    const ultimo = coords[coords.length - 1];
    const fechado = ptIguais(primeiro.lat, primeiro.lon, ultimo.lat, ultimo.lon);

    if (!fechado) {
      erros.push({
        gleba: num,
        linha: ultimo.linha,
        seq: pontos[pontos.length - 1]?.seq ?? "-",
        tipo: "POLIGONO NAO FECHADO",
        msg: "O ultimo ponto nao repete o primeiro. Adicione no final uma linha igual a primeira coordenada.",
      });
    }

    const semFechamento = fechado ? coords.slice(0, -1) : coords;
    const verticesUnicos = new Set(semFechamento.map((coord) => `${coord.lat.toFixed(8)},${coord.lon.toFixed(8)}`));
    if (verticesUnicos.size < 3) {
      erros.push({
        gleba: num,
        linha: pontos[0]?.linha ?? "-",
        seq: "-",
        tipo: "PONTOS INSUFICIENTES",
        msg: `Foram encontrados apenas ${verticesUnicos.size} vertice(s) unico(s). O minimo exigido e 3.`,
      });
    }

    const contagem = {};
    for (const coord of coords) {
      const chave = `${coord.lat.toFixed(8)},${coord.lon.toFixed(8)}`;
      if (!contagem[chave]) contagem[chave] = [];
      contagem[chave].push(coord.linha);
    }

    for (const [coord, linhas] of Object.entries(contagem)) {
      if (linhas.length > 2) {
        erros.push({
          gleba: num,
          linha: linhas[0],
          seq: "-",
          tipo: "PONTO DUPLICADO EM EXCESSO",
          msg: `[${coord}] aparece ${linhas.length} vezes nas linhas ${linhas.slice(0, 4).join(", ")}.`,
        });
      }
    }
  }

  return { erros, grupos };
}

function iniciarValidacao() {
  if (!state.arquivoCarregado) {
    setStatus("Selecione um arquivo antes de validar.");
    return;
  }

  elements.btnValidar.disabled = true;
  elements.btnExportar.disabled = true;
  setStatus("Processando a planilha...");
  setProgress("indeterminate");

  lerPlanilha(state.arquivoCarregado, (error, rows) => {
    if (error) {
      setProgress("error");
      setStatus("Erro ao ler o arquivo.");
      renderizarRelatorioErro("Nao foi possivel abrir a planilha.", error.message);
      elements.btnValidar.disabled = false;
      return;
    }

    if (!rows.length) {
      setProgress("error");
      setStatus("A planilha esta vazia.");
      renderizarRelatorioErro("A planilha esta vazia ou nao contem linhas utilizaveis.");
      elements.btnValidar.disabled = false;
      return;
    }

    const cols = detectarColunas(rows);
    if (!cols || cols.error) {
      setProgress("error");
      setStatus("Estrutura da planilha invalida.");
      renderizarRelatorioErro(cols?.error ?? "Nao foi possivel identificar as colunas da planilha.");
      elements.btnValidar.disabled = false;
      return;
    }

    window.setTimeout(() => {
      const { erros, grupos } = validar(rows, cols);
      state.errosCache = erros;
      state.gruposCache = grupos;
      exibirResultado(erros, grupos);
      elements.btnValidar.disabled = false;
    }, 30);
  });
}

function exibirResultado(erros, grupos) {
  const total = Object.keys(grupos).length;
  const totalErros = erros.length;
  const glebasComErro = new Set(erros.map((erro) => erro.gleba)).size;
  const glebasOk = Math.max(total - glebasComErro, 0);

  elements.statGlebas.textContent = String(total);
  elements.statErros.textContent = String(totalErros);
  elements.statOk.textContent = String(glebasOk);

  if (!total) {
    setProgress("error");
    setStatus("Nenhuma gleba valida foi encontrada na planilha.");
  } else if (!erros.length) {
    setProgress("ok");
    setStatus(`${total} gleba(s) analisadas sem inconsistencias.`);
  } else {
    setProgress("error");
    setStatus(`${totalErros} erro(s) encontrados em ${glebasComErro} gleba(s).`);
  }

  elements.btnExportar.disabled = total === 0;
  renderizarRelatorioVisual(erros, grupos);
  renderizarGlebas(erros, grupos);
  showTab("relatorio");
}

function montarTexto(erros, grupos) {
  const agora = new Date().toLocaleString("pt-BR");
  const nomeArquivo = state.arquivoCarregado ? state.arquivoCarregado.name : "-";
  const total = Object.keys(grupos).length;
  const linhas = [
    "==================================================",
    "RELATORIO DE VALIDACAO - SICOR",
    "==================================================",
    "",
    `Arquivo : ${nomeArquivo}`,
    `Horario : ${agora}`,
    `Glebas  : ${total}`,
    `Erros   : ${erros.length}`,
    "",
  ];

  if (!total) {
    linhas.push("Nenhuma gleba foi encontrada na planilha.");
    return linhas.join("\n");
  }

  if (!erros.length) {
    linhas.push("RESULTADO");
    linhas.push("Todas as glebas foram validadas sem inconsistencias.");
    return linhas.join("\n");
  }

  const porTipo = {};
  for (const erro of erros) {
    if (!porTipo[erro.tipo]) porTipo[erro.tipo] = [];
    porTipo[erro.tipo].push(erro);
  }

  linhas.push("ERROS ENCONTRADOS");
  linhas.push("--------------------------------------------------");

  for (const [tipo, lista] of Object.entries(porTipo).sort(([a], [b]) => a.localeCompare(b, "pt-BR"))) {
    linhas.push(`${ERROR_ICONS[tipo] ?? "[ERRO]"} ${tipo} (${lista.length})`);
    for (const erro of lista) {
      const seq = String(erro.seq ?? "").trim();
      const ponto = seq && seq !== "-" ? ` | Ponto ${seq}` : "";
      linhas.push(`Gleba ${erro.gleba} | Linha ${erro.linha}${ponto}`);
      linhas.push(`  -> ${erro.msg}`);
    }
    linhas.push("");
  }

  linhas.push("COMO CORRIGIR");
  linhas.push("- Poligono nao fechado: repita a primeira coordenada na ultima linha da gleba.");
  linhas.push("- Pontos insuficientes: inclua pelo menos tres vertices unicos.");
  linhas.push("- Ponto duplicado em excesso: remova repeticoes fora do fechamento.");
  linhas.push("- Coordenada invalida: revise latitude e longitude na planilha.");
  return linhas.join("\n");
}

function renderizarGlebas(erros, grupos) {
  elements.glebasList.replaceChildren();

  const errosPorGleba = {};
  for (const erro of erros) {
    if (!errosPorGleba[erro.gleba]) errosPorGleba[erro.gleba] = [];
    errosPorGleba[erro.gleba].push(erro);
  }

  const entries = Object.entries(grupos).sort(([a], [b]) => a.localeCompare(b, "pt-BR", { numeric: true }));
  if (!entries.length) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = "Nenhuma gleba validada ainda.";
    elements.glebasList.appendChild(empty);
    return;
  }

  for (const [num, pontos] of entries) {
    const listaErros = errosPorGleba[num] ?? [];
    const card = document.createElement("article");
    card.className = `gleba-card ${listaErros.length ? "err" : "ok"}`;

    const header = document.createElement("button");
    header.type = "button";
    header.className = "gleba-header";
    header.setAttribute("aria-expanded", listaErros.length ? "true" : "false");

    const headingWrap = document.createElement("div");
    const title = document.createElement("h3");
    title.className = "gleba-title";
    title.textContent = `${listaErros.length ? "Com pendencias" : "Sem pendencias"} - Gleba ${num}`;
    const meta = document.createElement("p");
    meta.className = "gleba-meta";
    meta.textContent = `${pontos.length} linha(s) analisadas`;
    headingWrap.append(title, meta);

    const toggle = document.createElement("span");
    toggle.className = "gleba-toggle";
    toggle.setAttribute("aria-hidden", "true");
    toggle.textContent = "v";
    header.append(headingWrap, toggle);

    const body = document.createElement("div");
    body.className = "gleba-body";

    header.addEventListener("click", () => {
      const open = card.classList.toggle("open");
      header.setAttribute("aria-expanded", open ? "true" : "false");
    });

    if (listaErros.length) {
      listaErros.forEach((erro) => {
        const item = document.createElement("div");
        item.className = "error-item";

        const badge = document.createElement("span");
        badge.className = `badge ${ERROR_BADGES[erro.tipo] ?? "badge-red"}`;
        badge.textContent = erro.tipo;

        const message = document.createElement("div");
        message.textContent = `Linha ${erro.linha}: ${erro.msg}`;

        item.append(badge, message);
        body.appendChild(item);
      });
      card.classList.add("open");
    } else {
      const ok = document.createElement("p");
      ok.className = "ok-label";
      ok.textContent = "Gleba valida. Nenhum erro encontrado.";
      body.appendChild(ok);
    }

    card.append(header, body);
    elements.glebasList.appendChild(card);
  }
}

function atualizarResumoRelatorio(total, erros, ok, pendentes) {
  elements.reportTotalGlebas.textContent = String(total);
  elements.reportTotalErros.textContent = String(erros);
  elements.reportTotalOk.textContent = String(ok);
  elements.reportTotalPendentes.textContent = String(pendentes);
}

function criarIconeRelatorio(type) {
  const icon = document.createElement("span");
  icon.className = `report-icon report-icon-${type}`;
  icon.setAttribute("aria-hidden", "true");
  icon.textContent = UI_ICONS[type] ?? UI_ICONS.info;
  return icon;
}

function criarBlocoVisaoGeral(type, title, text) {
  const item = document.createElement("article");
  item.className = "report-overview-item";

  const icon = criarIconeRelatorio(type);
  const content = document.createElement("div");
  const heading = document.createElement("h4");
  heading.className = "report-overview-title";
  heading.textContent = title;
  const body = document.createElement("p");
  body.className = "report-overview-text";
  body.textContent = text;

  content.append(heading, body);
  item.append(icon, content);
  return item;
}

function detectarTipoGrupo(tipo) {
  if (tipo === "PONTOS INSUFICIENTES" || tipo === "PONTO DUPLICADO EM EXCESSO") {
    return "warning";
  }
  return "danger";
}

function renderizarRelatorioInicial() {
  atualizarResumoRelatorio("-", "-", "-", "-");
  elements.reportOverviewList.replaceChildren(
    criarBlocoVisaoGeral("info", "Pronto para validar", "Carregue uma planilha para visualizar o resumo operacional e as ocorrencias por tipo.")
  );

  const empty = document.createElement("p");
  empty.className = "empty-state";
  empty.textContent = "As inconsistencias aparecerao aqui apos a validacao.";
  elements.reportGroups.replaceChildren(empty);
}

function renderizarRelatorioErro(message, detail = "") {
  atualizarResumoRelatorio("0", "0", "0", "0");
  elements.reportOverviewList.replaceChildren(
    criarBlocoVisaoGeral("danger", "Falha na leitura da planilha", detail ? `${message} Detalhe tecnico: ${detail}` : message)
  );

  const empty = document.createElement("p");
  empty.className = "empty-state";
  empty.textContent = "Corrija a estrutura do arquivo e tente novamente.";
  elements.reportGroups.replaceChildren(empty);
}

function renderizarRelatorioVisual(erros, grupos) {
  const total = Object.keys(grupos).length;
  const pendentes = new Set(erros.map((erro) => erro.gleba)).size;
  const ok = Math.max(total - pendentes, 0);
  atualizarResumoRelatorio(total, erros.length, ok, pendentes);

  const overviewNodes = [];
  if (!total) {
    overviewNodes.push(
      criarBlocoVisaoGeral("warning", "Nenhuma gleba identificada", "A planilha foi lida, mas nao houve glebas validas para processar.")
    );
  } else if (!erros.length) {
    overviewNodes.push(
      criarBlocoVisaoGeral("success", "Validacao concluida com sucesso", "Todas as glebas analisadas estao validas e prontas para conferencia final.")
    );
    overviewNodes.push(
      criarBlocoVisaoGeral("info", "Leitura limpa", "Nenhum alerta ou erro foi encontrado nas coordenadas processadas.")
    );
  } else {
    overviewNodes.push(
      criarBlocoVisaoGeral("danger", "Foram encontradas inconsistencias", `${erros.length} ocorrencia(s) distribuidas em ${pendentes} gleba(s).`)
    );
    overviewNodes.push(
      criarBlocoVisaoGeral("warning", "Acao recomendada", "Revise primeiro os erros de fechamento e coordenadas invalidas, pois eles tendem a bloquear o processamento.")
    );
  }
  elements.reportOverviewList.replaceChildren(...overviewNodes);

  const porTipo = {};
  for (const erro of erros) {
    if (!porTipo[erro.tipo]) porTipo[erro.tipo] = [];
    porTipo[erro.tipo].push(erro);
  }

  if (!erros.length) {
    const successCard = document.createElement("article");
    successCard.className = "report-group-card";

    const head = document.createElement("div");
    head.className = "report-group-head";
    const titleWrap = document.createElement("div");
    const title = document.createElement("h4");
    title.className = "report-group-title";
    title.textContent = "Sem ocorrencias";
    const meta = document.createElement("p");
    meta.className = "report-group-meta";
    meta.textContent = "Todas as validacoes passaram sem pendencias.";
    titleWrap.append(title, meta);

    const badge = document.createElement("span");
    badge.className = "report-group-badge report-group-success";
    badge.textContent = `${UI_ICONS.success} Validado`;
    head.append(titleWrap, badge);
    successCard.appendChild(head);
    elements.reportGroups.replaceChildren(successCard);
    return;
  }

  const groupCards = Object.entries(porTipo)
    .sort(([a], [b]) => a.localeCompare(b, "pt-BR"))
    .map(([tipo, lista]) => {
      const card = document.createElement("article");
      card.className = "report-group-card";

      const head = document.createElement("div");
      head.className = "report-group-head";

      const titleWrap = document.createElement("div");
      const title = document.createElement("h4");
      title.className = "report-group-title";
      title.textContent = tipo;
      const meta = document.createElement("p");
      meta.className = "report-group-meta";
      meta.textContent = `${lista.length} ocorrencia(s) encontradas`;
      titleWrap.append(title, meta);

      const badge = document.createElement("span");
      const severity = detectarTipoGrupo(tipo);
      badge.className = `report-group-badge report-group-${severity}`;
      badge.textContent = `${severity === "danger" ? UI_ICONS.danger : UI_ICONS.warning} ${severity === "danger" ? "Erro" : "Alerta"}`;
      head.append(titleWrap, badge);

      const items = document.createElement("div");
      items.className = "report-items";

      lista.forEach((erro) => {
        const item = document.createElement("article");
        item.className = "report-item";

        const icon = criarIconeRelatorio(severity);
        const content = document.createElement("div");

        const metaRow = document.createElement("div");
        metaRow.className = "report-item-meta";
        const chipGleba = document.createElement("span");
        chipGleba.className = "report-chip";
        chipGleba.textContent = `Gleba ${erro.gleba}`;
        const chipLinha = document.createElement("span");
        chipLinha.className = "report-chip";
        chipLinha.textContent = `Linha ${erro.linha}`;
        metaRow.append(chipGleba, chipLinha);

        if (erro.seq && erro.seq !== "-") {
          const chipPonto = document.createElement("span");
          chipPonto.className = "report-chip";
          chipPonto.textContent = `Ponto ${erro.seq}`;
          metaRow.appendChild(chipPonto);
        }

        const text = document.createElement("p");
        text.className = "report-item-text";
        text.textContent = erro.msg;

        content.append(metaRow, text);
        item.append(icon, content);
        items.appendChild(item);
      });

      card.append(head, items);
      return card;
    });

  elements.reportGroups.replaceChildren(...groupCards);
}

function showTab(name) {
  elements.tabButtons.forEach((button) => {
    const active = button.dataset.tab === name;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", active ? "true" : "false");
    button.tabIndex = active ? 0 : -1;
  });

  elements.tabContents.forEach((panel) => {
    const active = panel.id === `tab-${name}`;
    panel.classList.toggle("active", active);
    panel.hidden = !active;
  });
}

function onTabKeydown(event) {
  const currentIndex = elements.tabButtons.indexOf(event.currentTarget);
  if (currentIndex < 0) return;

  let targetIndex = currentIndex;
  if (event.key === "ArrowRight") {
    targetIndex = (currentIndex + 1) % elements.tabButtons.length;
  } else if (event.key === "ArrowLeft") {
    targetIndex = (currentIndex - 1 + elements.tabButtons.length) % elements.tabButtons.length;
  } else {
    return;
  }

  event.preventDefault();
  const targetButton = elements.tabButtons[targetIndex];
  targetButton.focus();
  showTab(targetButton.dataset.tab);
}

function setProgress(status) {
  const fill = elements.progressFill;
  fill.classList.remove("indeterminate");

  if (status === "indeterminate") {
    fill.style.width = "0";
    fill.style.background = "var(--primary)";
    fill.classList.add("indeterminate");
    return;
  }
  if (status === "ok") {
    fill.style.width = "100%";
    fill.style.background = "var(--success)";
    return;
  }
  if (status === "error") {
    fill.style.width = "100%";
    fill.style.background = "var(--danger)";
    return;
  }

  fill.style.width = "0";
  fill.style.background = "var(--primary)";
}

function setStatus(text) {
  elements.status.textContent = text;
}

function limpar() {
  state.arquivoCarregado = null;
  state.errosCache = [];
  state.gruposCache = {};
  elements.fileInput.value = "";
  elements.fileInputSide.value = "";
  elements.dropZone.classList.remove("is-loaded", "is-hover");
  elements.dropKicker.textContent = "Planilha Excel";
  elements.dropTitle.textContent = "Arraste o arquivo aqui ou clique para selecionar";
  elements.dropSub.textContent = "Formatos aceitos: .xls e .xlsx";
  elements.btnValidar.disabled = true;
  elements.btnExportar.disabled = true;
  elements.statGlebas.textContent = "-";
  elements.statErros.textContent = "-";
  elements.statOk.textContent = "-";
  setProgress("reset");
  setStatus("Aguardando arquivo para iniciar a validacao.");
  renderizarRelatorioInicial();

  const empty = document.createElement("p");
  empty.className = "empty-state";
  empty.textContent = "Nenhum arquivo validado ainda.";
  elements.glebasList.replaceChildren(empty);
}

function exportarRelatorio() {
  const texto = montarTexto(state.errosCache, state.gruposCache);
  const blob = new Blob([texto], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "relatorio_glebas.txt";
  link.click();
  URL.revokeObjectURL(url);
  setStatus("Relatorio exportado com sucesso.");
}
