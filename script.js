let dados = [];

const selects = {
  mes: document.getElementById("mes"),
  turma: document.getElementById("turma"),
  disciplina: document.getElementById("disciplina"),
  unidade: document.getElementById("unidade")
};

const listaTurmas = document.getElementById("listaTurmas");
const limparTurma = document.getElementById("limparTurma");
const estadoInicial = document.getElementById("estadoInicial");
const listaResultado = document.getElementById("listaResultado");

function normalizar(texto) {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function limparValor(valor) {
  if (valor === undefined || valor === null) return "";
  return String(valor).trim().replace(/\s+/g, " ");
}

function chaveDisciplina(disciplina) {
  return normalizar(disciplina)
    .replace(/biostimulador/g, "bioestimulador");
}

function nomeDisciplinaPadrao(disciplina) {
  return limparValor(disciplina)
    .replace(/Biostimulador/gi, "Bioestimulador");
}

function nomeCurtoDisciplina(nomeCompleto) {
  const nome = nomeDisciplinaPadrao(nomeCompleto);
  const n = normalizar(nome);

  let tipo = "";
  // Práticas Clínicas
if (
    n.includes("pratica clinica") &&
    n.includes("toxina botulinica")
) {
    return "Prática Clínica: Toxina Botulínica";
}

if (
    n.includes("pratica clinica") &&
    n.includes("bioestimulador")
) {
    return "Prática Clínica: Bioestimuladores";
}

if (
    n.includes("pratica clinica") &&
    (
        n.includes("preenchimento facial") ||
        n.includes("preenchedor")
    )
) {
    return "Prática Clínica: Preenchedores";
}

if (
    n.includes("pratica clinica") &&
    n.includes("fios de sustentacao")
) {
    return "Prática Clínica: Fios de Sustentação";
}
  if (n.includes("teoria")) tipo = " - Teórica";
  if (n.includes("pratica")) tipo = " - Prática";

  if (n.includes("eletroterapia")) return "Eletroterapia" + tipo;
  if (n.includes("farmacologia")) return "Farmacologia" + tipo;
  if (n.includes("fios de sustentacao")) return "Fios de Sustentação" + tipo;
  if (n.includes("preenchimento facial")) return "Preenchimento Facial" + tipo;
  if (n.includes("bioestimulador")) return "Bioestimulador" + tipo;
  if (n.includes("toxina botulinica")) return "Toxina Botulínica" + tipo;
  if (n.includes("tricologia")) return "Tricologia" + tipo;

 if (
    n.includes("intradermicos") ||
    n.includes("intramusculares")
) {
    let complemento = "";

    if (n.includes("pratica i")) complemento = " - Prática I";
    else if (n.includes("pratica ii")) complemento = " - Prática II";
    else if (n.includes("teoria")) complemento = " - Teórica";
    else if (n.includes("pratica")) complemento = " - Prática";

    return "Intradérmicos e Intramusculares" + complemento;
}

  return nome;
}
function ordemTipoDisciplina(texto) {
  const n = normalizar(texto);

  if (n.includes("teorica") || n.includes("teoria")) return 1;
  if (n.includes("pratica")) return 2;

  return 3;
}

function lerPlanilha(workbook) {
  const registros = [];

  const nomeAba = workbook.SheetNames[0];
  const planilha = workbook.Sheets[nomeAba];

  const linhas = XLSX.utils.sheet_to_json(planilha, {
    defval: "",
    raw: false
  });

  linhas.forEach(linha => {
    registros.push({
      mes: limparValor(linha["MÊS"] || linha["MES"]),
      dia: limparValor(linha["DIA"]),
      turma: limparValor(linha["TURMA"]),
      disciplina: nomeDisciplinaPadrao(linha["DISCIPLINA"]),
      unidade: limparValor(linha["UNIDADE"]),
      avaliacaoRegular: limparValor(
        linha["AVALIAÇÃO REGULAR"] || linha["AVALIACAO REGULAR"]
      ),
      avaliacaoSubstitutiva: limparValor(
        linha["AVALIAÇÃO SUBSTITUTIVA"] || linha["AVALIACAO SUBSTITUTIVA"]
      ),
      origem: limparValor(linha["ABA_ORIGEM"])
    });
  });

  return registros.filter(item => item.turma || item.disciplina);
}

function ordemMes(mes) {
  const meses = {
    janeiro: 1,
    fevereiro: 2,
    marco: 3,
    abril: 4,
    maio: 5,
    junho: 6,
    julho: 7,
    agosto: 8,
    setembro: 9,
    outubro: 10,
    novembro: 11,
    dezembro: 12
  };

  return meses[normalizar(mes)] || 99;
}

function ehListaDeMeses(lista) {
  return lista.every(item => ordemMes(item) !== 99);
}

function unicos(lista) {
  const valores = [...new Set(lista.filter(Boolean))];

  if (ehListaDeMeses(valores)) {
    return valores.sort((a, b) => ordemMes(a) - ordemMes(b));
  }

  return valores.sort((a, b) =>
    String(a).localeCompare(String(b), "pt-BR")
  );
}

function disciplinasUnicas(lista) {
  const mapa = new Map();

  lista.forEach(item => {
    const chave = chaveDisciplina(item.disciplina);
    if (!chave) return;

    if (!mapa.has(chave)) {
      mapa.set(chave, nomeCurtoDisciplina(item.disciplina));
    }
  });

  return [...mapa.entries()]
    .map(([value, text]) => ({ value, text }))
    .sort((a, b) => {

    const praticaA = a.text.startsWith("Prática Clínica:");
    const praticaB = b.text.startsWith("Prática Clínica:");

    // Práticas Clínicas sempre por último
    if (praticaA && !praticaB) return 1;
    if (!praticaA && praticaB) return -1;

    // Remove o tipo para comparar o nome
    const nomeA = a.text
        .replace("Prática Clínica: ", "")
        .replace(" - Teórica", "")
        .replace(" - Prática I", "")
        .replace(" - Prática II", "")
        .replace(" - Prática", "");

    const nomeB = b.text
        .replace("Prática Clínica: ", "")
        .replace(" - Teórica", "")
        .replace(" - Prática I", "")
        .replace(" - Prática II", "")
        .replace(" - Prática", "");

    const ordemNome = nomeA.localeCompare(nomeB, "pt-BR");

    if (ordemNome !== 0) return ordemNome;

    return ordemTipoDisciplina(a.text) - ordemTipoDisciplina(b.text);
});
}

function preencherSelect(select, valores, textoPadrao) {
  const valorAtual = select.value;

  select.innerHTML = `<option value="">${textoPadrao}</option>`;

  valores.forEach(valor => {
    const option = document.createElement("option");
    option.value = valor;
    option.textContent = valor;
    select.appendChild(option);
  });

  if ([...select.options].some(option => option.value === valorAtual)) {
    select.value = valorAtual;
  }
}

function preencherSelectDisciplina(valores) {
  const valorAtual = selects.disciplina.value;

  selects.disciplina.innerHTML =
    `<option value="">Todas as disciplinas</option>`;

  valores.forEach(item => {
    const option = document.createElement("option");
option.value = item.value;
option.textContent = item.text;

if (item.text.startsWith("Prática Clínica:")) {
    option.classList.add("pratica-clinica");
}

selects.disciplina.appendChild(option);
  });

  if ([...selects.disciplina.options].some(option => option.value === valorAtual)) {
    selects.disciplina.value = valorAtual;
  }
}

function numeroMes(mes) {
  const meses = {
    janeiro: "01",
    fevereiro: "02",
    marco: "03",
    abril: "04",
    maio: "05",
    junho: "06",
    julho: "07",
    agosto: "08",
    setembro: "09",
    outubro: "10",
    novembro: "11",
    dezembro: "12"
  };

  return meses[normalizar(mes)] || "";
}

function dataDaAulaValida(item) {
  const meses = {
    janeiro: 0,
    fevereiro: 1,
    marco: 2,
    abril: 3,
    maio: 4,
    junho: 5,
    julho: 6,
    agosto: 7,
        setembro: 8,
    outubro: 9,
    novembro: 10,
    dezembro: 11
  };

  const mes = meses[normalizar(item.mes)];
  const dia = parseInt(item.dia);

  if (mes === undefined || !dia) return true;

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const dataAula = new Date(2026, mes, dia);
  dataAula.setHours(0, 0, 0, 0);

  return dataAula >= hoje;
}

function dadosAtivos() {
  return dados.filter(item => dataDaAulaValida(item));
}

function dataFormatada(item) {
  const mes = numeroMes(item.mes);

  if (!mes) return item.dia || "-";

  return `${String(item.dia).padStart(2, "0")}/${mes}/2026`;
}

function ehPraticaClinica(item) {
  const n = normalizar(item.disciplina);

  return (
    n.includes("pratica clinica") &&
    (
      n.includes("preenchimento facial") ||
      n.includes("fios de sustentacao")
    ) &&
    item.avaliacaoRegular === "Prática Clínica"
  );
}

function turmasDisponiveis() {
  const mes = selects.mes.value;
  const disciplina = selects.disciplina.value;
  const unidade = selects.unidade.value;

  let lista = dadosAtivos();

  if (mes) lista = lista.filter(item => item.mes === mes);

  if (disciplina) {
    lista = lista.filter(item =>
      chaveDisciplina(item.disciplina) === disciplina
    );
  }

  if (unidade) lista = lista.filter(item => item.unidade === unidade);

  return unicos(lista.map(item => item.turma));
}

function renderizarListaTurmas() {
  const busca = normalizar(selects.turma.value);

  const turmas = turmasDisponiveis().filter(turma =>
    normalizar(turma).includes(busca)
  );

  listaTurmas.innerHTML = "";

  if (!turmas.length) {
    const vazio = document.createElement("div");
    vazio.className = "item-turma vazio";
    vazio.textContent = "Nenhuma turma encontrada";
    listaTurmas.appendChild(vazio);
    listaTurmas.classList.remove("escondido");
    return;
  }

  turmas.forEach(turma => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "item-turma";
    item.textContent = turma;

    item.addEventListener("mousedown", event => {
      event.preventDefault();

      selects.turma.value = turma;
      listaTurmas.classList.add("escondido");

      atualizarFiltros();
      consultar();
    });

    listaTurmas.appendChild(item);
  });

  listaTurmas.classList.remove("escondido");
}

function atualizarFiltros() {
  const base = dadosAtivos();

  const mesAtual = selects.mes.value;
  const turmaAtual = selects.turma.value;
  const disciplinaAtual = selects.disciplina.value;
  const unidadeAtual = selects.unidade.value;

  let listaSemDisciplina = base;
  let listaSemUnidade = base;

  if (mesAtual) {
    listaSemDisciplina = listaSemDisciplina.filter(item => item.mes === mesAtual);
    listaSemUnidade = listaSemUnidade.filter(item => item.mes === mesAtual);
  }

  if (turmaAtual) {
    listaSemDisciplina = listaSemDisciplina.filter(item =>
      normalizar(item.turma).includes(normalizar(turmaAtual))
    );

    listaSemUnidade = listaSemUnidade.filter(item =>
      normalizar(item.turma).includes(normalizar(turmaAtual))
    );
  }

  if (unidadeAtual) {
    listaSemDisciplina = listaSemDisciplina.filter(item => item.unidade === unidadeAtual);
  }

  if (disciplinaAtual) {
    listaSemUnidade = listaSemUnidade.filter(item =>
      chaveDisciplina(item.disciplina) === disciplinaAtual
    );
  }

  preencherSelect(
    selects.mes,
    unicos(base.map(item => item.mes)),
    "Todos os meses"
  );

  preencherSelectDisciplina(
    disciplinasUnicas(listaSemDisciplina)
  );

  preencherSelect(
    selects.unidade,
    unicos(listaSemUnidade.map(item => item.unidade)),
    "Todas as unidades"
  );

  selects.mes.value = mesAtual;
  selects.turma.value = turmaAtual;
  selects.disciplina.value = disciplinaAtual;
  selects.unidade.value = unidadeAtual;
  destacarDisciplinaSelecionada();
}

function filtrarDados() {
  const mes = selects.mes.value;
  const turma = normalizar(selects.turma.value);
  const disciplina = selects.disciplina.value;
  const unidade = selects.unidade.value;

  return dadosAtivos().filter(item => {
    return (
      (!mes || item.mes === mes) &&
      (!turma || normalizar(item.turma).includes(turma)) &&
      (!disciplina || chaveDisciplina(item.disciplina) === disciplina) &&
      (!unidade || item.unidade === unidade)
    );
  });
}

function consultar() {
  const lista = filtrarDados();

  if (!lista.length) {
    estadoInicial.textContent = "Nenhum resultado encontrado.";
    estadoInicial.classList.remove("escondido");
    listaResultado.classList.add("escondido");
    return;
  }

  estadoInicial.classList.add("escondido");
  listaResultado.classList.remove("escondido");

  listaResultado.innerHTML = lista.map(item => {
    const pratica = ehPraticaClinica(item);

    return `
      <article class="card">
        <h2>${item.disciplina || "-"}</h2>

        <div class="grid">
          <div>
            <span>Mês</span>
            <strong>${item.mes || "-"}</strong>
          </div>

          <div>
            <span>Data da aula</span>
            <strong>${dataFormatada(item)}</strong>
          </div>

          <div>
            <span>Código da turma</span>
            <strong>${item.turma || "-"}</strong>
          </div>

          <div>
            <span>Unidade</span>
            <strong>${item.unidade || "-"}</strong>
          </div>

          <div class="${pratica ? "sem-prova" : ""}">
            <span>Avaliação regular</span>
            <strong>${pratica ? "Não possui data de prova" : item.avaliacaoRegular || "Não informado"}</strong>
          </div>

          <div class="${pratica ? "sem-prova" : ""}">
            <span>Avaliação substitutiva</span>
            <strong>${pratica ? "Não possui data de prova" : item.avaliacaoSubstitutiva || "Não informado"}</strong>
          </div>
        </div>
      </article>
    `;
  }).join("");
}

function limparConsulta() {
  selects.mes.value = "";
  selects.turma.value = "";
  selects.disciplina.value = "";
  selects.unidade.value = "";

  listaTurmas.classList.add("escondido");

  atualizarFiltros();
  consultar();
}

selects.mes.addEventListener("change", () => {
  atualizarFiltros();
  consultar();
});

selects.turma.addEventListener("focus", () => {
  renderizarListaTurmas();
});

selects.turma.addEventListener("click", () => {
  renderizarListaTurmas();
});

selects.turma.addEventListener("input", () => {
  renderizarListaTurmas();
  atualizarFiltros();
  consultar();
});

selects.turma.addEventListener("blur", () => {
  setTimeout(() => {
    listaTurmas.classList.add("escondido");
  }, 150);
});

limparTurma.addEventListener("click", () => {
  selects.turma.value = "";
  selects.turma.focus();

  atualizarFiltros();
  consultar();
  renderizarListaTurmas();
});

selects.disciplina.addEventListener("change", () => {
  atualizarFiltros();
  consultar();
  destacarDisciplinaSelecionada();
});

selects.unidade.addEventListener("change", () => {
  atualizarFiltros();
  consultar();
});

document.getElementById("limpar").addEventListener("click", limparConsulta);

document.addEventListener("click", event => {
  const campoTurma = document.querySelector(".campo-turma");

  if (campoTurma && !campoTurma.contains(event.target)) {
    listaTurmas.classList.add("escondido");
  }
});

async function iniciar() {
  try {
    estadoInicial.textContent = "Carregando planilha...";

    const resposta = await fetch("./cronograma.xlsx?v=" + Date.now());

    if (!resposta.ok) {
      throw new Error("Planilha não encontrada.");
    }

    const arrayBuffer = await resposta.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    dados = lerPlanilha(workbook);

    atualizarFiltros();
    consultar();

  } catch (erro) {
    console.error(erro);
    estadoInicial.textContent =
      "Não foi possível carregar a planilha cronograma.xlsx.";
  }
}
function destacarDisciplinaSelecionada() {
  const textoSelecionado =
    selects.disciplina.options[selects.disciplina.selectedIndex]?.textContent || "";

  if (textoSelecionado.startsWith("Prática Clínica:")) {
    selects.disciplina.classList.add("pratica-clinica-selecionada");
  } else {
    selects.disciplina.classList.remove("pratica-clinica-selecionada");
  }
}
iniciar();