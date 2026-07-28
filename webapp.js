/**
 * Página pública de consulta. A implantação do Web App deve ser
 * configurada para executar como o proprietário do projeto.
 */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("portal_individual")
    .setTitle("Consulta do Coro")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.DEFAULT
    );
}

/**
 * Função pública chamada pelo browser através de google.script.run.
 */
function obterOpcoesConsultaPublica() {
  const contexto = lerContextoConsultaPublica_();
  const nomes = obterNomesDistribuicaoPublica_(
    contexto.spreadsheet
  );
  const meses = Array.from(
    new Set(
      contexto.linhas.map(linha =>
        normalizarChaveMesPublica_(
          linha[contexto.indices.mes]
        )
      )
    )
  )
    .filter(Boolean)
    .sort()
    .reverse()
    .map(valor => ({
      valor,
      texto: formatarChaveMes_(valor)
    }));

  return {
    nomes,
    meses,
    ultimaAtualizacao:
      obterUltimaAtualizacaoPublica_(contexto.spreadsheet)
  };
}

/**
 * Devolve apenas dados do snapshot público, sem tocar nas folhas
 * privadas nem executar cálculos de distribuição.
 */
function consultarResumoIndividual(nome, mes) {
  const nomePedido = String(nome || "").trim();
  const mesPedido = String(mes || "").trim();

  if (!nomePedido || !/^\d{4}-\d{2}$/.test(mesPedido)) {
    throw new Error("Seleciona um membro e um mês válidos.");
  }

  const contexto = lerContextoConsultaPublica_();
  const chaveNome = normalizarNome_(nomePedido);
  const linhas = contexto.linhas.filter(linha =>
    normalizarNome_(linha[contexto.indices.nome]) ===
      chaveNome &&
    normalizarChaveMesPublica_(
      linha[contexto.indices.mes]
    ) ===
      mesPedido
  );

  const nomesPublicados = obterNomesDistribuicaoPublica_(
    contexto.spreadsheet
  );
  const nomePublicado = nomesPublicados.find(item =>
    normalizarNome_(item) === chaveNome
  );

  if (!nomePublicado) {
    throw new Error(
      `O membro "${nomePedido}" não existe na distribuição publicada.`
    );
  }

  const ensaios = [];
  const concertos = [];
  let pontosObtidos = 0;
  let pontosMaximos = 0;
  let atividades = 0;

  linhas.forEach(linha => {
    const tipo = String(
      linha[contexto.indices.tipo] ?? ""
    ).trim();
    const registo = {
      data: formatarDataPortal_(
        linha[contexto.indices.data]
      ),
      origem: String(
        linha[contexto.indices.origem] ?? ""
      ).trim(),
      tipo,
      resposta: String(
        linha[contexto.indices.resposta] ?? ""
      ).trim(),
      pontosObtidos: Number(
        linha[contexto.indices.pontosObtidos]
      ) || 0,
      pontosMaximos: Number(
        linha[contexto.indices.pontosMaximos]
      ) || 0
    };

    if (
      normalizarTexto_(tipo) ===
      normalizarTexto_(
        DISTRIBUICAO_CONFIG.valores.tipoConcerto
      )
    ) {
      concertos.push(registo);

      if (
        converterBooleano_(
          linha[contexto.indices.contaAtividade]
        )
      ) {
        atividades += 1;
      }
    } else {
      ensaios.push(registo);
      pontosObtidos += registo.pontosObtidos;
      pontosMaximos += registo.pontosMaximos;
    }
  });

  return {
    nome: nomePublicado,
    mes: mesPedido,
    mesTexto: formatarChaveMes_(mesPedido),
    ultimaAtualizacao:
      obterUltimaAtualizacaoPublica_(contexto.spreadsheet),
    resumo: {
      atividades,
      ensaios: ensaios.length,
      pontosObtidos,
      pontosMaximos,
      assiduidade:
        pontosMaximos > 0
          ? pontosObtidos / pontosMaximos
          : 0
    },
    concertos,
    ensaios,
    distribuicao: obterDistribuicaoPublicaMembro_(
      contexto.spreadsheet,
      nomePedido
    )
  };
}

function obterNomesDistribuicaoPublica_(spreadsheet) {
  const folha = spreadsheet.getSheetByName(
    SHEETS.DISTRIBUTION
  );

  if (!folha || folha.getLastRow() < 2) {
    return [];
  }

  const dados = folha.getDataRange().getValues();
  const indiceNome = encontrarIndiceCabecalho_(
    dados[0],
    ["Nome"],
    true,
    folha.getName()
  );

  return Array.from(
    new Set(
      dados
        .slice(1)
        .map(linha =>
          String(linha[indiceNome] ?? "").trim()
        )
        .filter(nome =>
          nome &&
          normalizarTexto_(nome) !== "total"
        )
    )
  )
    .sort((a, b) =>
      a.localeCompare(b, "pt", { sensitivity: "base" })
    );
}

function lerContextoConsultaPublica_() {
  const spreadsheet = SpreadsheetApp.openById(
    PUBLICATION.SPREADSHEET_ID
  );
  const folha = spreadsheet.getSheetByName(
    PUBLICATION.DATA_SHEET
  );

  if (!folha || folha.getLastRow() < 1) {
    throw new Error(
      "Ainda não existem dados publicados para consulta. Executa primeiro «Publicar distribuição»."
    );
  }

  const dados = folha.getDataRange().getValues();
  const cabecalhos = dados[0];

  return {
    spreadsheet,
    linhas: dados.slice(1),
    indices: {
      nome: encontrarIndiceCabecalho_(
        cabecalhos,
        ["Nome"],
        true,
        folha.getName()
      ),
      data: encontrarIndiceCabecalho_(
        cabecalhos,
        ["Data"],
        true,
        folha.getName()
      ),
      mes: encontrarIndiceCabecalho_(
        cabecalhos,
        ["Mês"],
        true,
        folha.getName()
      ),
      origem: encontrarIndiceCabecalho_(
        cabecalhos,
        ["Origem"],
        true,
        folha.getName()
      ),
      tipo: encontrarIndiceCabecalho_(
        cabecalhos,
        ["Tipo"],
        true,
        folha.getName()
      ),
      resposta: encontrarIndiceCabecalho_(
        cabecalhos,
        ["Resposta"],
        true,
        folha.getName()
      ),
      pontosObtidos: encontrarIndiceCabecalho_(
        cabecalhos,
        ["Pontos obtidos"],
        true,
        folha.getName()
      ),
      pontosMaximos: encontrarIndiceCabecalho_(
        cabecalhos,
        ["Pontos máximos"],
        true,
        folha.getName()
      ),
      contaAtividade: encontrarIndiceCabecalho_(
        cabecalhos,
        ["Conta como atividade"],
        true,
        folha.getName()
      )
    }
  };
}

function obterDistribuicaoPublicaMembro_(
  spreadsheet,
  nome
) {
  const folha = spreadsheet.getSheetByName(
    SHEETS.DISTRIBUTION
  );

  if (!folha || folha.getLastRow() < 2) {
    return null;
  }

  const dados = folha.getDataRange().getValues();
  const cabecalhos = dados[0];
  const indiceNome = encontrarIndiceCabecalho_(
    cabecalhos,
    ["Nome"],
    true,
    folha.getName()
  );
  const linha = dados.slice(1).find(item =>
    normalizarNome_(item[indiceNome]) ===
    normalizarNome_(nome)
  );

  if (!linha) {
    return null;
  }

  const obter = cabecalho => {
    const indice = encontrarIndiceCabecalho_(
      cabecalhos,
      [cabecalho],
      false
    );

    return indice >= 0 ? linha[indice] : null;
  };

  return {
    atividades: Number(obter("Atividades")) || 0,
    assiduidade: Number(obter("Assiduidade")) || 0,
    pontos: Number(obter("Pontos")) || 0,
    valorFundoComum:
      Number(obter("Valor Fundo Comum")) || 0,
    valorIndividual:
      Number(obter("Valor Individual")) || 0,
    valorApoios:
      Number(obter("Valor Apoios")) || 0,
    valorFinal:
      Number(obter("Valor Final")) || 0
  };
}

function obterUltimaAtualizacaoPublica_(spreadsheet) {
  const folha = spreadsheet.getSheetByName(
    PUBLICATION.INFO_SHEET
  );

  if (!folha || folha.getLastRow() < 2) {
    return "";
  }

  const dados = folha.getDataRange().getValues();
  const linha = dados.find(item =>
    normalizarTexto_(item[0]) ===
    normalizarTexto_("Última atualização")
  );

  return linha
    ? formatarDataHoraPortal_(linha[1])
    : "";
}

function formatarChaveMes_(chave) {
  const correspondencia = String(chave || "")
    .match(/^(\d{4})-(\d{2})$/);

  if (!correspondencia) {
    return String(chave || "");
  }

  const nomesMeses = [
    "Janeiro", "Fevereiro", "Março", "Abril",
    "Maio", "Junho", "Julho", "Agosto",
    "Setembro", "Outubro", "Novembro", "Dezembro"
  ];
  const indiceMes = Number(correspondencia[2]) - 1;

  return `${nomesMeses[indiceMes]} de ${correspondencia[1]}`;
}

function normalizarChaveMesPublica_(valor) {
  if (
    valor instanceof Date &&
    !isNaN(valor.getTime())
  ) {
    return [
      valor.getFullYear(),
      String(valor.getMonth() + 1).padStart(2, "0")
    ].join("-");
  }

  const texto = String(valor || "").trim();

  return /^\d{4}-\d{2}$/.test(texto)
    ? texto
    : "";
}

function formatarDataPortal_(valor) {
  const data = converterData_(valor);

  if (!(data instanceof Date) || isNaN(data.getTime())) {
    return "";
  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );
}

function formatarDataHoraPortal_(valor) {
  const data = converterData_(valor);

  if (!(data instanceof Date) || isNaN(data.getTime())) {
    return "";
  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );
}

function converterBooleano_(valor) {
  return valor === true ||
    normalizarTexto_(valor) === "true" ||
    normalizarTexto_(valor) === "sim";
}
