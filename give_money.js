/**
 * ============================================================
 * CONFIGURAÇÃO
 * ============================================================
 */

const DISTRIBUICAO_CONFIG = {
  folhas: {
    get membros() {
      return SHEETS.MEMBERS;
    },
    get movimentos() {
      return SHEETS.MOVEMENTS;
    },
    get presencasGerais() {
      return SHEETS.GENERAL_ATTENDANCE;
    },
    get ensaiosNaipe() {
      return SHEETS.SECTION_ATTENDANCE;
    },
    get distribuicao() {
      return SHEETS.DISTRIBUTION;
    }
  },

  pesoAtividades: 0.75,
  valores: {
    estadoAtivo: "Ativo",
    tipoConcerto: "Concerto",
    tipoEnsaio: "Ensaio",
    presente: "Presente",
    atraso: "Atraso",
    falta: "Falta"
  },
  

  cabecalhos: {
    membros: {
      ordem: ["Ordem"],
      nome: ["Nome"],
      entrada: ["Entrada", "Data de Entrada"],
      estado: ["Estado"]
    },

    movimentos: {
      receita: ["Receita"],
      despesa: ["Despesa"]
    },

    presencas: {
      timestamp: [
        "Timestamp",
        "Timestap",
        "Carimbo de data/hora"
      ],
      tipoAtividade: [
        "Tipo de Atividade",
        "Tipo de atividade",
        "Tipo da Atividade"
      ],

      dataAtividade: [
        "Data ",
        "Data da Atividade",
        "Data de Atividade",
        "Data do Ensaio",
        "Data do Ensaio:"
      ]
    },

    distribuicao: {
      ordem: "Ordem",
      nome: "Nome",
      atividades: "Atividades",
      assiduidade: "Assiduidade",
      pontos: "Pontos",
      valorFundoComum: "Valor Fundo Comum",
      valorIndividual: "Valor Individual",
      valorApoios: "Valor Apoios",
      valorFinal: "Valor Final"
    }
  }
};


/**
 * ============================================================
 * FUNÇÃO PRINCIPAL
 * ============================================================
 */

/**
 * Gera ou atualiza a folha Distribuição.
 *
 * Preserva os valores manuais das colunas:
 * - Valor Individual
 * - Valor Apoios
 */
function gerarDistribuicao() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const membros = lerMembrosAtivos_(spreadsheet);

  if (membros.length === 0) {
    throw new Error(
      'Não foram encontrados membros com Estado igual a "Ativo".'
    );
  }

  const saldoFundoComum = calcularSaldoMovimentos_(
    spreadsheet,
    DISTRIBUICAO_CONFIG.folhas.movimentos
  );

  const analiseParticipacao =
  analisarParticipacao_(
    spreadsheet,
    membros
  );
  const metricasPorMembro = analiseParticipacao.metricas;

const totalPontos = membros.reduce(
  (soma, membro) => {
    const metricas =
      metricasPorMembro[membro.chaveNome];

    return soma + (metricas?.pontos || 0);
  },
  0
);

  const valoresManuais = lerValoresManuaisExistentes_(
    spreadsheet
  );

  const linhas = membros.map(membro => {
    const metricas =
    metricasPorMembro[membro.chaveNome] || {
    atividades: 0,
    assiduidade: 0,
    pontos: 0
  };

    const atividades = metricas.atividades;
    const assiduidade = metricas.assiduidade;
    const pontos = metricas.pontos;

    const valorFundoComum =
      totalPontos > 0
        ? pontos * saldoFundoComum / totalPontos
        : 0;

    const valoresExistentes =
      valoresManuais[membro.chaveNome] || {};

    return {
      ordem: membro.ordem,
      nome: membro.nome,
      atividades: atividades,  
      assiduidade: assiduidade,
      pontos: pontos,
      valorFundoComum: valorFundoComum,
      valorIndividual:
        converterNumero_(valoresExistentes.valorIndividual),
      valorApoios:
        converterNumero_(valoresExistentes.valorApoios)
    };
  });

  escreverDistribuicao_(
    spreadsheet,
    linhas,
    {
      saldoFundoComum,
      totalPontos,
      fontesOmitidas: analiseParticipacao.fontesOmitidas
    }
  );

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    "Distribuição atualizada",
    [
      `Membros ativos: ${membros.length}`,
      `Total de pontos: ${totalPontos}`,
      `Fundo comum: ${formatarEuro_(saldoFundoComum)}`,
      analiseParticipacao.fontesOmitidas.length > 0
        ? `Aviso — folhas ainda ausentes: ${analiseParticipacao.fontesOmitidas.join(", ")}`
        : "Todas as fontes de presenças foram processadas."
    ].join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * ============================================================
 * LEITURA DOS MEMBROS
 * ============================================================
 */

/**
 * Lê apenas os membros cujo Estado seja Ativo.
 */
function lerMembrosAtivos_(spreadsheet) {
  const nomeFolha =
    DISTRIBUICAO_CONFIG.folhas.membros;

  const folha = obterFolhaObrigatoria_(
    spreadsheet,
    nomeFolha
  );

  const ultimaLinha = folha.getLastRow();
  const ultimaColuna = folha.getLastColumn();

  if (ultimaLinha < 2 || ultimaColuna < 1) {
    throw new Error(
      `A folha "${nomeFolha}" não contém membros.`
    );
  }

  const dados = folha
    .getRange(1, 1, ultimaLinha, ultimaColuna)
    .getValues();

  const cabecalhos = dados[0];

  const indiceOrdem = encontrarIndiceCabecalho_(
    cabecalhos,
    DISTRIBUICAO_CONFIG.cabecalhos.membros.ordem,
    false
  );

  const indiceNome = encontrarIndiceCabecalho_(
    cabecalhos,
    DISTRIBUICAO_CONFIG.cabecalhos.membros.nome,
    true,
    nomeFolha
  );

  const indiceEntrada = encontrarIndiceCabecalho_(
    cabecalhos,
    DISTRIBUICAO_CONFIG.cabecalhos.membros.entrada,
    true,
    nomeFolha
  );

  const indiceEstado = encontrarIndiceCabecalho_(
    cabecalhos,
    DISTRIBUICAO_CONFIG.cabecalhos.membros.estado,
    true,
    nomeFolha
  );

  const membros = [];

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];

    const nome = String(
      linha[indiceNome] ?? ""
    ).trim();

    if (!nome) {
      continue;
    }

    const estado = String(
      linha[indiceEstado] ?? ""
    ).trim();

    if (
      normalizarTexto_(estado) !==
      normalizarTexto_(
        DISTRIBUICAO_CONFIG.valores.estadoAtivo
      )
    ) {
      continue;
    }

    const ordemOriginal =
      indiceOrdem >= 0
        ? linha[indiceOrdem]
        : membros.length + 1;

    const ordemNumerica =
      Number(ordemOriginal);

    membros.push({
      ordem:
        Number.isFinite(ordemNumerica)
          ? ordemNumerica
          : membros.length + 1,
      nome: nome,
      chaveNome: normalizarNome_(nome),
      entrada: converterData_(linha[indiceEntrada]),
      linhaOriginal: i + 1
    });
  }

  membros.sort((a, b) => {
    if (a.ordem !== b.ordem) {
      return a.ordem - b.ordem;
    }

    return a.nome.localeCompare(
      b.nome,
      "pt",
      { sensitivity: "base" }
    );
  });

  return membros;
}


/**
 * ============================================================
 * MOVIMENTOS FINANCEIROS
 * ============================================================
 */

/**
 * Calcula:
 *
 * soma das receitas - soma das despesas
 *
 * A leitura termina na primeira linha completamente vazia.
 */
function calcularSaldoMovimentos_(
  spreadsheet,
  nomeFolha
) {
  const folha = obterFolhaObrigatoria_(
    spreadsheet,
    nomeFolha
  );

  const ultimaLinha = folha.getLastRow();
  const ultimaColuna = folha.getLastColumn();

  if (ultimaLinha < 2 || ultimaColuna < 1) {
    return 0;
  }

  const dados = folha
    .getRange(1, 1, ultimaLinha, ultimaColuna)
    .getValues();

  const cabecalhos = dados[0];

  const indiceReceita = encontrarIndiceCabecalho_(
    cabecalhos,
    DISTRIBUICAO_CONFIG.cabecalhos.movimentos.receita,
    true,
    nomeFolha
  );

  const indiceDespesa = encontrarIndiceCabecalho_(
    cabecalhos,
    DISTRIBUICAO_CONFIG.cabecalhos.movimentos.despesa,
    true,
    nomeFolha
  );

  let totalReceita = 0;
  let totalDespesa = 0;

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];

    if (linhaCompletamenteVazia_(linha)) {
      break;
    }

    totalReceita += converterNumero_(
      linha[indiceReceita]
    );

    totalDespesa += converterNumero_(
      linha[indiceDespesa]
    );
  }

  return totalReceita - totalDespesa;
}


/**
 * ============================================================
 * PONTOS DOS CONCERTOS
 * ============================================================
 */

/**
 * Calcula, para cada membro ativo:
 *
 * - Atividades: número de concertos em que esteve Presente
 * - Assiduidade:
 *     pontos obtidos nos ensaios /
 *     pontos máximos dos ensaios desde a data de entrada
 * - Pontos:
 *     Atividades ×
 *     (pesoAtividades + (1 - pesoAtividades) × Assiduidade)
 */
function calcularMetricasParticipacao_(
  spreadsheet,
  membros
) {
  return analisarParticipacao_(spreadsheet, membros).metricas;
}

/**
 * Lê todas as fontes de presença e devolve métricas, detalhes
 * de diagnóstico e a lista de folhas opcionais ainda ausentes.
 */
function analisarParticipacao_(spreadsheet, membros) {
  const metricas = {};
  const detalhes = [];
  const avisos = [];
  const fontesOmitidas = [];
  const membrosPorChave = {};

  membros.forEach(membro => {
    if (membrosPorChave[membro.chaveNome]) {
      throw new Error(
        `Existem dois membros ativos com o mesmo nome normalizado: "${membro.nome}".`
      );
    }

    membrosPorChave[membro.chaveNome] = membro;
    metricas[membro.chaveNome] = {
      atividades: 0,
      pontosEnsaios: 0,
      ensaiosPossiveis: 0,
      assiduidade: 0,
      pontos: 0
    };
  });

  const fonteGeral = lerFontePresencas_(
    spreadsheet,
    DISTRIBUICAO_CONFIG.folhas.presencasGerais,
    {
      tipo: "geral",
      obrigatoria: true,
      membrosPorChave
    }
  );

  processarFontePresencas_(
    fonteGeral,
    membrosPorChave,
    metricas,
    detalhes,
    avisos
  );

  DISTRIBUICAO_CONFIG.folhas.ensaiosNaipe
    .forEach(nomeFolha => {
      const fonte = lerFontePresencas_(
        spreadsheet,
        nomeFolha,
        {
          tipo: "naipe",
          obrigatoria: false,
          membrosPorChave
        }
      );

      if (fonte.ausente) {
        fontesOmitidas.push(nomeFolha);
        avisos.push(fonte.aviso);
        return;
      }

      processarFontePresencas_(
        fonte,
        membrosPorChave,
        metricas,
        detalhes,
        avisos
      );
    });

  const pesoAtividades =
    DISTRIBUICAO_CONFIG.pesoAtividades;

  if (
    typeof pesoAtividades !== "number" ||
    pesoAtividades < 0 ||
    pesoAtividades > 1
  ) {
    throw new Error(
      "pesoAtividades deve ser um número entre 0 e 1."
    );
  }

  membros.forEach(membro => {
    const dadosMembro =
      metricas[membro.chaveNome];

    const maximoEnsaios =
      dadosMembro.ensaiosPossiveis * 5;

    dadosMembro.assiduidade =
      maximoEnsaios > 0
        ? dadosMembro.pontosEnsaios / maximoEnsaios
        : 0;

    dadosMembro.pontos =
      dadosMembro.atividades *
      (
        pesoAtividades +
        (1 - pesoAtividades) *
        dadosMembro.assiduidade
      );
  });

  avisos.forEach(aviso => Logger.log(aviso));

  return {
    metricas,
    detalhes,
    avisos,
    fontesOmitidas
  };
}

/**
 * Lê e valida uma fonte com Timestamp | Data | membros...
 * A fonte geral exige adicionalmente Tipo de atividade.
 */
function lerFontePresencas_(
  spreadsheet,
  nomeFolha,
  opcoes
) {
  const folha = spreadsheet.getSheetByName(nomeFolha);

  if (!folha) {
    if (!opcoes.obrigatoria) {
      return {
        ausente: true,
        aviso: `A folha opcional "${nomeFolha}" ainda não existe e foi omitida.`
      };
    }

    throw new Error(`Não existe a folha obrigatória "${nomeFolha}".`);
  }

  const ultimaLinha = folha.getLastRow();
  const ultimaColuna = folha.getLastColumn();

  if (ultimaLinha < 1 || ultimaColuna < 1) {
    if (!opcoes.obrigatoria) {
      return {
        ausente: true,
        aviso: `A folha opcional "${nomeFolha}" está vazia e foi omitida.`
      };
    }

    throw new Error(`A folha "${nomeFolha}" não contém cabeçalhos.`);
  }

  const dados = folha
    .getRange(1, 1, Math.max(ultimaLinha, 1), ultimaColuna)
    .getValues();
  const cabecalhos = dados[0];
  const configCabecalhos = DISTRIBUICAO_CONFIG.cabecalhos.presencas;

  const indiceTimestamp = encontrarIndiceCabecalho_(
    cabecalhos,
    configCabecalhos.timestamp,
    true,
    nomeFolha
  );
  const indiceData = encontrarIndiceCabecalho_(
    cabecalhos,
    configCabecalhos.dataAtividade,
    true,
    nomeFolha
  );
  const indiceTipo = opcoes.tipo === "geral"
    ? encontrarIndiceCabecalho_(
        cabecalhos,
        configCabecalhos.tipoAtividade,
        true,
        nomeFolha
      )
    : -1;

  const indicesMetadados = new Set([
    indiceTimestamp,
    indiceData,
    indiceTipo
  ]);
  const colunasPresencas = [];
  const chavesEncontradas = new Set();
  const nomesDesconhecidos = [];

  cabecalhos.forEach((cabecalho, indiceColuna) => {
    if (indicesMetadados.has(indiceColuna)) {
      return;
    }

    const nomeExtraido = extrairNomeCabecalhoPresenca_(cabecalho);

    if (!nomeExtraido) {
      return;
    }

    const chaveNome = normalizarNome_(nomeExtraido);

    if (!opcoes.membrosPorChave[chaveNome]) {
      nomesDesconhecidos.push(nomeExtraido);
      return;
    }

    if (chavesEncontradas.has(chaveNome)) {
      throw new Error(
        `A folha "${nomeFolha}" contém mais de uma coluna para "${nomeExtraido}".`
      );
    }

    chavesEncontradas.add(chaveNome);
    colunasPresencas.push({
      indiceColuna,
      chaveNome,
      nomeCabecalho: nomeExtraido
    });
  });

  if (colunasPresencas.length === 0) {
    if (opcoes.obrigatoria) {
      throw new Error(
        `Não foram encontradas colunas de membros ativos em "${nomeFolha}".`
      );
    }

  }

  return {
    ausente: false,
    nomeFolha,
    tipo: opcoes.tipo,
    dados,
    indiceData,
    indiceTipo,
    colunasPresencas,
    nomesDesconhecidos,
    semMembrosAtivos: colunasPresencas.length === 0
  };
}

function processarFontePresencas_(
  fonte,
  membrosPorChave,
  metricas,
  detalhes,
  avisos
) {
  if (fonte.semMembrosAtivos) {
    avisos.push(
      `A folha "${fonte.nomeFolha}" não contém colunas correspondentes a membros ativos.`
    );
  }

  fonte.nomesDesconhecidos.forEach(nome => {
    avisos.push(
      `A folha "${fonte.nomeFolha}" contém o membro desconhecido "${nome}".`
    );
  });

  for (let i = 1; i < fonte.dados.length; i++) {
    const linha = fonte.dados[i];

    if (linhaCompletamenteVazia_(linha)) {
      continue;
    }

    const dataOriginal = linha[fonte.indiceData];
    const dataAtividade = converterData_(dataOriginal);
    const dataValida =
      dataAtividade instanceof Date &&
      !isNaN(dataAtividade.getTime());
    const tipoOriginal = fonte.tipo === "naipe"
      ? DISTRIBUICAO_CONFIG.valores.tipoEnsaio
      : String(linha[fonte.indiceTipo] ?? "").trim();
    const tipoNormalizado = fonte.tipo === "naipe"
      ? normalizarTexto_(DISTRIBUICAO_CONFIG.valores.tipoEnsaio)
      : normalizarTexto_(tipoOriginal);
    const tipoRegisto = fonte.tipo === "naipe"
      ? "Ensaio de naipe"
      : tipoOriginal;

    for (const coluna of fonte.colunasPresencas) {
      const membro = membrosPorChave[coluna.chaveNome];
      const respostaOriginal = String(
        linha[coluna.indiceColuna] ?? ""
      ).trim();
      const resultado = avaliarRegistoPresenca_({
        fonte: fonte.nomeFolha,
        linha: i + 1,
        dataAtividade,
        dataValida,
        tipoNormalizado,
        tipoRegisto,
        respostaOriginal,
        membro,
        origemNaipe: fonte.tipo === "naipe"
      });

      if (resultado.contaComoAtividade) {
        metricas[coluna.chaveNome].atividades += 1;
      }

      if (resultado.entraDenominador) {
        metricas[coluna.chaveNome].ensaiosPossiveis += 1;
        metricas[coluna.chaveNome].pontosEnsaios +=
          resultado.pontosObtidos;
      }

      detalhes.push(resultado);

      if (resultado.gerarAviso) {
        avisos.push(
          `${fonte.nomeFolha}, linha ${i + 1}, ${membro.nome}: ${resultado.motivo}`
        );
      }
    }
  }
}

function avaliarRegistoPresenca_(dados) {
  const respostaNormalizada = normalizarTexto_(dados.respostaOriginal);
  const presente = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.presente
  );
  const atraso = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.atraso
  );
  const falta = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.falta
  );
  const tipoConcerto = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.tipoConcerto
  );
  const tipoEnsaio = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.tipoEnsaio
  );

  let entraCalculo = false;
  let entraNumerador = false;
  let entraDenominador = false;
  let pontosObtidos = 0;
  let pontosMaximos = 0;
  let contaComoAtividade = false;
  let motivo = "";
  let gerarAviso = false;

  if (!dados.dataValida) {
    motivo = "Data inválida: registo ignorado";
    gerarAviso = true;
  } else if (dados.tipoNormalizado === tipoConcerto) {
    if (
      compararApenasData_(
        dados.dataAtividade,
        ACTIVITIES_START_DATE
      ) < 0
    ) {
      motivo = "Concerto anterior a 01/05/2026";
    } else if (respostaNormalizada === presente) {
      entraCalculo = true;
      contaComoAtividade = true;
      motivo = "Concerto presente: conta como atividade";
    } else if (
      !respostaNormalizada ||
      respostaNormalizada === falta
    ) {
      entraCalculo = true;
      motivo = "Concerto sem presença: não conta como atividade";
    } else {
      entraCalculo = true;
      motivo = `Resposta inesperada em concerto: "${dados.respostaOriginal}"`;
      gerarAviso = true;
    }
  } else if (dados.tipoNormalizado === tipoEnsaio) {
    const entradaValida =
      dados.membro.entrada instanceof Date &&
      !isNaN(dados.membro.entrada.getTime());

    if (!entradaValida) {
      motivo = "Data de entrada inválida: ensaio ignorado";
      gerarAviso = true;
    } else if (
      compararApenasData_(
        dados.dataAtividade,
        dados.membro.entrada
      ) < 0
    ) {
      motivo = "Ensaio anterior à entrada";
    } else {
      entraCalculo = true;
      entraNumerador = true;
      entraDenominador = true;
      pontosMaximos = 5;

      if (respostaNormalizada === presente) {
        pontosObtidos = 5;
        motivo = "Presente";
      } else if (respostaNormalizada === atraso) {
        pontosObtidos = 4;
        motivo = "Atraso";
      } else if (!respostaNormalizada) {
        motivo = "Resposta vazia: conta como 0";
      } else if (respostaNormalizada === falta) {
        motivo = "Falta";
      } else {
        motivo = `Resposta inesperada: "${dados.respostaOriginal}", conta como 0`;
        gerarAviso = true;
      }
    }
  } else {
    motivo = `Tipo de atividade inesperado: "${dados.tipoRegisto}"`;
    gerarAviso = true;
  }

  return {
    folha: dados.fonte,
    linha: dados.linha,
    data: dados.dataValida ? dados.dataAtividade : null,
    tipoRegisto: dados.tipoRegisto,
    ehEnsaio: dados.tipoNormalizado === tipoEnsaio,
    nomeMembro: dados.membro.nome,
    chaveNome: dados.membro.chaveNome,
    resposta: dados.respostaOriginal,
    dataEntrada: dados.membro.entrada,
    entraCalculo,
    entraNumerador,
    entraDenominador,
    pontosObtidos,
    pontosMaximos,
    contaComoAtividade,
    motivo,
    gerarAviso
  };
}

/**
 * Compara duas datas ignorando horas e minutos.
 *
 * Resultado:
 * - negativo: dataA anterior a dataB
 * - zero: mesmo dia
 * - positivo: dataA posterior a dataB
 */
function compararApenasData_(dataA, dataB) {
  const a = new Date(
    dataA.getFullYear(),
    dataA.getMonth(),
    dataA.getDate()
  );

  const b = new Date(
    dataB.getFullYear(),
    dataB.getMonth(),
    dataB.getDate()
  );

  return a.getTime() - b.getTime();
}

/**
 * Extrai o nome de:
 *
 * Registo de Presenças [Inês Regina]
 */
function extrairNomeCabecalhoPresenca_(cabecalho) {
  const texto = String(cabecalho ?? "").trim();

  const correspondencia = texto.match(
    /^Registo\s+de\s+Presenças\s*\[(.+?)\]\s*$/i
  );

  if (correspondencia) {
    return correspondencia[1].trim();
  }

  return texto || null;
}


/**
 * ============================================================
 * VALORES MANUAIS EXISTENTES
 * ============================================================
 */

/**
 * Lê os valores existentes de:
 *
 * - Valor Individual
 * - Valor Apoios
 *
 * Estes valores são associados ao nome do membro, para não
 * serem apagados quando a folha Distribuição é regenerada.
 */
function lerValoresManuaisExistentes_(spreadsheet) {
  const nomeFolha =
    DISTRIBUICAO_CONFIG.folhas.distribuicao;

  const folha = spreadsheet.getSheetByName(nomeFolha);

  if (!folha) {
    return {};
  }

  const ultimaLinha = folha.getLastRow();
  const ultimaColuna = folha.getLastColumn();

  if (ultimaLinha < 2 || ultimaColuna < 1) {
    return {};
  }

  const dados = folha
    .getRange(1, 1, ultimaLinha, ultimaColuna)
    .getValues();

  const cabecalhos = dados[0];

  const indiceNome = encontrarIndiceCabecalho_(
    cabecalhos,
    [
      DISTRIBUICAO_CONFIG.cabecalhos.distribuicao.nome
    ],
    false
  );

  const indiceIndividual =
    encontrarIndiceCabecalho_(
      cabecalhos,
      [
        DISTRIBUICAO_CONFIG.cabecalhos.distribuicao
          .valorIndividual
      ],
      false
    );

  const indiceApoios =
    encontrarIndiceCabecalho_(
      cabecalhos,
      [
        DISTRIBUICAO_CONFIG.cabecalhos.distribuicao
          .valorApoios
      ],
      false
    );

  if (
    indiceNome < 0 ||
    indiceIndividual < 0 ||
    indiceApoios < 0
  ) {
    return {};
  }

  const valores = {};

  for (let i = 1; i < dados.length; i++) {
    const nome = String(
      dados[i][indiceNome] ?? ""
    ).trim();

    if (!nome) {
      continue;
    }

    valores[normalizarNome_(nome)] = {
      valorIndividual:
        dados[i][indiceIndividual],
      valorApoios:
        dados[i][indiceApoios]
    };
  }

  return valores;
}


/**
 * ============================================================
 * ESCRITA DA DISTRIBUIÇÃO
 * ============================================================
 */

function escreverDistribuicao_(
  spreadsheet,
  linhas,
  resumo
) {
  const nomeFolha =
    DISTRIBUICAO_CONFIG.folhas.distribuicao;

  let folha = spreadsheet.getSheetByName(nomeFolha);

  if (!folha) {
    folha = spreadsheet.insertSheet(nomeFolha);
  }

  const filtroExistente = folha.getFilter();

  if (filtroExistente) {
    filtroExistente.remove();
  }

  folha.clear();

  const h =
    DISTRIBUICAO_CONFIG.cabecalhos.distribuicao;

  const cabecalhos = [
    h.ordem,
    h.nome,
    h.atividades,
    h.assiduidade,
    h.pontos,
    h.valorFundoComum,
    h.valorIndividual,
    h.valorApoios,
    h.valorFinal
  ];

  folha
    .getRange(1, 1, 1, cabecalhos.length)
    .setValues([cabecalhos]);

  if (linhas.length > 0) {
    const valores = linhas.map(linha => [
      linha.ordem,
      linha.nome,
      linha.atividades,
      linha.assiduidade,
      linha.pontos,
      linha.valorFundoComum,
      linha.valorIndividual,
      linha.valorApoios,
      ""
    ]);

    folha
      .getRange(
        2,
        1,
        valores.length,
        valores[0].length
      )
      .setValues(valores);

    const formulasValorFinal =
      linhas.map((_, indice) => {
        const linhaFolha = indice + 2;

        return [
          `=SUM(F${linhaFolha}:H${linhaFolha})`
        ];
      });

    folha
      .getRange(
        2,
        9,
        formulasValorFinal.length,
        1
      )
      .setFormulas(formulasValorFinal);
  }

  const linhaTotal = linhas.length + 2;

  folha
    .getRange(linhaTotal, 1, 1, 9)
    .setValues([[
      "",
      "TOTAL",
      "",
      "",
      "",
      "",
      "",
      "",
      ""
    ]]);

  if (linhas.length > 0) {
    folha
      .getRange(linhaTotal, 3)
      .setFormula(
        `=SUM(C2:C${linhaTotal - 1})`
      );

    folha
      .getRange(linhaTotal, 5)
      .setFormula(
        `=SUM(E2:E${linhaTotal - 1})`
      );

    folha
      .getRange(linhaTotal, 6)
      .setFormula(
        `=SUM(F2:F${linhaTotal - 1})`
      );

    folha
      .getRange(linhaTotal, 7)
      .setFormula(
        `=SUM(G2:G${linhaTotal - 1})`
      );

    folha
      .getRange(linhaTotal, 8)
      .setFormula(
        `=SUM(H2:H${linhaTotal - 1})`
      );

    folha
      .getRange(linhaTotal, 9)
      .setFormula(
        `=SUM(I2:I${linhaTotal - 1})`
      );
  }

  aplicarFormatacaoDistribuicao_(
    folha,
    linhas.length,
    linhaTotal
  );

  escreverResumoTecnico_(
    folha,
    resumo,
    linhaTotal
  );
}


/**
 * ============================================================
 * FORMATAÇÃO
 * ============================================================
 */

function aplicarFormatacaoDistribuicao_(
  folha,
  numeroMembros,
  linhaTotal
) {
  const numeroLinhasDados =
    Math.max(numeroMembros, 1);

  folha.setFrozenRows(1);

  folha
    .getRange(1, 1, 1, 9)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  folha.setRowHeight(1, 42);

  folha
    .getRange(linhaTotal, 1, 1, 9)
    .setFontWeight("bold");

  folha
    .getRange(2, 1, numeroLinhasDados, 1)
    .setHorizontalAlignment("center");

  folha
    .getRange(2, 3, numeroLinhasDados, 3)
    .setHorizontalAlignment("center");

  if (numeroMembros > 0) {
    // Atividades
    folha
      .getRange(2, 3, numeroMembros, 1)
      .setNumberFormat("0");

    // Assiduidade
    folha
      .getRange(2, 4, numeroMembros, 1)
      .setNumberFormat("0.0%");

    // Pontos finais
    folha
      .getRange(2, 5, numeroMembros, 1)
      .setNumberFormat("0.000");

    // Fundo comum até valor final
    folha
      .getRange(2, 6, numeroMembros, 4)
      .setNumberFormat('#,##0.00 "€"');
  }

  folha
    .getRange(linhaTotal, 3)
    .setNumberFormat("0");

  folha
    .getRange(linhaTotal, 5)
    .setNumberFormat("0.000");

  folha
    .getRange(linhaTotal, 6, 1, 4)
    .setNumberFormat('#,##0.00 "€"');

  folha.autoResizeColumns(1, 9);

  folha.setColumnWidth(1, 70);
  folha.setColumnWidth(2, 190);
  folha.setColumnWidth(3, 90);
  folha.setColumnWidth(4, 105);
  folha.setColumnWidth(5, 90);
  folha.setColumnWidth(6, 150);
  folha.setColumnWidth(7, 135);
  folha.setColumnWidth(8, 120);
  folha.setColumnWidth(9, 120);

  folha
    .getRange(1, 1, linhaTotal, 9)
    .setVerticalAlignment("middle");

  folha
    .getRange(1, 1, linhaTotal, 9)
    .createFilter();
}

/**
 * Escreve alguns dados de controlo abaixo da tabela.
 */
function escreverResumoTecnico_(
  folha,
  resumo,
  linhaTotal
) {
  const linhaInicio = linhaTotal + 3;

  const dados = [
    ["Controlo", "Valor"],
    [
      "Saldo Fundo Comum",
      resumo.saldoFundoComum
    ],
    [
      "Total de pontos",
      resumo.totalPontos
    ],
    [
      "Fontes de naipe omitidas",
      resumo.fontesOmitidas.length > 0
        ? resumo.fontesOmitidas.join(", ")
        : "Nenhuma"
    ]
  ];

  folha
    .getRange(
      linhaInicio,
      1,
      dados.length,
      dados[0].length
    )
    .setValues(dados);

  folha
    .getRange(linhaInicio, 1, 1, 2)
    .setFontWeight("bold");

  folha
    .getRange(linhaInicio + 1, 2)
    .setNumberFormat('#,##0.00 "€"');
}


/**
 * ============================================================
 * MENU
 * ============================================================
 */

function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu("Coro")
    .addItem(
      "Atualizar distribuição",
      "gerarDistribuicao"
    )
    .addItem(
      "Gerar mapa mensal",
      "gerarMapaAssiduidadeMensal"
    )
    .addSeparator()
    .addItem(
      "Publicar distribuição",
      "publicarDistribuicao"
    )
    .addToUi();
}


/**
 * ============================================================
 * FUNÇÕES AUXILIARES
 * ============================================================
 */

function obterFolhaObrigatoria_(
  spreadsheet,
  nomeFolha
) {
  const folha = spreadsheet.getSheetByName(nomeFolha);

  if (!folha) {
    throw new Error(
      `Não foi encontrada a folha "${nomeFolha}".`
    );
  }

  return folha;
}


/**
 * Encontra uma coluna através do cabeçalho.
 *
 * Devolve índice baseado em zero.
 */
function encontrarIndiceCabecalho_(
  cabecalhos,
  alternativas,
  obrigatorio,
  nomeFolha = ""
) {
  const cabecalhosNormalizados =
    cabecalhos.map(normalizarTexto_);

  for (const alternativa of alternativas) {
    const indice = cabecalhosNormalizados.indexOf(
      normalizarTexto_(alternativa)
    );

    if (indice >= 0) {
      return indice;
    }
  }

  if (obrigatorio) {
    const contexto = nomeFolha
      ? ` na folha "${nomeFolha}"`
      : "";

    throw new Error(
      `Não foi encontrado o cabeçalho ${
        alternativas
          .map(valor => `"${valor}"`)
          .join(" ou ")
      }${contexto}.`
    );
  }

  return -1;
}


/**
 * Normalização genérica:
 *
 * - remove espaços adicionais;
 * - converte para minúsculas;
 * - remove acentos.
 */
function normalizarTexto_(valor) {
  return String(valor ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}


/**
 * Normalização usada para associar nomes entre folhas.
 */
function normalizarNome_(nome) {
  return String(nome ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleLowerCase("pt-PT");
}


/**
 * Aceita números ou texto, incluindo:
 *
 * 1200
 * "1 200"
 * "1.200,50"
 * "1200,50 €"
 */
function converterNumero_(valor) {
  if (
    valor === null ||
    valor === undefined ||
    valor === ""
  ) {
    return 0;
  }

  if (
    typeof valor === "number" &&
    Number.isFinite(valor)
  ) {
    return valor;
  }

  let texto = String(valor)
    .trim()
    .replace(/\s/g, "")
    .replace(/€/g, "");

  if (!texto) {
    return 0;
  }

  const temVirgula = texto.includes(",");
  const temPonto = texto.includes(".");

  if (temVirgula && temPonto) {
    texto = texto
      .replace(/\./g, "")
      .replace(",", ".");
  } else if (temVirgula) {
    texto = texto.replace(",", ".");
  }

  const numero = Number(texto);

  return Number.isFinite(numero)
    ? numero
    : 0;
}


/**
 * Converte valores de data vindos da folha.
 */
function converterData_(valor) {
  if (
    valor instanceof Date &&
    !isNaN(valor.getTime())
  ) {
    return valor;
  }

  if (
    valor === null ||
    valor === undefined ||
    valor === ""
  ) {
    return null;
  }

  const texto = String(valor).trim();

  const formatoPortugues = texto.match(
    /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/
  );

  if (formatoPortugues) {
    const dia = Number(formatoPortugues[1]);
    const mes = Number(formatoPortugues[2]) - 1;
    const ano = Number(formatoPortugues[3]);

    const data = new Date(ano, mes, dia);

    return isNaN(data.getTime())
      ? null
      : data;
  }

  const data = new Date(valor);

  return isNaN(data.getTime())
    ? null
    : data;
}


/**
 * Determina se todas as células de uma linha estão vazias.
 */
function linhaCompletamenteVazia_(linha) {
  return linha.every(valor =>
    valor === "" ||
    valor === null ||
    valor === undefined
  );
}


function formatarEuro_(valor) {
  return Number(valor || 0).toLocaleString(
    "pt-PT",
    {
      style: "currency",
      currency: "EUR"
    }
  );
}
