/**
 * ============================================================
 * CONFIGURAÇÃO
 * ============================================================
 */

const DISTRIBUICAO_CONFIG = {
  folhas: {
    membros: "Membros",
    movimentosPreWCG: "Movimento_PREWCG",
    movimentos: "Movimentos",
    respostas: "Form_Responses",
    distribuicao: "Distribuição"
  },

  /**
   * Membros com data de entrada anterior a esta data
   * recebem a divisão da caixa pré-WCG.
   *
   * new Date(ano, mês-1, dia)
   * Maio = 4 porque os meses começam em zero.
   */
  dataLimitePreWCG: new Date(2026, 4, 1),

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

    respostas: {
      tipoAtividade: [
        "Tipo de Atividade",
        "Tipo de atividade",
        "Tipo da Atividade"
      ],

      dataAtividade: [
        "Data ",
        "Data da Atividade",
        "Data de Atividade",
        "Data do Ensaio:"
      ]
    },

    distribuicao: {
      ordem: "Ordem",
      nome: "Nome",
      valorPreWCG: "Valor Caixa Antes WCG",
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

  const saldoPreWCG = calcularSaldoMovimentos_(
    spreadsheet,
    DISTRIBUICAO_CONFIG.folhas.movimentosPreWCG
  );

  const saldoFundoComum = calcularSaldoMovimentos_(
    spreadsheet,
    DISTRIBUICAO_CONFIG.folhas.movimentos
  );

  const membrosPreWCG = membros.filter(membro =>
    membro.entrada instanceof Date &&
    !isNaN(membro.entrada.getTime()) &&
    membro.entrada < DISTRIBUICAO_CONFIG.dataLimitePreWCG
  );

  const valorPreWCGPorMembro =
    membrosPreWCG.length > 0
      ? saldoPreWCG / membrosPreWCG.length
      : 0;

  const chavesMembrosPreWCG = new Set(
    membrosPreWCG.map(membro => membro.chaveNome)
  );

const metricasPorMembro =
  calcularMetricasParticipacao_(
    spreadsheet,
    membros
  );

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

    const valorPreWCG =
      chavesMembrosPreWCG.has(membro.chaveNome)
        ? valorPreWCGPorMembro
        : 0;

    const valorFundoComum =
      totalPontos > 0
        ? pontos * saldoFundoComum / totalPontos
        : 0;

    const valoresExistentes =
      valoresManuais[membro.chaveNome] || {};

    return {
      ordem: membro.ordem,
      nome: membro.nome,
      valorPreWCG: valorPreWCG,
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
      saldoPreWCG,
      saldoFundoComum,
      totalPontos,
      numeroMembrosPreWCG: membrosPreWCG.length
    }
  );

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    "Distribuição atualizada",
    [
      `Membros ativos: ${membros.length}`,
      `Membros elegíveis pré-WCG: ${membrosPreWCG.length}`,
      `Total de pontos: ${totalPontos}`,
      `Caixa pré-WCG: ${formatarEuro_(saldoPreWCG)}`,
      `Fundo comum: ${formatarEuro_(saldoFundoComum)}`
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
  const nomeFolha =
    DISTRIBUICAO_CONFIG.folhas.respostas;

  const folha = obterFolhaObrigatoria_(
    spreadsheet,
    nomeFolha
  );

  const ultimaLinha = folha.getLastRow();
  const ultimaColuna = folha.getLastColumn();

  const metricas = {};

  membros.forEach(membro => {
    metricas[membro.chaveNome] = {
      atividades: 0,
      pontosEnsaios: 0,
      ensaiosPossiveis: 0,
      assiduidade: 0,
      pontos: 0
    };
  });

  if (ultimaLinha < 2 || ultimaColuna < 1) {
    return metricas;
  }

  const dados = folha
    .getRange(1, 1, ultimaLinha, ultimaColuna)
    .getValues();

  const cabecalhos = dados[0];

  const indiceDataAtividade =
    encontrarIndiceCabecalho_(
      cabecalhos,
      DISTRIBUICAO_CONFIG.cabecalhos.respostas
        .dataAtividade,
      true,
      nomeFolha
    );

  const indiceTipoAtividade =
    encontrarIndiceCabecalho_(
      cabecalhos,
      DISTRIBUICAO_CONFIG.cabecalhos.respostas
        .tipoAtividade,
      true,
      nomeFolha
    );

  const membrosPorChave = {};

  membros.forEach(membro => {
    membrosPorChave[membro.chaveNome] = membro;
  });

  const colunasPresencas = [];

  cabecalhos.forEach((cabecalho, indiceColuna) => {
    const nomeExtraido =
      extrairNomeCabecalhoPresenca_(cabecalho);

    if (!nomeExtraido) {
      return;
    }

    const chaveNome = normalizarNome_(nomeExtraido);

    if (!membrosPorChave[chaveNome]) {
      return;
    }

    colunasPresencas.push({
      indiceColuna,
      chaveNome
    });
  });

  if (colunasPresencas.length === 0) {
    throw new Error(
      [
        `Não foram encontradas colunas de presenças válidas em "${nomeFolha}".`,
        "Formato esperado:",
        '"Registo de Presenças [Nome do Membro]"'
      ].join("\n")
    );
  }

  const tipoConcerto = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.tipoConcerto
  );

  const tipoEnsaio = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.tipoEnsaio
  );

  const presente = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.presente
  );

  const atraso = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.atraso
  );

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];

    if (linhaCompletamenteVazia_(linha)) {
      continue;
    }

    const dataAtividade =
      converterData_(linha[indiceDataAtividade]);

    if (
      !(dataAtividade instanceof Date) ||
      isNaN(dataAtividade.getTime())
    ) {
      continue;
    }

    const tipoAtividade =
      normalizarTexto_(
        linha[indiceTipoAtividade]
      );

    for (const coluna of colunasPresencas) {
      const membro =
        membrosPorChave[coluna.chaveNome];

      if (!membro) {
        continue;
      }

      /**
       * O membro só pode ser contabilizado em atividades
       * ocorridas na data de entrada ou depois dela.
       */
      if (
        membro.entrada instanceof Date &&
        !isNaN(membro.entrada.getTime()) &&
        compararApenasData_(
          dataAtividade,
          membro.entrada
        ) < 0
      ) {
        continue;
      }

      const resposta =
        normalizarTexto_(
          linha[coluna.indiceColuna]
        );

      /**
       * Concertos:
       * Presente = 1 atividade.
       */
      if (tipoAtividade === tipoConcerto) {
        if (resposta === presente) {
          metricas[coluna.chaveNome].atividades += 1;
        }

        continue;
      }

      /**
       * Ensaios:
       * todos os ensaios posteriores à entrada entram no
       * denominador, independentemente da resposta.
       */
      if (tipoAtividade === tipoEnsaio) {
        metricas[coluna.chaveNome].ensaiosPossiveis += 1;

        if (resposta === presente) {
          metricas[coluna.chaveNome].pontosEnsaios += 5;
        } else if (resposta === atraso) {
          metricas[coluna.chaveNome].pontosEnsaios += 4;
        }
      }
    }
  }

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

  return metricas;
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

  if (!correspondencia) {
    return null;
  }

  return correspondencia[1].trim();
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
    h.valorPreWCG,
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
      linha.valorPreWCG,
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
          `=SUM(C${linhaFolha};G${linhaFolha}:I${linhaFolha})`
        ];
      });

    folha
      .getRange(
        2,
        10,
        formulasValorFinal.length,
        1
      )
      .setFormulas(formulasValorFinal);
  }

  const linhaTotal = linhas.length + 2;

  folha
    .getRange(linhaTotal, 1, 1, 10)
    .setValues([[
      "",
      "TOTAL",
      "",
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
      .getRange(linhaTotal, 4)
      .setFormula(
        `=SUM(D2:D${linhaTotal - 1})`
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

    folha
      .getRange(linhaTotal, 10)
      .setFormula(
        `=SUM(J2:J${linhaTotal - 1})`
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
    .getRange(1, 1, 1, 10)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  folha.setRowHeight(1, 42);

  folha
    .getRange(linhaTotal, 1, 1, 10)
    .setFontWeight("bold");

  folha
    .getRange(2, 1, numeroLinhasDados, 1)
    .setHorizontalAlignment("center");

  folha
    .getRange(2, 4, numeroLinhasDados, 3)
    .setHorizontalAlignment("center");

  if (numeroMembros > 0) {
    // Valor Caixa Antes WCG
    folha
      .getRange(2, 3, numeroMembros, 1)
      .setNumberFormat('#,##0.00 "€"');

    // Atividades
    folha
      .getRange(2, 4, numeroMembros, 1)
      .setNumberFormat("0");

    // Assiduidade
    folha
      .getRange(2, 5, numeroMembros, 1)
      .setNumberFormat("0.0%");

    // Pontos finais
    folha
      .getRange(2, 6, numeroMembros, 1)
      .setNumberFormat("0.000");

    // Fundo comum até valor final
    folha
      .getRange(2, 7, numeroMembros, 4)
      .setNumberFormat('#,##0.00 "€"');
  }

  folha
    .getRange(linhaTotal, 3)
    .setNumberFormat('#,##0.00 "€"');

  folha
    .getRange(linhaTotal, 4)
    .setNumberFormat("0");

  folha
    .getRange(linhaTotal, 6)
    .setNumberFormat("0.000");

  folha
    .getRange(linhaTotal, 7, 1, 4)
    .setNumberFormat('#,##0.00 "€"');

  folha.autoResizeColumns(1, 10);

  folha.setColumnWidth(1, 70);
  folha.setColumnWidth(2, 190);
  folha.setColumnWidth(3, 155);
  folha.setColumnWidth(4, 90);
  folha.setColumnWidth(5, 105);
  folha.setColumnWidth(6, 90);
  folha.setColumnWidth(7, 150);
  folha.setColumnWidth(8, 135);
  folha.setColumnWidth(9, 120);
  folha.setColumnWidth(10, 120);

  folha
    .getRange(1, 1, linhaTotal, 10)
    .setVerticalAlignment("middle");

  folha
    .getRange(1, 1, linhaTotal, 10)
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
      "Saldo Caixa Antes WCG",
      resumo.saldoPreWCG
    ],
    [
      "Número de membros pré-WCG",
      resumo.numeroMembrosPreWCG
    ],
    [
      "Saldo Fundo Comum",
      resumo.saldoFundoComum
    ],
    [
      "Total de pontos",
      resumo.totalPontos
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

  folha
    .getRange(linhaInicio + 3, 2)
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
  return normalizarTexto_(nome);
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