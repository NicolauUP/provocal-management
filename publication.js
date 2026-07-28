/**
 * Publica manualmente uma cópia de consulta da distribuição e
 * dos mapas mensais já gerados.
 *
 * A publicação:
 * - só pode ser iniciada por um email autorizado;
 * - exige confirmação;
 * - copia valores calculados, nunca fórmulas;
 * - não copia as folhas privadas nem o resumo técnico interno.
 */
function publicarDistribuicao() {
  const emailPublicador = obterEmailPublicador_();

  if (!utilizadorPodePublicar_(emailPublicador)) {
    throw new Error(
      `O utilizador "${emailPublicador || "não identificado"}" não está autorizado a publicar.`
    );
  }

  const spreadsheetOrigem =
    SpreadsheetApp.getActiveSpreadsheet();
  const folhasOrigem =
    listarFolhasPublicaveis_(spreadsheetOrigem);
  const ui = SpreadsheetApp.getUi();
  const confirmacao = ui.alert(
    "Publicar distribuição",
    [
      "Serão substituídas no ficheiro público:",
      ...folhasOrigem.map(item => `• ${item.nome}`),
      "",
      "Serão copiados apenas valores e formatação.",
      "Queres continuar?"
    ].join("\n"),
    ui.ButtonSet.YES_NO
  );

  if (confirmacao !== ui.Button.YES) {
    return;
  }

  const spreadsheetDestino =
    SpreadsheetApp.openById(PUBLICATION.SPREADSHEET_ID);

  if (spreadsheetDestino.getId() === spreadsheetOrigem.getId()) {
    throw new Error(
      "A spreadsheet pública não pode ser a mesma spreadsheet privada."
    );
  }

  folhasOrigem.forEach(item => {
    publicarFolha_(
      item.folha,
      spreadsheetDestino,
      item.nome,
      item.tipo
    );
  });

  publicarDadosConsulta_(
    spreadsheetOrigem,
    spreadsheetDestino
  );
  publicarDadosPortalPrivado_(spreadsheetOrigem);

  escreverInformacaoPublicacao_(
    spreadsheetDestino,
    folhasOrigem.map(item => item.nome)
  );

  SpreadsheetApp.flush();

  ui.alert(
    "Publicação concluída",
    [
      `Folhas publicadas: ${folhasOrigem.length}`,
      `Destino: ${spreadsheetDestino.getName()}`,
      "Os dados foram copiados como valores, sem fórmulas.",
      "O resumo individual do portal também foi atualizado."
    ].join("\n"),
    ui.ButtonSet.OK
  );
}

/**
 * Publica um snapshot mínimo e privado para o projeto Apps Script do portal.
 * Só membros ativos com Email preenchido entram neste ficheiro. Não é uma
 * folha pública e não deve ser partilhada com membros do coro.
 */
function publicarDadosPortalPrivado_(spreadsheetOrigem) {
  const membros = lerMembrosAtivos_(spreadsheetOrigem);
  const membrosComEmail = validarMembrosPortalPrivado_(membros);
  const analise = analisarParticipacao_(
    spreadsheetOrigem,
    membrosComEmail
  );
  const dados = criarDadosPortalPrivado_(
    membrosComEmail,
    analise.detalhes
  );
  const spreadsheetDestino = SpreadsheetApp.openById(
    PUBLICATION.PRIVATE_PORTAL_SPREADSHEET_ID
  );
  let folha = spreadsheetDestino.getSheetByName(
    PUBLICATION.PRIVATE_PORTAL_DATA_SHEET
  );

  if (!folha) {
    folha = spreadsheetDestino.insertSheet(
      PUBLICATION.PRIVATE_PORTAL_DATA_SHEET
    );
  }

  prepararFolhaDestino_(
    folha,
    Math.max(dados.length, 1),
    dados[0].length
  );
  folha
    .getRange(1, 1, dados.length, dados[0].length)
    .setValues(dados);
  folha
    .getRange(1, 1, 1, dados[0].length)
    .setFontWeight("bold")
    .setBackground("#1F4E78")
    .setFontColor("#FFFFFF");
  folha.getRange(2, 3, Math.max(dados.length - 1, 1), 1)
    .setNumberFormat("dd/MM/yyyy");
  folha.getRange(2, 4, Math.max(dados.length - 1, 1), 1)
    .setNumberFormat("@");
  folha.setFrozenRows(1);
  folha.autoResizeColumns(1, dados[0].length);

  publicarValoresPortalPrivados_(
    spreadsheetOrigem,
    spreadsheetDestino,
    membrosComEmail
  );
  publicarResumoGlobalPortalPrivado_(
    spreadsheetOrigem,
    spreadsheetDestino
  );
}

function publicarValoresPortalPrivados_(
  spreadsheetOrigem,
  spreadsheetDestino,
  membros
) {
  const folhaOrigem = obterFolhaObrigatoria_(
    spreadsheetOrigem,
    SHEETS.DISTRIBUTION
  );
  const dadosOrigem = folhaOrigem.getDataRange().getValues();
  const cabecalhos = dadosOrigem[0];
  const indiceNome = encontrarIndiceCabecalho_(
    cabecalhos,
    [DISTRIBUICAO_CONFIG.cabecalhos.distribuicao.nome],
    true,
    folhaOrigem.getName()
  );
  const campos = [
    "pontos",
    "valorFundoComum",
    "valorIndividual",
    "valorApoios",
    "valorFinal"
  ];
  const indices = {};

  campos.forEach(campo => {
    indices[campo] = encontrarIndiceCabecalho_(
      cabecalhos,
      [DISTRIBUICAO_CONFIG.cabecalhos.distribuicao[campo]],
      true,
      folhaOrigem.getName()
    );
  });

  const valoresPorMembro = {};
  dadosOrigem.slice(1).forEach(linha => {
    valoresPorMembro[normalizarNome_(linha[indiceNome])] = linha;
  });

  const dadosDestino = [[
    "Email",
    "Nome",
    "Pontos Finais",
    "Valor Fundo Comum",
    "Valor Individual",
    "Valor Apoios",
    "Valor Final"
  ], ...membros.map(membro => {
    const linha = valoresPorMembro[membro.chaveNome] || [];
    return [
      membro.email,
      membro.nome,
      Number(linha[indices.pontos]) || 0,
      Number(linha[indices.valorFundoComum]) || 0,
      Number(linha[indices.valorIndividual]) || 0,
      Number(linha[indices.valorApoios]) || 0,
      Number(linha[indices.valorFinal]) || 0
    ];
  })];
  let folhaDestino = spreadsheetDestino.getSheetByName(
    PUBLICATION.PRIVATE_PORTAL_VALUES_SHEET
  );

  if (!folhaDestino) {
    folhaDestino = spreadsheetDestino.insertSheet(
      PUBLICATION.PRIVATE_PORTAL_VALUES_SHEET
    );
  }

  prepararFolhaDestino_(
    folhaDestino,
    dadosDestino.length,
    dadosDestino[0].length
  );
  folhaDestino.getRange(1, 1, dadosDestino.length, dadosDestino[0].length)
    .setValues(dadosDestino);
  folhaDestino.getRange(1, 1, 1, dadosDestino[0].length)
    .setFontWeight("bold")
    .setBackground("#1F4E78")
    .setFontColor("#FFFFFF");
  folhaDestino.getRange(2, 4, Math.max(dadosDestino.length - 1, 1), 4)
    .setNumberFormat("#,##0.00 €");
  folhaDestino.setFrozenRows(1);
  folhaDestino.autoResizeColumns(1, dadosDestino[0].length);
}

function publicarResumoGlobalPortalPrivado_(origem, destino) {
  const folha = obterFolhaObrigatoria_(origem, SHEETS.DISTRIBUTION);
  const dados = folha.getDataRange().getValues();
  const cabecalhos = dados[0];
  const indiceNome = encontrarIndiceCabecalho_(cabecalhos, ["Nome"], true, folha.getName());
  const indicePontos = encontrarIndiceCabecalho_(cabecalhos, ["Pontos"], true, folha.getName());
  const indiceFundo = encontrarIndiceCabecalho_(cabecalhos, ["Valor Fundo Comum"], true, folha.getName());
  const total = dados.slice(1).find(linha => normalizarTexto_(linha[indiceNome]) === "total") || [];
  const pontos = Number(total[indicePontos]) || 0;
  const fundo = Number(total[indiceFundo]) || 0;
  let destinoFolha = destino.getSheetByName(PUBLICATION.PRIVATE_PORTAL_GLOBAL_SHEET);
  if (!destinoFolha) destinoFolha = destino.insertSheet(PUBLICATION.PRIVATE_PORTAL_GLOBAL_SHEET);
  const resumo = [["Fundo comum distribuível", "Valor por ponto", "Pontos finais do coro"], [fundo, pontos ? fundo / pontos : 0, pontos]];
  prepararFolhaDestino_(destinoFolha, 2, 3);
  destinoFolha.getRange(1, 1, 2, 3).setValues(resumo);
  destinoFolha.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1F4E78").setFontColor("#FFFFFF");
  destinoFolha.getRange(2, 1, 1, 2).setNumberFormat("#,##0.00 €");
  destinoFolha.setFrozenRows(1);
}

function validarMembrosPortalPrivado_(membros) {
  const emails = {};

  return membros.filter(membro => {
    const email = String(membro.email || "").trim().toLowerCase();

    if (!email) {
      return false;
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      throw new Error(
        `O Email do membro "${membro.nome}" não é válido.`
      );
    }

    if (emails[email]) {
      throw new Error(
        `O Email "${email}" está repetido em Membros.`
      );
    }

    emails[email] = true;
    return true;
  });
}

function criarDadosPortalPrivado_(membros, detalhes) {
  const cabecalhos = [
    "Email",
    "Nome",
    "Data",
    "Mês",
    "Origem",
    "Tipo",
    "Nome da Atividade",
    "Resposta",
    "Pontos obtidos",
    "Pontos máximos",
    "Conta como atividade"
  ];
  const emailPorMembro = {};
  const tipoConcerto = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.tipoConcerto
  );

  membros.forEach(membro => {
    emailPorMembro[membro.chaveNome] = membro.email;
  });

  const linhas = detalhes
    .filter(detalhe => {
      if (!(detalhe.data instanceof Date)) {
        return false;
      }

      if (compararApenasData_(detalhe.data, ACTIVITIES_START_DATE) < 0) {
        return false;
      }

      if (detalhe.ehEnsaio) {
        return detalhe.entraDenominador;
      }

      return (
        normalizarTexto_(detalhe.tipoRegisto) === tipoConcerto &&
        detalhe.entraCalculo
      );
    })
    .sort((a, b) =>
      a.nomeMembro.localeCompare(b.nomeMembro, "pt") ||
      a.data.getTime() - b.data.getTime()
    )
    .map(detalhe => [
      emailPorMembro[detalhe.chaveNome],
      detalhe.nomeMembro,
      detalhe.data,
      criarChaveMes_(detalhe.data),
      obterNomeOrigemPublica_(detalhe.folha),
      detalhe.ehEnsaio
        ? detalhe.tipoRegisto
        : DISTRIBUICAO_CONFIG.valores.tipoConcerto,
      detalhe.nomeConcerto || "",
      detalhe.resposta || "Sem resposta",
      detalhe.pontosObtidos,
      detalhe.pontosMaximos,
      detalhe.contaComoAtividade
    ]);

  return [cabecalhos, ...linhas];
}

/**
 * Cria no ficheiro público o snapshot que alimenta o portal.
 * A folha é ocultada apenas para manter a navegação limpa; não
 * é usada como mecanismo de segurança.
 */
function publicarDadosConsulta_(
  spreadsheetOrigem,
  spreadsheetDestino
) {
  const membros = lerMembrosAtivos_(spreadsheetOrigem);
  const analise = analisarParticipacao_(
    spreadsheetOrigem,
    membros
  );
  const dados = criarDadosConsultaPublica_(
    analise.detalhes
  );
  let folha = spreadsheetDestino.getSheetByName(
    PUBLICATION.DATA_SHEET
  );

  if (!folha) {
    folha = spreadsheetDestino.insertSheet(
      PUBLICATION.DATA_SHEET
    );
  }

  prepararFolhaDestino_(
    folha,
    Math.max(dados.length, 1),
    dados[0].length
  );
  folha
    .getRange(
      2,
      3,
      Math.max(dados.length - 1, 1),
      1
    )
    .setNumberFormat("@");
  folha
    .getRange(1, 1, dados.length, dados[0].length)
    .setValues(dados);
  folha
    .getRange(1, 1, 1, dados[0].length)
    .setFontWeight("bold")
    .setBackground("#1F4E78")
    .setFontColor("#FFFFFF");
  folha
    .getRange(2, 2, Math.max(dados.length - 1, 1), 1)
    .setNumberFormat("dd/MM/yyyy");
  folha.setFrozenRows(1);
  folha.autoResizeColumns(1, dados[0].length);

  if (!folha.isSheetHidden()) {
    folha.hideSheet();
  }
}

function criarDadosConsultaPublica_(detalhes) {
  const cabecalhos = [
    "Nome",
    "Data",
    "Mês",
    "Origem",
    "Tipo",
    "Resposta",
    "Pontos obtidos",
    "Pontos máximos",
    "Conta como atividade",
    "Motivo"
  ];
  const tipoConcerto = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.tipoConcerto
  );
  const detalhesPublicaveis = detalhes.filter(detalhe => {
    if (!(detalhe.data instanceof Date)) {
      return false;
    }

    if (detalhe.ehEnsaio) {
      return detalhe.entraDenominador;
    }

    return (
      normalizarTexto_(detalhe.tipoRegisto) === tipoConcerto &&
      detalhe.entraCalculo
    );
  });

  detalhesPublicaveis.sort((a, b) =>
    a.nomeMembro.localeCompare(
      b.nomeMembro,
      "pt",
      { sensitivity: "base" }
    ) ||
    a.data.getTime() - b.data.getTime() ||
    a.folha.localeCompare(b.folha, "pt")
  );

  const linhas = detalhesPublicaveis.map(detalhe => [
    detalhe.nomeMembro,
    detalhe.data,
    criarChaveMes_(detalhe.data),
    obterNomeOrigemPublica_(detalhe.folha),
    detalhe.ehEnsaio
      ? detalhe.tipoRegisto
      : DISTRIBUICAO_CONFIG.valores.tipoConcerto,
    detalhe.resposta || "Sem resposta",
    detalhe.pontosObtidos,
    detalhe.pontosMaximos,
    detalhe.contaComoAtividade,
    detalhe.motivo
  ]);

  return [cabecalhos, ...linhas];
}

function criarChaveMes_(data) {
  return [
    data.getFullYear(),
    String(data.getMonth() + 1).padStart(2, "0")
  ].join("-");
}

function obterNomeOrigemPublica_(nomeFolha) {
  if (nomeFolha === SHEETS.GENERAL_ATTENDANCE) {
    return "Geral";
  }

  return nomeFolha.replace(/^Ensaios_/, "");
}

function obterEmailPublicador_() {
  return String(
    Session.getEffectiveUser().getEmail() || ""
  )
    .trim()
    .toLowerCase();
}

function utilizadorPodePublicar_(email) {
  const emailNormalizado = String(email || "")
    .trim()
    .toLowerCase();

  return PUBLICATION.AUTHORIZED_PUBLISHERS.some(
    autorizado =>
      String(autorizado).trim().toLowerCase() ===
      emailNormalizado
  );
}

function listarFolhasPublicaveis_(spreadsheet) {
  const folhaDistribuicao = spreadsheet.getSheetByName(
    SHEETS.DISTRIBUTION
  );

  if (!folhaDistribuicao) {
    throw new Error(
      `Não existe a folha obrigatória "${SHEETS.DISTRIBUTION}".`
    );
  }

  const mapasMensais = spreadsheet
    .getSheets()
    .filter(folha =>
      /^Assiduidade_\d{4}_\d{2}$/.test(folha.getName())
    )
    .sort((a, b) =>
      a.getName().localeCompare(b.getName())
    )
    .map(folha => ({
      nome: folha.getName(),
      folha,
      tipo: "mapa"
    }));

  return [
    {
      nome: SHEETS.DISTRIBUTION,
      folha: folhaDistribuicao,
      tipo: "distribuicao"
    },
    ...mapasMensais
  ];
}

function publicarFolha_(
  folhaOrigem,
  spreadsheetDestino,
  nomeFolha,
  tipo
) {
  const intervaloOrigem = tipo === "distribuicao"
    ? obterIntervaloDistribuicaoPublicavel_(folhaOrigem)
    : obterIntervaloUtilizado_(folhaOrigem);
  let folhaDestino =
    spreadsheetDestino.getSheetByName(nomeFolha);

  if (!folhaDestino) {
    folhaDestino = spreadsheetDestino.insertSheet(nomeFolha);
  }

  prepararFolhaDestino_(
    folhaDestino,
    intervaloOrigem.getNumRows(),
    intervaloOrigem.getNumColumns()
  );
  copiarValoresEFormatacao_(
    intervaloOrigem,
    folhaDestino
  );

  if (tipo === "mapa") {
    const numeroMembros =
      Math.max(intervaloOrigem.getNumRows() - 1, 0);
    const numeroEnsaios =
      Math.max(intervaloOrigem.getNumColumns() - 3, 0);

    if (numeroMembros > 0) {
      applyConditionalFormatting(
        folhaDestino,
        numeroMembros,
        numeroEnsaios
      );
    }
  }
}

function obterIntervaloDistribuicaoPublicavel_(folha) {
  const ultimaLinha = folha.getLastRow();
  const ultimaColuna = folha.getLastColumn();

  if (ultimaLinha < 2 || ultimaColuna < 1) {
    throw new Error(
      `A folha "${folha.getName()}" não contém uma distribuição publicável.`
    );
  }

  const dados = folha
    .getRange(1, 1, ultimaLinha, ultimaColuna)
    .getValues();
  const indiceNome = encontrarIndiceCabecalho_(
    dados[0],
    [
      DISTRIBUICAO_CONFIG.cabecalhos.distribuicao.nome
    ],
    true,
    folha.getName()
  );
  const ultimaColunaPublicavel =
    encontrarUltimaColunaPreenchida_(dados[0]);
  let indiceLinhaTotal = -1;

  for (let i = 1; i < dados.length; i++) {
    if (normalizarTexto_(dados[i][indiceNome]) === "total") {
      indiceLinhaTotal = i;
      break;
    }
  }

  if (indiceLinhaTotal < 0) {
    throw new Error(
      `Não foi encontrada a linha "TOTAL" na folha "${folha.getName()}".`
    );
  }

  return folha.getRange(
    1,
    1,
    indiceLinhaTotal + 1,
    ultimaColunaPublicavel
  );
}

function obterIntervaloUtilizado_(folha) {
  const ultimaLinha = folha.getLastRow();
  const ultimaColuna = folha.getLastColumn();

  if (ultimaLinha < 1 || ultimaColuna < 1) {
    throw new Error(
      `A folha "${folha.getName()}" está vazia e não pode ser publicada.`
    );
  }

  return folha.getRange(
    1,
    1,
    ultimaLinha,
    ultimaColuna
  );
}

function encontrarUltimaColunaPreenchida_(cabecalhos) {
  for (let i = cabecalhos.length - 1; i >= 0; i--) {
    if (String(cabecalhos[i] ?? "").trim()) {
      return i + 1;
    }
  }

  throw new Error("A linha de cabeçalhos está vazia.");
}

function prepararFolhaDestino_(
  folha,
  numeroLinhas,
  numeroColunas
) {
  const filtro = folha.getFilter();

  if (filtro) {
    filtro.remove();
  }

  folha.clearConditionalFormatRules();
  folha.clear();

  if (folha.getMaxRows() < numeroLinhas) {
    folha.insertRowsAfter(
      folha.getMaxRows(),
      numeroLinhas - folha.getMaxRows()
    );
  }

  if (folha.getMaxColumns() < numeroColunas) {
    folha.insertColumnsAfter(
      folha.getMaxColumns(),
      numeroColunas - folha.getMaxColumns()
    );
  }
}

function copiarValoresEFormatacao_(intervaloOrigem, folhaDestino) {
  const numeroLinhas = intervaloOrigem.getNumRows();
  const numeroColunas = intervaloOrigem.getNumColumns();
  const intervaloDestino = folhaDestino.getRange(
    1,
    1,
    numeroLinhas,
    numeroColunas
  );

  intervaloDestino
    .setValues(intervaloOrigem.getValues())
    .setNumberFormats(intervaloOrigem.getNumberFormats())
    .setBackgrounds(intervaloOrigem.getBackgrounds())
    .setFontColors(intervaloOrigem.getFontColors())
    .setFontWeights(intervaloOrigem.getFontWeights())
    .setFontStyles(intervaloOrigem.getFontStyles())
    .setFontSizes(intervaloOrigem.getFontSizes())
    .setHorizontalAlignments(
      intervaloOrigem.getHorizontalAlignments()
    )
    .setVerticalAlignments(
      intervaloOrigem.getVerticalAlignments()
    )
    .setWraps(intervaloOrigem.getWraps());

  for (let coluna = 1; coluna <= numeroColunas; coluna++) {
    folhaDestino.setColumnWidth(
      coluna,
      intervaloOrigem.getSheet().getColumnWidth(coluna)
    );
  }

  for (let linha = 1; linha <= numeroLinhas; linha++) {
    folhaDestino.setRowHeight(
      linha,
      intervaloOrigem.getSheet().getRowHeight(linha)
    );
  }

  folhaDestino.setFrozenRows(
    intervaloOrigem.getSheet().getFrozenRows()
  );
  folhaDestino.setFrozenColumns(
    intervaloOrigem.getSheet().getFrozenColumns()
  );
  intervaloDestino.createFilter();
}

function escreverInformacaoPublicacao_(
  spreadsheet,
  nomesFolhas
) {
  let folha = spreadsheet.getSheetByName(
    PUBLICATION.INFO_SHEET
  );

  if (!folha) {
    folha = spreadsheet.insertSheet(
      PUBLICATION.INFO_SHEET
    );
  }

  const filtro = folha.getFilter();

  if (filtro) {
    filtro.remove();
  }

  folha.clear();
  folha
    .getRange(1, 1, 4, 2)
    .setValues([
      ["Informação", "Valor"],
      ["Última atualização", new Date()],
      ["Folhas publicadas", nomesFolhas.join(", ")],
      [
        "Nota",
        "Cópia de consulta. Os dados e o portal são atualizados manualmente."
      ]
    ]);
  folha
    .getRange(1, 1, 1, 2)
    .setFontWeight("bold");
  folha
    .getRange(2, 2)
    .setNumberFormat("dd/MM/yyyy HH:mm");
  folha.autoResizeColumns(1, 2);
  folha.setFrozenRows(1);
}
