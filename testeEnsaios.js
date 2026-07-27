/**
 * Gera um diagnóstico detalhado das cinco fontes de presenças.
 * Alterar o nome abaixo antes de executar manualmente.
 */
function diagnosticarAssiduidade() {
  const NOME_MEMBRO = "Catarina Lopes";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const membros = lerMembrosAtivos_(spreadsheet);
  const chaveProcurada = normalizarNome_(NOME_MEMBRO);
  const membro = membros.find(
    item => item.chaveNome === chaveProcurada
  );

  if (!membro) {
    throw new Error(
      `O membro ativo "${NOME_MEMBRO}" não foi encontrado na folha "${DISTRIBUICAO_CONFIG.folhas.membros}".`
    );
  }

  const analise = analisarParticipacao_(spreadsheet, membros);
  const metricas = analise.metricas[membro.chaveNome];
  const detalhes = analise.detalhes.filter(
    detalhe => detalhe.chaveNome === membro.chaveNome
  );

  const nomeFolhaDebug = "DEBUG_Assiduidade";
  let folhaDebug = spreadsheet.getSheetByName(nomeFolhaDebug);

  if (!folhaDebug) {
    folhaDebug = spreadsheet.insertSheet(nomeFolhaDebug);
  }

  const filtro = folhaDebug.getFilter();

  if (filtro) {
    filtro.remove();
  }

  folhaDebug.clear();

  const resumo = [
    ["Campo", "Valor"],
    ["Membro", membro.nome],
    ["Data de entrada", membro.entrada || "INVÁLIDA"],
    ["Atividades", metricas.atividades],
    ["Ensaios possíveis", metricas.ensaiosPossiveis],
    ["Pontos obtidos", metricas.pontosEnsaios],
    ["Pontos máximos", metricas.ensaiosPossiveis * 5],
    ["Assiduidade", metricas.assiduidade],
    [
      "Folhas de naipe omitidas",
      analise.fontesOmitidas.length > 0
        ? analise.fontesOmitidas.join(", ")
        : "Nenhuma"
    ]
  ];

  folhaDebug
    .getRange(1, 1, resumo.length, 2)
    .setValues(resumo);
  folhaDebug
    .getRange(8, 2)
    .setNumberFormat("0.00%");

  const linhaCabecalho = resumo.length + 2;
  const cabecalhos = [
    "Folha de origem",
    "Linha",
    "Data",
    "Tipo de registo",
    "Nome do membro",
    "Resposta",
    "Data de entrada",
    "Entra no cálculo?",
    "Entra no numerador?",
    "Entra no denominador?",
    "Pontos obtidos",
    "Pontos máximos",
    "Conta como atividade?",
    "Motivo"
  ];

  folhaDebug
    .getRange(linhaCabecalho, 1, 1, cabecalhos.length)
    .setValues([cabecalhos])
    .setFontWeight("bold");

  if (detalhes.length > 0) {
    const linhas = detalhes.map(detalhe => [
      detalhe.folha,
      detalhe.linha,
      detalhe.data || "",
      detalhe.tipoRegisto,
      detalhe.nomeMembro,
      detalhe.resposta,
      detalhe.dataEntrada || "",
      detalhe.entraCalculo,
      detalhe.entraNumerador,
      detalhe.entraDenominador,
      detalhe.pontosObtidos,
      detalhe.pontosMaximos,
      detalhe.contaComoAtividade,
      detalhe.motivo
    ]);

    folhaDebug
      .getRange(
        linhaCabecalho + 1,
        1,
        linhas.length,
        cabecalhos.length
      )
      .setValues(linhas);
  }

  folhaDebug.setFrozenRows(linhaCabecalho);
  folhaDebug.autoResizeColumns(1, cabecalhos.length);

  SpreadsheetApp.getUi().alert(
    "Diagnóstico concluído",
    [
      `Membro: ${membro.nome}`,
      `Atividades: ${metricas.atividades}`,
      `Ensaios possíveis: ${metricas.ensaiosPossiveis}`,
      `Assiduidade: ${(metricas.assiduidade * 100).toFixed(2)}%`,
      analise.fontesOmitidas.length > 0
        ? `Folhas ainda ausentes: ${analise.fontesOmitidas.join(", ")}`
        : "Todas as fontes foram processadas.",
      "",
      `Consulta a folha "${nomeFolhaDebug}".`
    ].join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
