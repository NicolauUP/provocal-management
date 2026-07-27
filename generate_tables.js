/**
 * Gera o resumo de assiduidade a partir dos ensaios gerais e
 * dos ensaios de naipe disponíveis.
 */
function gerarResumo() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const membros = lerMembrosAtivos_(spreadsheet);

  if (membros.length === 0) {
    throw new Error(
      'Não foram encontrados membros com Estado igual a "Ativo".'
    );
  }

  const analise = analisarParticipacao_(spreadsheet, membros);
  const detalhesEnsaios = analise.detalhes.filter(detalhe =>
    detalhe.ehEnsaio &&
    detalhe.data instanceof Date &&
    detalhe.data >= START_DATE &&
    detalhe.data <= END_DATE
  );
  const registosPorChave = {};

  detalhesEnsaios.forEach(detalhe => {
    const chave = `${detalhe.folha}|${detalhe.linha}`;

    if (!registosPorChave[chave]) {
      registosPorChave[chave] = {
        chave,
        folha: detalhe.folha,
        data: detalhe.data,
        detalhesPorMembro: {}
      };
    }

    registosPorChave[chave].detalhesPorMembro[
      detalhe.chaveNome
    ] = detalhe;
  });

  const registos = Object.values(registosPorChave)
    .sort((a, b) =>
      a.data.getTime() - b.data.getTime() ||
      a.folha.localeCompare(b.folha, "pt")
    );
  const cabecalhos = [
    "Nome",
    ...registos.map(registo => registo.data),
    "Total",
    "%"
  ];
  const linhas = membros.map(membro => {
    let totalPontos = 0;
    let totalMaximo = 0;
    const pontosPorEnsaio = registos.map(registo => {
      const detalhe =
        registo.detalhesPorMembro[membro.chaveNome];

      if (!detalhe || !detalhe.entraDenominador) {
        return "";
      }

      totalPontos += detalhe.pontosObtidos;
      totalMaximo += detalhe.pontosMaximos;
      return detalhe.pontosObtidos;
    });

    return [
      membro.nome,
      ...pontosPorEnsaio,
      totalPontos,
      totalMaximo > 0 ? totalPontos / totalMaximo : 0
    ];
  });

  let folhaResumo = spreadsheet.getSheetByName(SHEETS.SUMMARY);

  if (!folhaResumo) {
    folhaResumo = spreadsheet.insertSheet(SHEETS.SUMMARY);
  } else {
    const filtro = folhaResumo.getFilter();

    if (filtro) {
      filtro.remove();
    }

    folhaResumo.clear();
  }

  folhaResumo
    .getRange(1, 1, 1, cabecalhos.length)
    .setValues([cabecalhos]);
  folhaResumo
    .getRange(2, 1, linhas.length, cabecalhos.length)
    .setValues(linhas);

  customizeSheet(
    folhaResumo,
    membros.length,
    registos.length
  );
  applyConditionalFormatting(
    folhaResumo,
    membros.length,
    registos.length
  );

  if (analise.fontesOmitidas.length > 0) {
    SpreadsheetApp.getUi().alert(
      "Resumo gerado com avisos",
      `Folhas ainda ausentes: ${analise.fontesOmitidas.join(", ")}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
