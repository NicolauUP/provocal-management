/**
 * Mantém o resumo histórico configurado em START_DATE/END_DATE.
 */
function gerarResumo() {
  gerarMapaAssiduidadePeriodo_(
    START_DATE,
    END_DATE,
    SHEETS.SUMMARY,
    "Resumo de assiduidade"
  );
}

/**
 * Pede um mês no formato MM/AAAA e gera uma folha independente.
 */
function gerarMapaAssiduidadeMensal() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt(
    "Mapa mensal de assiduidade",
    "Indica o mês no formato MM/AAAA.",
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const correspondencia = resposta
    .getResponseText()
    .trim()
    .match(/^(\d{1,2})\/(\d{4})$/);

  if (!correspondencia) {
    throw new Error(
      "Mês inválido. Usa o formato MM/AAAA, por exemplo 07/2026."
    );
  }

  const mes = Number(correspondencia[1]);
  const ano = Number(correspondencia[2]);

  if (mes < 1 || mes > 12) {
    throw new Error("O mês deve estar entre 1 e 12.");
  }

  const inicio = new Date(ano, mes - 1, 1);
  const fim = new Date(ano, mes, 0, 23, 59, 59, 999);
  const mesFormatado = String(mes).padStart(2, "0");
  const nomeFolha = `Assiduidade_${ano}_${mesFormatado}`;

  gerarMapaAssiduidadePeriodo_(
    inicio,
    fim,
    nomeFolha,
    `Mapa de assiduidade de ${mesFormatado}/${ano}`
  );
}

/**
 * Constrói uma tabela:
 * Nome | ensaio 1 | ensaio 2 | ... | Total | %
 *
 * Usa ensaios gerais e de naipe. Uma célula fica vazia quando o
 * ensaio não se aplica ao membro; respostas elegíveis valem 5/4/0.
 */
function gerarMapaAssiduidadePeriodo_(
  inicio,
  fim,
  nomeFolha,
  titulo
) {
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
    detalhe.data >= inicio &&
    detalhe.data <= fim
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
    ...registos.map(criarCabecalhoRegistoAssiduidade_),
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

  let folhaResumo = spreadsheet.getSheetByName(nomeFolha);

  if (!folhaResumo) {
    folhaResumo = spreadsheet.insertSheet(nomeFolha);
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

  const mensagens = [
    `${titulo} gerado.`,
    `Ensaios no período: ${registos.length}`
  ];

  if (analise.fontesOmitidas.length > 0) {
    mensagens.push(
      `Folhas ainda ausentes: ${analise.fontesOmitidas.join(", ")}`
    );
  }

  SpreadsheetApp.getUi().alert(
    "Mapa de assiduidade",
    mensagens.join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function criarCabecalhoRegistoAssiduidade_(registo) {
  const dia = String(registo.data.getDate()).padStart(2, "0");
  const mes = String(registo.data.getMonth() + 1).padStart(2, "0");
  const origem = registo.folha === SHEETS.GENERAL_ATTENDANCE
    ? "Geral"
    : registo.folha.replace(/^Ensaios_/, "");

  return `${dia}/${mes} ${origem}`;
}
