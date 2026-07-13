function diagnosticarAssiduidade() {
  const NOME_MEMBRO = "Catarina Lopes"; // alterar para o membro a testar

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const folhaMembros = obterFolhaObrigatoria_(
    ss,
    DISTRIBUICAO_CONFIG.folhas.membros
  );

  const folhaRespostas = obterFolhaObrigatoria_(
    ss,
    DISTRIBUICAO_CONFIG.folhas.respostas
  );

  /**
   * ==========================================================
   * ENCONTRAR O MEMBRO
   * ==========================================================
   */

  const dadosMembros = folhaMembros
    .getDataRange()
    .getValues();

  const cabecalhosMembros = dadosMembros[0];

  const indiceNomeMembros = encontrarIndiceCabecalho_(
    cabecalhosMembros,
    DISTRIBUICAO_CONFIG.cabecalhos.membros.nome,
    true,
    folhaMembros.getName()
  );

  const indiceEntrada = encontrarIndiceCabecalho_(
    cabecalhosMembros,
    DISTRIBUICAO_CONFIG.cabecalhos.membros.entrada,
    true,
    folhaMembros.getName()
  );

  const indiceEstado = encontrarIndiceCabecalho_(
    cabecalhosMembros,
    DISTRIBUICAO_CONFIG.cabecalhos.membros.estado,
    true,
    folhaMembros.getName()
  );

  const chaveProcurada = normalizarNome_(NOME_MEMBRO);

  let membro = null;

  for (let i = 1; i < dadosMembros.length; i++) {
    const nome = String(
      dadosMembros[i][indiceNomeMembros] ?? ""
    ).trim();

    if (normalizarNome_(nome) !== chaveProcurada) {
      continue;
    }

    membro = {
      nome,
      entradaOriginal: dadosMembros[i][indiceEntrada],
      entrada: converterData_(
        dadosMembros[i][indiceEntrada]
      ),
      estado: dadosMembros[i][indiceEstado],
      linha: i + 1
    };

    break;
  }

  if (!membro) {
    throw new Error(
      `O membro "${NOME_MEMBRO}" não foi encontrado na folha Membros.`
    );
  }

  /**
   * ==========================================================
   * IDENTIFICAR COLUNAS EM FORM_RESPONSES
   * ==========================================================
   */

  const dadosRespostas = folhaRespostas
    .getDataRange()
    .getValues();

  const cabecalhosRespostas = dadosRespostas[0];

  const indiceData = encontrarIndiceCabecalho_(
    cabecalhosRespostas,
    DISTRIBUICAO_CONFIG.cabecalhos.respostas.dataAtividade,
    true,
    folhaRespostas.getName()
  );

  const indiceTipo = encontrarIndiceCabecalho_(
    cabecalhosRespostas,
    DISTRIBUICAO_CONFIG.cabecalhos.respostas.tipoAtividade,
    true,
    folhaRespostas.getName()
  );

  let indicePresencaMembro = -1;
  let cabecalhoPresencaEncontrado = "";

  cabecalhosRespostas.forEach((cabecalho, indice) => {
    const nomeExtraido =
      extrairNomeCabecalhoPresenca_(cabecalho);

    if (
      nomeExtraido &&
      normalizarNome_(nomeExtraido) === chaveProcurada
    ) {
      indicePresencaMembro = indice;
      cabecalhoPresencaEncontrado = cabecalho;
    }
  });

  if (indicePresencaMembro < 0) {
    throw new Error(
      `Não encontrei a coluna de presenças de "${NOME_MEMBRO}" em Form_Responses.`
    );
  }

  /**
   * ==========================================================
   * ANALISAR LINHA A LINHA
   * ==========================================================
   */

  const tipoEnsaioConfigurado = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.tipoEnsaio
  );

  const presente = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.presente
  );

  const atraso = normalizarTexto_(
    DISTRIBUICAO_CONFIG.valores.atraso
  );

  let ensaiosPossiveis = 0;
  let pontosObtidos = 0;

  const linhasDebug = [];

  for (let i = 1; i < dadosRespostas.length; i++) {
    const linha = dadosRespostas[i];

    if (linhaCompletamenteVazia_(linha)) {
      continue;
    }

    const dataOriginal = linha[indiceData];
    const dataAtividade = converterData_(dataOriginal);

    const tipoOriginal = String(
      linha[indiceTipo] ?? ""
    ).trim();

    const tipoNormalizado =
      normalizarTexto_(tipoOriginal);

    const respostaOriginal = String(
      linha[indicePresencaMembro] ?? ""
    ).trim();

    const respostaNormalizada =
      normalizarTexto_(respostaOriginal);

    const dataValida =
      dataAtividade instanceof Date &&
      !isNaN(dataAtividade.getTime());

    const posteriorEntrada =
      dataValida &&
      membro.entrada instanceof Date &&
      !isNaN(membro.entrada.getTime())
        ? compararApenasData_(
            dataAtividade,
            membro.entrada
          ) >= 0
        : false;

    const tipoContaComoEnsaio =
      tipoNormalizado === tipoEnsaioConfigurado;

    let contaDenominador = false;
    let pontosLinha = 0;
    let motivo = "";

    if (!dataValida) {
      motivo = "Data inválida";
    } else if (!posteriorEntrada) {
      motivo = "Anterior à entrada";
    } else if (!tipoContaComoEnsaio) {
      motivo = `Tipo diferente de "${DISTRIBUICAO_CONFIG.valores.tipoEnsaio}"`;
    } else {
      contaDenominador = true;
      ensaiosPossiveis += 1;

      if (respostaNormalizada === presente) {
        pontosLinha = 5;
        motivo = "Presente";
      } else if (respostaNormalizada === atraso) {
        pontosLinha = 4;
        motivo = "Atraso";
      } else if (!respostaNormalizada) {
        motivo = "Resposta vazia: conta como 0";
      } else {
        motivo = `Resposta "${respostaOriginal}": conta como 0`;
      }

      pontosObtidos += pontosLinha;
    }

    linhasDebug.push([
      i + 1,
      dataOriginal,
      dataValida
        ? Utilities.formatDate(
            dataAtividade,
            Session.getScriptTimeZone(),
            "dd/MM/yyyy"
          )
        : "",
      tipoOriginal,
      tipoNormalizado,
      respostaOriginal,
      posteriorEntrada,
      tipoContaComoEnsaio,
      contaDenominador,
      pontosLinha,
      motivo
    ]);
  }

  const pontosMaximos = ensaiosPossiveis * 5;

  const assiduidade =
    pontosMaximos > 0
      ? pontosObtidos / pontosMaximos
      : 0;

  /**
   * ==========================================================
   * ESCREVER FOLHA DE DEBUG
   * ==========================================================
   */

  const nomeFolhaDebug = "DEBUG_Assiduidade";

  let folhaDebug =
    ss.getSheetByName(nomeFolhaDebug);

  if (!folhaDebug) {
    folhaDebug = ss.insertSheet(nomeFolhaDebug);
  }

  const filtro = folhaDebug.getFilter();

  if (filtro) {
    filtro.remove();
  }

  folhaDebug.clear();

  const resumo = [
    ["Campo", "Valor"],
    ["Membro", membro.nome],
    ["Estado", membro.estado],
    ["Data de entrada original", membro.entradaOriginal],
    [
      "Data de entrada interpretada",
      membro.entrada
        ? Utilities.formatDate(
            membro.entrada,
            Session.getScriptTimeZone(),
            "dd/MM/yyyy"
          )
        : "INVÁLIDA"
    ],
    [
      "Coluna de presença",
      cabecalhoPresencaEncontrado
    ],
    [
      "Tipo de ensaio configurado",
      DISTRIBUICAO_CONFIG.valores.tipoEnsaio
    ],
    ["Ensaios possíveis", ensaiosPossiveis],
    ["Pontos obtidos", pontosObtidos],
    ["Pontos máximos", pontosMaximos],
    ["Assiduidade", assiduidade]
  ];

  folhaDebug
    .getRange(1, 1, resumo.length, 2)
    .setValues(resumo);

  folhaDebug
    .getRange(resumo.length, 2)
    .setNumberFormat("0.00%");

  const linhaCabecalho = resumo.length + 2;

  const cabecalhosDebug = [
    "Linha Form",
    "Data original",
    "Data interpretada",
    "Tipo original",
    "Tipo normalizado",
    "Resposta",
    "Posterior à entrada?",
    "É tipo Ensaio?",
    "Conta no denominador?",
    "Pontos",
    "Motivo"
  ];

  folhaDebug
    .getRange(
      linhaCabecalho,
      1,
      1,
      cabecalhosDebug.length
    )
    .setValues([cabecalhosDebug])
    .setFontWeight("bold");

  if (linhasDebug.length > 0) {
    folhaDebug
      .getRange(
        linhaCabecalho + 1,
        1,
        linhasDebug.length,
        cabecalhosDebug.length
      )
      .setValues(linhasDebug);
  }

  folhaDebug.setFrozenRows(linhaCabecalho);
  folhaDebug.autoResizeColumns(
    1,
    cabecalhosDebug.length
  );

  SpreadsheetApp.getUi().alert(
    "Diagnóstico concluído",
    [
      `Membro: ${membro.nome}`,
      `Ensaios possíveis: ${ensaiosPossiveis}`,
      `Pontos obtidos: ${pontosObtidos}`,
      `Assiduidade: ${(assiduidade * 100).toFixed(2)}%`,
      "",
      'Consulta a folha "DEBUG_Assiduidade".'
    ].join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}