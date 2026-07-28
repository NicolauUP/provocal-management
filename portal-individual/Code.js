/**
 * Projeto Apps Script exclusivamente para o portal individual.
 * Nunca contém funções de cálculo, publicação ou escrita na spreadsheet
 * administrativa.
 */
const PORTAL_CONFIG = {
  DATA_SPREADSHEET_ID: "1UwlcSqP5xDMIaSGmxwvRlu8R3jsp3U2PYHAl1-Nl4ME",
  DATA_SHEET: "_Dados_Portal_Privado",
  VALUES_SHEET: "_Valores_Portal_Privado",
  GLOBAL_SHEET: "_Resumo_Coro_Privado",
  OAUTH_CLIENT_ID:
    "521246659306-jhk4d66f6frllrvr5sdmjk4lrcpjhjvd.apps.googleusercontent.com",
  OAUTH_CLIENT_SECRET_PROPERTY: "GOOGLE_OAUTH_CLIENT_SECRET",
  WEB_APP_URL:
    "https://script.google.com/macros/s/AKfycbxeFWoAQQ_cgTQbvuKNM4_QrJQrCkcr1EEjE3vmytDTnbOO9LIBwSp4UE2F41PThQ6hPg/exec",
  SESSION_SECONDS: 21600
};

function doGet(evento) {
  const pagina = HtmlService.createTemplateFromFile("index");
  const parametros = (evento && evento.parameter) || {};
  pagina.sessao = String(parametros.sessao || "");

  if (parametros.code || parametros.error) {
    try {
      pagina.sessao = concluirOAuthPortal_(parametros);
    } catch (erro) {
      pagina.sessao = "";
    }
  }

  return pagina
    .evaluate()
    .setTitle("Consulta individual do Coro")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Cria uma autorização OAuth por redirecionamento direto para a web app.
 */
function obterUrlLoginPortal() {
  obterSegredoOAuthPortal_();

  const state = Utilities.getUuid() + Utilities.getUuid();
  const cache = CacheService.getScriptCache();

  cache.put(`portal-oauth:${state}`, "pendente", 1800);

  return criarUrlOAuthGoogle_({
    client_id: PORTAL_CONFIG.OAUTH_CLIENT_ID,
    redirect_uri: PORTAL_CONFIG.WEB_APP_URL,
    response_type: "code",
    scope: "openid email",
    prompt: "select_account",
    state
  });
}

function concluirOAuthPortal_(parametros) {
  const state = String(parametros.state || "");
  const cache = CacheService.getScriptCache();

  if (!state || !cache.get(`portal-oauth:${state}`)) {
    throw new Error("O pedido de autenticação expirou. Tenta novamente.");
  }

  cache.remove(`portal-oauth:${state}`);

  if (parametros.error || !parametros.code) {
    throw new Error("O início de sessão foi cancelado ou recusado.");
  }

  const identidade = trocarCodigoOAuthPorIdentidade_(parametros.code);
  const membro = obterMembroPortalPorEmail_(identidade.email);

  if (!membro) {
    throw new Error("Este email não está associado a um membro ativo.");
  }

  const sessao = Utilities.getUuid() + Utilities.getUuid();
  cache.put(
    `portal-sessao:${sessao}`,
    JSON.stringify({ email: identidade.email }),
    PORTAL_CONFIG.SESSION_SECONDS
  );
  return sessao;
}

function obterResumoGeralPortal(sessao) {
  const membro = obterMembroAutenticadoPortal_(sessao);
  let pontosObtidos = 0;
  let pontosMaximos = 0;
  let atividades = 0;
  let atividadesPossiveis = 0;
  let ensaiosFrequentados = 0;
  const mensal = {};

  membro.linhas.forEach(linha => {
    if (linha.tipo === "Concerto") {
      atividadesPossiveis += 1;
      if (linha.contaComoAtividade) {
        atividades += 1;
      }
      return;
    }

    pontosObtidos += linha.pontosObtidos;
    pontosMaximos += linha.pontosMaximos;
    if (linha.pontosObtidos > 0) ensaiosFrequentados += 1;
    mensal[linha.mes] = mensal[linha.mes] || { pontos: 0, maximos: 0 };
    mensal[linha.mes].pontos += linha.pontosObtidos;
    mensal[linha.mes].maximos += linha.pontosMaximos;
  });

  return {
    nome: membro.nome,
    atividades,
    atividadesPossiveis,
    ensaios: membro.linhas.filter(linha => linha.tipo !== "Concerto").length,
    ensaiosFrequentados,
    pontosObtidos,
    pontosMaximos,
    assiduidade: pontosMaximos ? pontosObtidos / pontosMaximos : 0,
    valores: membro.valores,
    global: obterResumoGlobalPortal_(),
    atividadesRealizadas: membro.linhas,
    ensaiosLista: membro.linhas.filter(linha => linha.tipo !== "Concerto")
      .sort((a, b) => b.dataChave.localeCompare(a.dataChave)),
    concertosLista: membro.linhas.filter(linha => linha.tipo === "Concerto")
      .sort((a, b) => b.dataChave.localeCompare(a.dataChave)),
    evolucaoMensal: Object.keys(mensal).sort().map(mes => ({
      mes,
      percentagem: mensal[mes].maximos
        ? mensal[mes].pontos / mensal[mes].maximos
        : 0
    }))
  };
}

function obterResumoGlobalPortal_() {
  const folha = SpreadsheetApp.openById(PORTAL_CONFIG.DATA_SPREADSHEET_ID)
    .getSheetByName(PORTAL_CONFIG.GLOBAL_SHEET);
  if (!folha || folha.getLastRow() < 2) return { fundo: 0, valorPorPonto: 0, pontos: 0 };
  const linha = folha.getRange(2, 1, 1, 3).getValues()[0];
  return { fundo: Number(linha[0]) || 0, valorPorPonto: Number(linha[1]) || 0, pontos: Number(linha[2]) || 0 };
}

function obterMembroAutenticadoPortal_(sessao) {
  const valor = CacheService.getScriptCache().get(
    `portal-sessao:${String(sessao || "")}`
  );

  if (!valor) {
    throw new Error("A sessão expirou. Inicia sessão novamente.");
  }

  const dados = JSON.parse(valor);
  const membro = obterMembroPortalPorEmail_(dados.email);

  if (!membro) {
    throw new Error("Este email já não está associado a um membro ativo.");
  }

  return membro;
}

function trocarCodigoOAuthPorIdentidade_(codigo) {
  const resposta = UrlFetchApp.fetch(
    "https://oauth2.googleapis.com/token",
    {
      method: "post",
      payload: {
        code: String(codigo),
        client_id: PORTAL_CONFIG.OAUTH_CLIENT_ID,
        client_secret: obterSegredoOAuthPortal_(),
        redirect_uri: PORTAL_CONFIG.WEB_APP_URL,
        grant_type: "authorization_code"
      },
      muteHttpExceptions: true
    }
  );

  if (resposta.getResponseCode() !== 200) {
    throw new Error("O Google não aceitou a autenticação.");
  }

  const idToken = JSON.parse(resposta.getContentText()).id_token;
  return validarIdTokenGoogle_(idToken);
}

function validarIdTokenGoogle_(idToken) {
  const resposta = UrlFetchApp.fetch(
    "https://oauth2.googleapis.com/tokeninfo?id_token=" +
      encodeURIComponent(String(idToken || "")),
    { muteHttpExceptions: true }
  );

  if (resposta.getResponseCode() !== 200) {
    throw new Error("Não foi possível validar a sessão Google.");
  }

  const dados = JSON.parse(resposta.getContentText());
  const agora = Math.floor(Date.now() / 1000);

  if (
    dados.aud !== PORTAL_CONFIG.OAUTH_CLIENT_ID ||
    !["accounts.google.com", "https://accounts.google.com"].includes(dados.iss) ||
    Number(dados.exp) <= agora ||
    String(dados.email_verified) !== "true" ||
    !dados.email
  ) {
    throw new Error("A sessão Google não é válida.");
  }

  return { email: String(dados.email).trim().toLowerCase() };
}

function obterMembroPortalPorEmail_(email) {
  const folha = SpreadsheetApp.openById(
    PORTAL_CONFIG.DATA_SPREADSHEET_ID
  ).getSheetByName(PORTAL_CONFIG.DATA_SHEET);

  if (!folha || folha.getLastRow() < 2) {
    throw new Error("Ainda não existem dados publicados para o portal.");
  }

  const dados = folha.getDataRange().getValues();
  const cabecalhos = dados[0].map(valor => String(valor).trim());
  const indices = {
    email: cabecalhos.indexOf("Email"),
    nome: cabecalhos.indexOf("Nome"),
    data: cabecalhos.indexOf("Data"),
    mes: cabecalhos.indexOf("Mês"),
    origem: cabecalhos.indexOf("Origem"),
    tipo: cabecalhos.indexOf("Tipo"),
    nomeConcerto: cabecalhos.indexOf("Nome da Atividade"),
    resposta: cabecalhos.indexOf("Resposta"),
    pontosObtidos: cabecalhos.indexOf("Pontos obtidos"),
    pontosMaximos: cabecalhos.indexOf("Pontos máximos"),
    contaComoAtividade: cabecalhos.indexOf("Conta como atividade")
  };

  if (Object.keys(indices).some(chave => indices[chave] < 0)) {
    throw new Error("O snapshot privado do portal não tem os cabeçalhos esperados.");
  }

  const linhas = dados.slice(1)
    .filter(linha => String(linha[indices.email]).trim().toLowerCase() === email)
    .map(linha => ({
      mes: normalizarMesPortal_(linha[indices.mes]),
      data: formatarDataPortal_(linha[indices.data]),
      dataChave: formatarDataChavePortal_(linha[indices.data]),
      origem: String(linha[indices.origem] || ""),
      tipo: String(linha[indices.tipo] || ""),
      nomeConcerto: String(linha[indices.nomeConcerto] || ""),
      resposta: String(linha[indices.resposta] || ""),
      pontosObtidos: Number(linha[indices.pontosObtidos]) || 0,
      pontosMaximos: Number(linha[indices.pontosMaximos]) || 0,
      contaComoAtividade: Boolean(linha[indices.contaComoAtividade])
    }));

  const valores = obterValoresPortalPorEmail_(email);

  if (!linhas.length && !valores) return null;

  return {
    nome: valores ? valores.nome : String(dados.find(linha =>
      String(linha[indices.email]).trim().toLowerCase() === email
    )[indices.nome]).trim(),
    meses: Array.from(new Set(linhas.map(linha => linha.mes)))
      .filter(Boolean)
      .sort()
      .reverse(),
    linhas,
    valores: valores || {
      fundoComum: 0, individual: 0, apoios: 0, final: 0
    }
  };
}

function obterValoresPortalPorEmail_(email) {
  const folha = SpreadsheetApp.openById(PORTAL_CONFIG.DATA_SPREADSHEET_ID)
    .getSheetByName(PORTAL_CONFIG.VALUES_SHEET);
  if (!folha || folha.getLastRow() < 2) return null;
  const dados = folha.getDataRange().getValues();
  const cabecalhos = dados[0].map(valor => String(valor).trim());
  const indice = nome => cabecalhos.indexOf(nome);
  const indices = {
    email: indice("Email"), nome: indice("Nome"),
    fundoComum: indice("Valor Fundo Comum"),
    individual: indice("Valor Individual"), apoios: indice("Valor Apoios"),
    final: indice("Valor Final")
  };
  if (Object.keys(indices).some(chave => indices[chave] < 0)) {
    throw new Error("O snapshot de valores do portal não tem os cabeçalhos esperados.");
  }
  const linha = dados.slice(1).find(item =>
    String(item[indices.email]).trim().toLowerCase() === email
  );
  if (!linha) return null;
  return {
    nome: String(linha[indices.nome] || "").trim(),
    pontos: Number(linha[indice("Pontos Finais")]) || 0,
    fundoComum: Number(linha[indices.fundoComum]) || 0,
    individual: Number(linha[indices.individual]) || 0,
    apoios: Number(linha[indices.apoios]) || 0,
    final: Number(linha[indices.final]) || 0
  };
}

function obterSegredoOAuthPortal_() {
  const segredo = PropertiesService.getScriptProperties().getProperty(
    PORTAL_CONFIG.OAUTH_CLIENT_SECRET_PROPERTY
  );

  if (!segredo) {
    throw new Error("O segredo OAuth do portal ainda não foi configurado.");
  }

  return segredo;
}

function criarUrlOAuthGoogle_(parametros) {
  return "https://accounts.google.com/o/oauth2/v2/auth?" +
    Object.keys(parametros)
      .map(chave => `${encodeURIComponent(chave)}=${encodeURIComponent(parametros[chave])}`)
      .join("&");
}

function normalizarMesPortal_(valor) {
  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "yyyy-MM");
  }

  return String(valor || "").trim();
}

function formatarDataPortal_(valor) {
  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }

  return String(valor || "");
}

function formatarDataChavePortal_(valor) {
  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return "";
}
