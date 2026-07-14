# Projeto de gestão do coro

## Ambiente

- Google Apps Script com runtime V8.
- O código é sincronizado com Google Apps Script através de clasp.
- Não usar APIs específicas de Node.js no código executado pelo Apps Script.
- Todas as funções e constantes dos ficheiros `.gs` partilham o namespace global.
- Não duplicar nomes de funções, constantes ou `onOpen`.

## Regras de segurança

- Não alterar regras de negócio sem indicação explícita.
- Não executar `clasp push` automaticamente.
- Não apagar ou renomear folhas, colunas ou dados existentes.
- Não alterar a folha `Form_Responses`.
- Preservar os valores manuais das colunas:
  - `Valor Individual`
  - `Valor Apoios`
- Antes de qualquer refatoração, identificar dependências e possíveis alterações de comportamento.
- Preferir alterações pequenas, verificáveis e separadas em commits.

## Folhas

- `Membros`
- `Form_Responses`
- `Movimento_PREWCG`
- `Movimentos`
- `Distribuição`
- `Resumo`

## Regras de negócio da distribuição

- Apenas membros com `Estado = Ativo` entram nos cálculos.
- O valor pré-WCG é dividido igualmente pelos membros ativos cuja data de entrada é anterior a 01/05/2026.
- O saldo financeiro é `soma das receitas - soma das despesas`.
- Em concertos:
  - `Presente` vale uma atividade.
  - Concertos nunca entram no cálculo da assiduidade.
- Em ensaios:
  - `Presente` vale 5 pontos.
  - `Atraso` vale 4 pontos.
  - `Falta` vale 0 pontos.
- Apenas ensaios ocorridos na data de entrada do membro ou posteriormente entram no denominador.
- Assiduidade é:
  `pontos obtidos nos ensaios / (número de ensaios possíveis * 5)`.
- Pontos finais são:
  `Atividades * (pesoAtividades + (1 - pesoAtividades) * Assiduidade)`.
- `pesoAtividades` é uma variável de configuração, atualmente 0.75.
- O fundo comum é distribuído proporcionalmente aos pontos finais.
- Valor final é:
  `Valor Caixa Antes WCG + Valor Fundo Comum + Valor Individual + Valor Apoios`.

## Estilo pretendido

- Funções pequenas e com uma responsabilidade clara.
- Separar leitura de dados, cálculo, escrita e formatação.
- Evitar índices de colunas codificados diretamente.
- Localizar colunas através dos cabeçalhos.
- Centralizar nomes de folhas, cabeçalhos, pesos e valores válidos na configuração.
- Funções puras para cálculos sempre que possível.
- As funções que usam `SpreadsheetApp` devem limitar-se à leitura e escrita.
- Eliminar código duplicado e funções de diagnóstico obsoletas apenas depois de confirmar que já não são utilizadas.