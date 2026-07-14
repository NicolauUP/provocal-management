# Análise integral e plano de refatoração

## 1. Âmbito e método

Esta análise cobre todos os ficheiros presentes no repositório em 13/07/2026:

- `.clasp.json`
- `.gitignore`
- `AGENTS.md`
- `README.md`
- `TODO.md`
- `appsscript.json`
- `config.js`
- `formatting.js`
- `generate_tables.js`
- `give_money.js`
- `testeEnsaios.js`

Foi feita uma auditoria estática. Não foi executado código contra a folha de cálculo real e não foi feito `clasp push`. Por isso, referências externas que o Google Apps Script possa invocar diretamente — botões, desenhos, macros, acionadores instaláveis ou chamadas manuais pelo editor — não podem ser confirmadas apenas pelo repositório.

Os ficheiros `.js` são enviados pelo `clasp` de acordo com `.clasp.json`. Tal como acontece com `.gs`, as declarações de topo partilham o namespace do projeto Apps Script. A separação por ficheiro não cria módulos nem isolamento.

## 2. Resumo executivo

O núcleo de `give_money.js` implementa corretamente a maior parte das regras de distribuição: filtra membros ativos, separa concertos de ensaios, calcula a assiduidade dos ensaios, aplica o peso `0.75`, distribui o fundo proporcionalmente e preserva numericamente as duas colunas manuais. Contudo, o projeto tem dois caminhos diferentes para assiduidade:

1. `calcularMetricasParticipacao_`, usado na distribuição, distingue explicitamente `Concerto` e `Ensaio`.
2. `gerarResumo`, usado na folha `Resumo`, não lê o tipo de atividade e trata todas as linhas dentro do intervalo de datas como ensaios. Assim, concertos podem entrar na assiduidade do resumo.

Há ainda um erro bloqueante em `gerarResumo`: declara `point`, mas usa `points`. A primeira atividade válida provoca `ReferenceError: points is not defined`, pelo que o resumo não chega normalmente a ser gerado.

Conclusões principais:

- Existem **8 constantes globais** e **23 funções globais**.
- Existe **um único `onOpen`**, em `give_money.js:1119`; não há duplicação atual.
- Não foram encontrados nomes de funções ou constantes de topo exatamente duplicados.
- Há configuração funcionalmente duplicada entre `config.js` e `DISTRIBUICAO_CONFIG`.
- `DISTRIBUTION_START_DATE`, `ACTIVITY_TYPES` e `ATTENDANCE_STATES` não são usados.
- `gerarResumo` e `diagnosticarAssiduidade` só aparecem nas próprias declarações. Podem ser pontos de entrada manuais externos, portanto não devem ser removidos sem confirmar acionadores, botões, macros e histórico de utilização.
- A maior concentração de responsabilidades está em `give_money.js` (1 336 linhas) e em `gerarResumo`/`diagnosticarAssiduidade`, que misturam leitura, transformação, cálculo, escrita e/ou interface.
- O principal problema de desempenho escalável é a escrita célula a célula no resumo; a distribuição já faz as escritas de dados principais em lote.
- Não são usadas APIs específicas de Node.js.
- Nenhum fluxo escreve em `Form_Responses`; todos os acessos a essa folha são de leitura.

## 3. Inventário de constantes globais

| Símbolo | Ficheiro | Responsabilidade | Dependentes | Estado |
|---|---|---|---|---|
| `SHEETS` | `config.js:6` | Nomes das seis folhas oficiais | `gerarResumo` usa apenas `RESPONSES` e `SUMMARY` | Parcialmente usado; duplica `DISTRIBUICAO_CONFIG.folhas` |
| `POINTS` | `config.js:19` | Mapa `Presente=5`, `Atraso=4`, `Falta=0` | `gerarResumo` | Usado, mas o resultado é guardado em `point` e depois é lido como `points` |
| `START_DATE` | `config.js:29` | Início do intervalo do resumo: 01/05/2026 | `gerarResumo` | Usado apenas no resumo |
| `END_DATE` | `config.js:30` | Fim do intervalo do resumo: 30/06/2026 | `gerarResumo` | Usado apenas no resumo |
| `DISTRIBUTION_START_DATE` | `config.js:34` | Data 01/05/2026 | Nenhum | Código morto atual; duplica `dataLimitePreWCG` |
| `ACTIVITY_TYPES` | `config.js:36` | Valores `Ensaio` e `Concerto` | Nenhum | Código morto atual; duplica `DISTRIBUICAO_CONFIG.valores` |
| `ATTENDANCE_STATES` | `config.js:41` | Valores de presença | Nenhum | Código morto atual; duplica `POINTS` e `DISTRIBUICAO_CONFIG.valores` |
| `DISTRIBUICAO_CONFIG` | `give_money.js:7` | Folhas, cabeçalhos, valores válidos, data pré-WCG e peso | Toda a distribuição e o diagnóstico | Ativo, mas sobrepõe conceitos de `config.js`; `valores.falta` não é consultado |

### Observações sobre configuração

- Os nomes de folhas estão centralizados duas vezes e em estilos diferentes: `SHEETS.RESPONSES` versus `DISTRIBUICAO_CONFIG.folhas.respostas`.
- A data de 01/05/2026 existe em `START_DATE`, `DISTRIBUTION_START_DATE` e `DISTRIBUICAO_CONFIG.dataLimitePreWCG`, embora tenha significados parcialmente diferentes.
- Estados e tipos existem em `POINTS`, `ACTIVITY_TYPES`, `ATTENDANCE_STATES` e `DISTRIBUICAO_CONFIG.valores`.
- Os cabeçalhos da distribuição estão centralizados, mas o código de escrita/formatação ainda usa diretamente os números de coluna `1` a `10` e fórmulas com `C`, `G:I` e `J`.
- Os cabeçalhos `Nome`, `Total` e `%` do resumo, assim como `DEBUG_Assiduidade`, não estão na configuração.
- A configuração usa simultaneamente inglês (`SHEETS`, `POINTS`, `START_DATE`) e português (`DISTRIBUICAO_CONFIG`, `folhas`, `cabecalhos`).

## 4. Inventário de funções e dependências

### 4.1 Resumo e formatação

| Função | Ficheiro | Responsabilidade atual | Dependências diretas | Efeitos / observações |
|---|---|---|---|---|
| `gerarResumo` | `generate_tables.js:1` | Lê respostas, filtra datas, calcula pontos, cria/limpa e preenche `Resumo`, chama formatação | `SpreadsheetApp`, `SHEETS`, `START_DATE`, `END_DATE`, `POINTS`, `customizeSheet`, `applyConditionalFormatting` | Mistura todas as camadas; escreve célula a célula; contém o erro `point`/`points`; não filtra tipos nem membros ativos |
| `customizeSheet` | `formatting.js:2` | Formata integralmente `Resumo` | `SpreadsheetApp.BorderStyle` | Só formatação, mas usa muitos índices derivados e faz zebra por linha com uma chamada remota por linha |
| `applyConditionalFormatting` | `formatting.js:160` | Substitui as regras condicionais do resumo | `SpreadsheetApp.newConditionalFormatRule` | Só formatação; seis regras recriadas em cada execução |

### 4.2 Distribuição

| Função | Ficheiro | Responsabilidade atual | Dependências diretas | Efeitos / observações |
|---|---|---|---|---|
| `gerarDistribuicao` | `give_money.js:93` | Orquestra leitura, cálculos, composição, escrita e alerta | Quase todas as funções de distribuição, `DISTRIBUICAO_CONFIG`, `SpreadsheetApp` | Contém também cálculos de elegibilidade, rateio e composição das linhas |
| `lerMembrosAtivos_` | `give_money.js:225` | Lê `Membros`, localiza cabeçalhos, filtra ativos, converte datas e ordena | `obterFolhaObrigatoria_`, `encontrarIndiceCabecalho_`, `normalizarTexto_`, `normalizarNome_`, `converterData_` | Mistura I/O com transformação, regra de estado e ordenação |
| `calcularSaldoMovimentos_` | `give_money.js:351` | Lê uma folha de movimentos e calcula receitas menos despesas até à primeira linha totalmente vazia | `obterFolhaObrigatoria_`, `encontrarIndiceCabecalho_`, `linhaCompletamenteVazia_`, `converterNumero_` | Mistura I/O e cálculo; usada para as duas folhas financeiras |
| `calcularMetricasParticipacao_` | `give_money.js:427` | Lê respostas, associa colunas a membros e calcula atividades, assiduidade e pontos | Utilitários de folha, texto, nome e data; `DISTRIBUICAO_CONFIG` | Mistura leitura, parsing, validação e regras centrais; o laço de cálculo em memória é adequado |
| `compararApenasData_` | `give_money.js:660` | Compara datas sem horas | `Date` | Pura |
| `extrairNomeCabecalhoPresenca_` | `give_money.js:681` | Extrai o nome de um cabeçalho de presença | Expressão regular | Pura; usada também pelo diagnóstico |
| `lerValoresManuaisExistentes_` | `give_money.js:711` | Lê e indexa `Valor Individual` e `Valor Apoios` por nome normalizado | `encontrarIndiceCabecalho_`, `normalizarNome_` | Mistura I/O e transformação; não preserva fórmulas/texto literalmente, apenas o valor devolvido por `getValues()` |
| `escreverDistribuicao_` | `give_money.js:799` | Cria/limpa a folha, escreve cabeçalhos/dados/fórmulas/totais e chama formatação e controlo | `aplicarFormatacaoDistribuicao_`, `escreverResumoTecnico_`, `DISTRIBUICAO_CONFIG` | Mistura construção do modelo de saída, fórmulas, escrita e coordenação de formatação |
| `aplicarFormatacaoDistribuicao_` | `give_money.js:964` | Formata a tabela da distribuição | API de folha | Responsabilidade coerente, mas depende de colunas codificadas de 1 a 10 |
| `escreverResumoTecnico_` | `give_money.js:1063` | Escreve e formata dados de controlo abaixo da tabela | API de folha | Mistura preparação de dados, escrita e formatação |
| `onOpen` | `give_money.js:1119` | Cria o menu `Coro` | `SpreadsheetApp.getUi`, `gerarDistribuicao` por nome textual | Único `onOpen` encontrado; o menu não expõe resumo nem diagnóstico |

### 4.3 Utilitários

| Função | Ficheiro | Responsabilidade | Dependentes | Estado |
|---|---|---|---|---|
| `obterFolhaObrigatoria_` | `give_money.js:1137` | Obtém uma folha ou lança erro | Distribuição e diagnóstico | Ativa; I/O |
| `encontrarIndiceCabecalho_` | `give_money.js:1158` | Localiza uma alternativa de cabeçalho normalizada | Leitores e diagnóstico | Pura; renormaliza todos os cabeçalhos em cada chamada |
| `normalizarTexto_` | `give_money.js:1202` | Limpa espaços, caixa e acentos | Vários fluxos | Pura |
| `normalizarNome_` | `give_money.js:1215` | Delega em `normalizarTexto_` para chaves de nomes | Distribuição e diagnóstico | Pura; semanticamente útil apesar de ser uma linha |
| `converterNumero_` | `give_money.js:1228` | Converte números/texto monetário para número | Saldos e valores manuais | Pura; entradas inválidas tornam-se silenciosamente `0` |
| `converterData_` | `give_money.js:1275` | Converte `Date` ou texto em data | Membros, métricas e diagnóstico | Pura em termos de I/O; aceita datas portuguesas impossíveis que o construtor normaliza, por exemplo 31/02 |
| `linhaCompletamenteVazia_` | `give_money.js:1320` | Testa se uma linha está vazia | Saldos, métricas e diagnóstico | Pura |
| `formatarEuro_` | `give_money.js:1329` | Formata euros para o alerta | `gerarDistribuicao` | Pura; depende do suporte de `Intl`/`toLocaleString` do V8 |

### 4.4 Diagnóstico

| Função | Ficheiro | Responsabilidade atual | Dependências diretas | Efeitos / observações |
|---|---|---|---|---|
| `diagnosticarAssiduidade` | `testeEnsaios.js:1` | Procura um membro fixo, recalcula a assiduidade, cria/limpa `DEBUG_Assiduidade`, formata e mostra alerta | Funções e configuração de `give_money.js`, `SpreadsheetApp`, `Utilities`, `Session` | Mistura leitura, cálculo, escrita, formatação e UI; duplica parcialmente a lógica central; nome do membro está codificado no corpo |

## 5. Mapa de dependências

```text
config.js
  └─ gerarResumo
       ├─ customizeSheet
       └─ applyConditionalFormatting

DISTRIBUICAO_CONFIG + utilitários de give_money.js
  ├─ gerarDistribuicao
  │    ├─ lerMembrosAtivos_
  │    ├─ calcularSaldoMovimentos_ (duas folhas)
  │    ├─ calcularMetricasParticipacao_
  │    ├─ lerValoresManuaisExistentes_
  │    └─ escreverDistribuicao_
  │         ├─ aplicarFormatacaoDistribuicao_
  │         └─ escreverResumoTecnico_
  ├─ onOpen ──(nome textual)──> gerarDistribuicao
  └─ diagnosticarAssiduidade
       ├─ obterFolhaObrigatoria_
       ├─ encontrarIndiceCabecalho_
       ├─ normalizarTexto_ / normalizarNome_
       ├─ converterData_ / compararApenasData_
       ├─ extrairNomeCabecalhoPresenca_
       └─ linhaCompletamenteVazia_
```

Dependências externas do runtime:

- `SpreadsheetApp`: acesso às folhas, formatação, menus e alertas.
- `Utilities.formatDate`: datas da folha de diagnóstico.
- `Session.getScriptTimeZone`: fuso usado no diagnóstico.
- Objetos V8 standard: `Date`, `Set`, `Number`, `String`, expressões regulares e métodos de arrays.
- Não há bibliotecas declaradas em `appsscript.json` e não há dependências Node.js.

## 6. Duplicações, código morto e inconsistências

### 6.1 Duplicações exatas

- Não há funções globais com o mesmo nome.
- Não há constantes globais com o mesmo nome.
- Há exatamente um `onOpen`, em `give_money.js:1119`.
- Não há outro gatilho simples (`onEdit`, `onFormSubmit`, etc.) no repositório.

### 6.2 Duplicação conceptual

- Nomes de folhas, tipos de atividade, estados de presença e data de maio estão definidos em mais de uma configuração.
- O cálculo de assiduidade em `diagnosticarAssiduidade` repete a lógica de `calcularMetricasParticipacao_` em vez de consumir uma função pura comum. Já existe divergência para data de entrada inválida: a distribuição inclui as atividades; o diagnóstico exclui todas.
- A extração de nomes no resumo usa dois `replace` simples; a distribuição/diagnóstico usa `extrairNomeCabecalhoPresenca_`, mais estrito e reutilizável.
- O resumo usa `POINTS`; a distribuição codifica `5` e `4` diretamente apesar de ter valores de estado em configuração.
- Datas são analisadas com `new Date(...)` diretamente no resumo e com `converterData_` no resto do projeto.

### 6.3 Código morto ou potencialmente morto

Confirmado por ausência de referências internas:

- `DISTRIBUTION_START_DATE`.
- `ACTIVITY_TYPES`.
- `ATTENDANCE_STATES`.
- `DISTRIBUICAO_CONFIG.valores.falta`.
- A variável local `lastColumn` de `gerarResumo:11`.
- A variável local `point` de `gerarResumo:99` é calculada mas nunca lida devido ao uso incorreto de `points`.

Potencialmente morto, mas não removível sem inspeção externa:

- `gerarResumo` só aparece na própria declaração.
- `diagnosticarAssiduidade` só aparece na própria declaração.
- Funções públicas Apps Script podem ser chamadas pelo editor, por acionadores instaláveis, macros, botões ou desenhos sem referência textual no repositório.

### 6.4 Nomes e documentação inconsistentes

- Mistura de inglês e português: `customizeSheet`, `applyConditionalFormatting`, `START_DATE` versus `gerarDistribuicao`, `aplicarFormatacaoDistribuicao_` e `dataLimitePreWCG`.
- Mistura de convenções de ficheiros: `generate_tables.js`, `give_money.js`, `testeEnsaios.js`.
- As funções internas da distribuição terminam em `_`, mas `customizeSheet` e `applyConditionalFormatting`, também auxiliares, não.
- `testeEnsaios.js` não é um teste automatizado; é uma rotina destrutiva de diagnóstico que cria/limpa uma folha.
- A secção de estrutura do `README.md` refere `Code.js`, que não existe, e não descreve os ficheiros atuais.
- O roadmap do `README.md` marca como pendentes várias funcionalidades já parcialmente implementadas.
- O comentário “PONTOS DOS CONCERTOS” em `give_money.js:410` já não descreve a função, que também calcula ensaios, assiduidade e pontos finais.
- `numberOfRehearsals` e a variável `rehearsal` no resumo representam todas as respostas/atividades no intervalo, não apenas ensaios.

## 7. Separação de responsabilidades

### Funções com separação aceitável

- `compararApenasData_`
- `extrairNomeCabecalhoPresenca_`
- `encontrarIndiceCabecalho_`
- `normalizarTexto_`
- `normalizarNome_`
- `converterNumero_`
- `converterData_`
- `linhaCompletamenteVazia_`
- `formatarEuro_`
- `customizeSheet`, `applyConditionalFormatting` e `aplicarFormatacaoDistribuicao_` estão limitadas a formatação, embora possam ser otimizadas/renomeadas.

### Funções com responsabilidades misturadas

- `gerarResumo`: leitura + interpretação de esquema + cálculo + escrita + formatação.
- `gerarDistribuicao`: orquestração + elegibilidade pré-WCG + rateio + composição + UI.
- `lerMembrosAtivos_`: leitura + normalização + regra de membro ativo + conversão + ordenação.
- `calcularSaldoMovimentos_`: leitura + delimitação das linhas + conversão + cálculo.
- `calcularMetricasParticipacao_`: leitura + descoberta do esquema + associação de membros + validação + cálculo.
- `lerValoresManuaisExistentes_`: leitura + associação e transformação dos valores.
- `escreverDistribuicao_`: criação/limpeza + montagem de cabeçalhos + dados + fórmulas + totais + coordenação da formatação.
- `escreverResumoTecnico_`: construção dos dados + escrita + formatação.
- `diagnosticarAssiduidade`: leitura + cálculo duplicado + construção do relatório + escrita + formatação + alerta.

O maior ganho arquitetural virá de fazer os adaptadores `SpreadsheetApp` devolverem arrays/objetos simples e passá-los a funções puras. Escritores devem receber matrizes já completas e não decidir regras de negócio.

## 8. Conformidade com `AGENTS.md`

| Regra | Estado | Evidência e risco |
|---|---|---|
| Runtime V8, sem APIs Node.js | Conforme | `appsscript.json` define V8; não há APIs Node.js |
| Namespace global sem nomes duplicados | Conforme atualmente | 8 constantes e 23 funções, sem nomes repetidos |
| Não duplicar `onOpen` | Conforme | Só `give_money.js:1119` |
| Não alterar/apagar folhas ou dados existentes | Parcial | Não renomeia/apaga folhas, mas `gerarResumo`, `escreverDistribuicao_` e o diagnóstico executam `clear()` nas folhas de destino |
| Não alterar `Form_Responses` | Conforme | Só há leituras desta folha |
| Preservar `Valor Individual` e `Valor Apoios` | Parcialmente conforme | Os valores calculados por `getValues()` são guardados por nome normalizado e reescritos; fórmulas, texto não numérico, validações e representação original não são preservados |
| Apenas membros ativos nos cálculos | Conforme na distribuição; não conforme no resumo/diagnóstico | `lerMembrosAtivos_` filtra corretamente; `gerarResumo` usa todas as colunas; diagnóstico aceita o membro mesmo que não esteja ativo |
| Pré-WCG dividido por ativos com entrada anterior a 01/05/2026 | Conforme na distribuição | Filtro estrito `< dataLimitePreWCG`; entradas inválidas não recebem |
| Saldo = receitas - despesas | Parcialmente conforme | Fórmula correta; linhas após a primeira linha completamente vazia são ignoradas |
| Concerto `Presente` = uma atividade | Conforme na distribuição | Comparação normalizada e incremento de 1 |
| Concertos nunca entram na assiduidade | Conforme na distribuição; violado no resumo | O ramo de concerto termina com `continue`; `gerarResumo` não lê a coluna de tipo |
| Ensaio: 5/4/0 | Conforme na distribuição e diagnóstico | Presente 5, atraso 4, restantes 0; resumo pretende usar o mapa correto, mas falha em runtime |
| Só ensaios na data de entrada ou posteriores entram no denominador | Parcialmente conforme | Inclusivo para datas válidas; entrada inválida faz a distribuição contar todos os ensaios. O resumo ignora a data de entrada |
| Assiduidade = pontos / (ensaios possíveis × 5) | Conforme na distribuição e diagnóstico | Quando não há ensaios possíveis devolve 0; resumo usa a mesma forma sem proteção para denominador zero |
| Pontos finais = Atividades × fórmula ponderada | Conforme | Implementação literal em `give_money.js:640-646` |
| `pesoAtividades` configurável e igual a 0.75 | Conforme localmente | Está em `DISTRIBUICAO_CONFIG` e é validado; não está na configuração central `config.js` |
| Fundo distribuído proporcionalmente aos pontos | Conforme no caso normal | Se o total de pontos for zero, atribui zero a todos e deixa o saldo sem distribuição; caso-limite não definido |
| Valor final = pré-WCG + fundo + individual + apoios | Conforme | Fórmula `SUM(C;G:I)` exclui corretamente métricas D:F |
| Evitar índices codificados e localizar por cabeçalhos | Parcial | Leitores da distribuição localizam cabeçalhos; resumo assume data na coluna B e membros desde C; escritores/formatadores usam 1..10 e letras fixas |
| Centralizar nomes, cabeçalhos, pesos e valores | Não conforme integralmente | Existem duas configurações e vários literais dispersos |
| Funções pequenas, puras e camadas separadas | Parcial | Há bons utilitários puros, mas os fluxos principais misturam camadas |

## 9. Como concertos podem entrar indevidamente na assiduidade

### Risco confirmado: `gerarResumo`

`gerarResumo` considera válida qualquer linha cuja data da coluna B esteja entre `START_DATE` e `END_DATE`. Nunca procura nem testa `Tipo de Atividade`. Consequências:

- Um concerto dentro de 01/05/2026–30/06/2026 é incluído no número de “ensaios válidos”.
- A presença nesse concerto entra nos pontos do membro e o concerto aumenta o denominador em 5.
- Um concerto com `Presente` pode produzir 5 pontos como se fosse ensaio.
- Um concerto com outra resposta pode reduzir a assiduidade como se fosse uma falta em ensaio.
- Como o resumo também ignora a data de entrada, a distorção afeta inclusive membros que ainda não tinham entrado.

Além disso, `firstMemberColumn = 2` assume que todas as colunas desde C são presenças. Qualquer metadado, incluindo a própria coluna `Tipo de Atividade` se estiver nessa zona, é tratado como membro. A extração de nome não valida o padrão do cabeçalho.

### Distribuição principal: proteção existente

Em `calcularMetricasParticipacao_`, o concerto é tratado antes do ensaio e termina com `continue`. Portanto, uma linha cujo tipo normalize exatamente para `Concerto` não alcança o denominador nem os pontos dos ensaios. Esta proteção é correta.

Riscos residuais da distribuição:

- Se a data de entrada do membro for vazia/inválida, a condição de exclusão não é executada e todos os ensaios válidos são contabilizados. Isto não transforma concertos em ensaios, mas viola a regra temporal da assiduidade.
- Tipos não reconhecidos são ignorados, o que é seguro relativamente a concertos, mas pode ocultar variantes ou erros de dados sem aviso.
- A distribuição exclui também concertos anteriores à entrada. O `AGENTS.md` restringe explicitamente o denominador dos ensaios, mas não declara explicitamente se concertos pré-entrada devem ou não contar como atividades. Convém congelar esta decisão com um teste de caracterização antes de refatorar.

### Divergência do diagnóstico

O diagnóstico só conta linhas cujo tipo seja exatamente `Ensaio`, portanto não inclui concertos. Porém, para entrada inválida define `posteriorEntrada = false` e exclui todos os ensaios, enquanto a distribuição os inclui. O diagnóstico pode assim reportar um resultado diferente daquele que é distribuído.

## 10. Defeitos e riscos funcionais priorizados

### Críticos

1. **Resumo bloqueado por variável inexistente** — `generate_tables.js:99-105` declara `point` e usa `points`. A função falha na primeira célula de atividade elegível.
2. **Concertos no resumo de assiduidade** — ausência total de filtro por tipo em `gerarResumo`.

### Altos

1. **Data de entrada inválida inclui todos os ensaios na distribuição** — o `continue` temporal só é aplicado quando a entrada é uma data válida.
2. **Resumo ignora membros ativos e datas de entrada** — não cumpre as regras usadas na distribuição.
3. **Preservação manual vulnerável** — `folha.clear()` apaga primeiro toda a distribuição e recria os valores por chave de nome. Nomes duplicados normalizados, renomes ou alterações durante a execução podem atribuir/perder valores.
4. **Chaves de nome não são únicas** — dois membros ativos com o mesmo nome normalizado partilham o mesmo objeto de métricas e os mesmos valores manuais. Colunas duplicadas de presença também podem somar no mesmo registo.
5. **Linhas financeiras posteriores a um intervalo vazio são ignoradas** — pode produzir saldos incompletos apesar de `getLastRow()` encontrar dados abaixo.

### Médios

1. `converterNumero_` converte entradas inválidas silenciosamente em zero, ocultando erros financeiros.
2. `converterData_` aceita datas portuguesas impossíveis normalizadas pelo construtor `Date`.
3. Se não houver pontos, o fundo comum não é distribuído e não há aviso específico.
4. Se `validRehearsals` for zero, o resumo calcula `0/0` e funções de formatação tentam intervalos com zero colunas.
5. O resumo limpa a folha antes de terminar os cálculos; uma exceção deixa o resultado anterior perdido e a folha parcial/vazia.
6. A distribuição não usa `LockService`; execuções concorrentes ou uma edição manual durante a regeneração podem perder valores.
7. Fórmulas usam letras/posições fixas e separador `;`; futuras mudanças de colunas ou de locale podem quebrá-las.
8. Não há validação explícita de colunas de presença duplicadas ou de membros ativos sem coluna correspondente. Estes membros recebem métricas zero silenciosamente.

### Baixos / manutenção

1. `lastColumn` não é usado.
2. Configuração, convenções de nomes e parsing estão duplicados.
3. O relatório técnico fica abaixo da tabela mas o filtro é criado apenas até `linhaTotal`, uma escolha implícita que deve ser caracterizada.
4. `README.md` e `TODO.md` já não descrevem fielmente o estado do código.

## 11. Desempenho e chamadas a `SpreadsheetApp`

### `gerarResumo`

- Faz uma escrita remota por cabeçalho de data, uma por nome, uma por presença, uma por total e uma por percentagem: ordem de grandeza `O(membros × atividades)` chamadas ao serviço.
- Converte a mesma data novamente dentro do laço de cada membro, também `O(membros × atividades)` conversões, quando cada data poderia ser analisada uma vez.
- `customizeSheet` aplica zebra com `getRange(...).setBackground(...)` linha a linha.
- A folha é limpa antes de se construir/validar toda a matriz em memória.

Recomendação futura sem mudar os valores calculados: construir uma matriz 2D completa em memória e usar um único `setValues`; pré-calcular as atividades elegíveis; aplicar zebra com `applyRowBanding()` ou com uma regra/faixas agrupadas após testes de equivalência visual.

### Distribuição

Pontos positivos:

- Cada origem principal é lida em lote com `getValues()`.
- O laço `linhas de respostas × colunas de membros` ocorre em memória, sem chamadas de folha no interior.
- As linhas de distribuição e as fórmulas finais são escritas em lote.

Oportunidades:

- `getLastRow`, `getLastColumn` e `getRange` são chamados separadamente em cada leitor; isto é aceitável, mas um repositório de dados pode tornar o padrão uniforme.
- `encontrarIndiceCabecalho_` normaliza a linha inteira para cada cabeçalho procurado. Criar uma vez um mapa `cabeçalho normalizado -> índice` evita trabalho repetido e permite detetar duplicados.
- Sete fórmulas de totais são escritas com sete chamadas. Podem ser enviadas numa única linha com `setFormulas` depois de preservar exatamente as células vazias atuais.
- A formatação faz vários pedidos fixos e dez chamadas separadas a `setColumnWidth`; não cresce com os dados, mas pode ser agrupada quando a API o permitir.
- `SpreadsheetApp.getUi()` é obtido duas vezes em cada alerta; impacto negligenciável.

### Diagnóstico

- As duas folhas são lidas em lote.
- `Utilities.formatDate` é chamado por linha e os dados de debug são todos acumulados em memória. Para folhas muito grandes, isto aumenta tempo e memória, embora não faça escrita por linha.
- A folha de debug é sempre totalmente limpa e reescrita.

## 12. Organização modular proposta

Como todos os ficheiros Apps Script continuam a partilhar o namespace, os nomes abaixo representam camadas lógicas, não módulos JavaScript. Manter apenas os pontos de entrada públicos sem `_`; dar nomes únicos e consistentes a todos os auxiliares.

```text
00_config.js
  CONFIG (folhas, cabeçalhos, valores, pesos, datas e layout)

10_normalizacao.js
  normalizarTexto_, normalizarNome_, converterNumero_, converterData_,
  compararApenasData_, linhaCompletamenteVazia_, extrairNomeCabecalhoPresenca_

20_esquema.js
  criarMapaCabecalhos_, encontrarIndiceCabecalho_, validarCabecalhos_

30_leitura.js
  lerFolha_, lerMembros_, lerMovimentos_, lerRespostas_,
  lerValoresManuaisExistentes_

40_calculos_assiduidade.js
  calcularPontosEnsaio_, calcularAssiduidade_,
  calcularMetricasParticipacaoPuras_

50_calculos_financeiros.js
  calcularSaldo_, selecionarElegiveisPreWCG_, calcularRateioPreWCG_,
  calcularRateioFundoComum_, calcularValorFinal_

60_modelos_saida.js
  criarMatrizResumo_, criarLinhasDistribuicao_, criarFormulasDistribuicao_

70_escrita.js
  escreverResumo_, escreverDistribuicao_, escreverDiagnostico_

80_formatacao.js
  formatarResumo_, aplicarFormatacaoCondicionalResumo_,
  formatarDistribuicao_, formatarDiagnostico_

90_comandos.js
  gerarResumo, gerarDistribuicao, diagnosticarAssiduidade

99_menu.js
  onOpen
```

Princípios da organização:

- Uma única constante global de configuração.
- Cabeçalhos resolvidos uma vez por folha e passados como mapa.
- Leitores apenas obtêm dados e metadados da folha.
- Cálculos recebem arrays/objetos e não conhecem `SpreadsheetApp`.
- Escritores recebem matrizes prontas e não recalculam regras.
- Formatação não decide dados nem regras financeiras.
- `onOpen` permanece único.
- O diagnóstico chama o mesmo cálculo puro da distribuição e apenas acrescenta rastreabilidade por linha.
- Os números de coluna de saída são derivados da ordem de cabeçalhos/configuração; fórmulas são geradas a partir desses índices.

## 13. Plano de refatoração por fases

As fases 0–7 abaixo destinam-se a preservar o comportamento observável atual. Defeitos conhecidos não devem ser corrigidos silenciosamente durante movimentações de código. As correções de regras estão separadas na secção 14.

### Fase 0 — Baseline e dependências externas

1. Registar os acionadores simples e instaláveis do projeto.
2. Inspecionar botões, desenhos e macros associados a `gerarResumo`, `gerarDistribuicao` e `diagnosticarAssiduidade`.
3. Guardar cópias de amostras anonimizadas das seis folhas, incluindo linhas vazias intermédias, nomes semelhantes, datas inválidas e valores manuais.
4. Exportar resultados atuais de `Distribuição`, `Resumo` e `DEBUG_Assiduidade` quando as funções executarem.
5. Documentar locale da spreadsheet, fuso do script e fórmulas efetivamente aceites.
6. Criar uma lista de funções públicas que têm de conservar o nome.

Critério: mapa de dependências externas completo e baseline reproduzível. Nenhum código alterado.

### Fase 1 — Testes de caracterização puros

1. Criar casos para normalização, números, datas, cabeçalhos e comparação de dias.
2. Capturar o comportamento atual para linha financeira vazia, entrada inválida, tipo desconhecido, ausência de ensaios, total de pontos zero e nomes duplicados.
3. Capturar separadamente a divergência atual entre distribuição, resumo e diagnóstico.
4. Testar fórmulas e matriz de saída sem escrever numa spreadsheet real, sempre que possível.

Critério: os testes descrevem o comportamento atual, inclusive anomalias, e não pressupõem ainda a correção funcional.

### Fase 2 — Consolidar configuração sem alterar consumidores

1. Definir uma configuração canónica com todas as folhas, cabeçalhos, estados, pontos, peso e datas.
2. Manter temporariamente aliases compatíveis (`SHEETS`, `POINTS`, etc.) para evitar uma migração grande.
3. Migrar um consumidor de cada vez para a configuração canónica.
4. Remover aliases e constantes sem uso apenas depois de pesquisa estática e confirmação de dependências externas.

Critério: mesmos valores e nomes públicos; nenhuma constante global duplicada conceptualmente.

### Fase 3 — Extrair funções puras

1. Separar `calcularSaldo_(linhas, indices)` de `calcularSaldoMovimentos_`.
2. Separar descoberta de colunas, parsing das respostas e cálculo das métricas.
3. Extrair elegibilidade/rateio pré-WCG, rateio do fundo e composição final de linhas de `gerarDistribuicao`.
4. Criar uma função única de cálculo de assiduidade que possa produzir também os detalhes usados pelo diagnóstico.
5. Manter wrappers antigos durante a transição.

Critério: funções de domínio não referem `SpreadsheetApp`; resultados iguais ao baseline.

### Fase 4 — Separar leitura, escrita e formatação

1. Criar leitores que devolvam `{dados, cabecalhos, indices}`.
2. Construir matrizes completas de `Resumo` e `Distribuição` antes de tocar nas folhas de destino.
3. Fazer os escritores escreverem apenas matrizes e fórmulas já preparadas.
4. Mover toda a formatação de distribuição para o ficheiro de formatação.
5. Fazer o diagnóstico reutilizar os leitores e o cálculo puro comuns.

Critério: cada função com `SpreadsheetApp` limita-se a leitura, escrita, formatação ou UI; não contém fórmulas de negócio.

### Fase 5 — Otimizar chamadas à spreadsheet

1. Substituir escritas célula a célula do resumo por `setValues` em lote.
2. Analisar datas e linhas elegíveis uma única vez.
3. Escrever totais da distribuição em lote.
4. Criar uma única vez o mapa normalizado dos cabeçalhos.
5. Reduzir a formatação linha a linha, confirmando equivalência visual.

Critério: número de chamadas do resumo deixa de crescer com `membros × atividades`; conteúdo e formatação permanecem equivalentes.

### Fase 6 — Robustez operacional sem mudar regras

1. Validar todas as pré-condições antes de limpar folhas de destino.
2. Preparar dados completos antes da primeira escrita.
3. Avaliar `LockService` para impedir regenerações concorrentes e proteger os valores manuais.
4. Adicionar mensagens de erro com folha, linha e cabeçalho, sem converter silenciosamente falhas em alterações financeiras.
5. Manter `Form_Responses` estritamente read-only nos adaptadores.

Critério: uma falha de validação não apaga o resultado anterior; execuções concorrentes não perdem edições.

### Fase 7 — Limpeza e documentação

1. Remover wrappers, aliases e código morto apenas após confirmação dos usos externos.
2. Uniformizar nomes em português ou inglês; recomenda-se português por coerência com o domínio atual.
3. Renomear `testeEnsaios.js` para refletir que é diagnóstico, preservando o nome público da função se necessário.
4. Atualizar `README.md` e `TODO.md` para o estado real.
5. Fazer commits pequenos por camada: configuração, funções puras, leitores, escritores, formatação, comandos, documentação.

Critério: namespace global inventariado, sem duplicações, ficheiros com responsabilidade única e documentação atual.

## 14. Correções funcionais que exigem fase e aprovação separadas

Estas alterações alinham o sistema com `AGENTS.md`, mas mudam o comportamento atual e não devem ser misturadas numa refatoração neutra:

1. Corrigir `point`/`points` para tornar `gerarResumo` executável.
2. Fazer o resumo filtrar explicitamente `Tipo de Atividade = Ensaio`.
3. Fazer o resumo considerar apenas membros ativos e apenas ensaios desde a data de entrada.
4. Definir a política para entrada vazia/inválida: erro, exclusão do membro ou assiduidade zero. Nunca assumir silenciosamente todos os ensaios.
5. Decidir explicitamente se concertos anteriores à entrada contam como atividades.
6. Decidir se uma linha vazia intermédia termina a leitura financeira ou deve ser ignorada.
7. Definir o que acontece ao fundo comum quando o total de pontos é zero.
8. Definir tratamento de tipos/respostas desconhecidos e de membros sem coluna de presença.
9. Definir uma chave estável e única de membro em vez de depender apenas do nome normalizado.
10. Decidir se os valores manuais devem preservar apenas o valor numérico ou também fórmula, validação, nota e formato.
11. Validar datas impossíveis e números financeiros inválidos em vez de os normalizar silenciosamente.

Cada correção deve ter testes da regra desejada, um commit próprio e comparação de resultados antes/depois com dados de referência.

## 15. Ordem recomendada e controlo de risco

Ordem prática:

1. Baseline e testes.
2. Configuração única com compatibilidade temporária.
3. Extração de cálculos puros.
4. Unificação do cálculo usado por distribuição e diagnóstico.
5. Separação de I/O e escrita em lote.
6. Robustez e concorrência.
7. Correções funcionais autorizadas, uma a uma.
8. Remoção de código obsoleto e atualização documental.

Antes de cada fase:

- Confirmar `git status` e preservar alterações do utilizador.
- Não executar `clasp push` automaticamente.
- Não modificar `Form_Responses`.
- Comparar saldos, atividades, assiduidade, pontos, rateios e valores finais membro a membro.
- Comparar especificamente `Valor Individual` e `Valor Apoios` antes/depois.
- Confirmar que continua a existir exatamente um `onOpen`.

O primeiro objetivo da refatoração deve ser criar uma fonte única de verdade para a assiduidade e para a configuração. Só depois é seguro otimizar ou remover o código atualmente duplicado.
