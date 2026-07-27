# Plano de evolução

## Estado atual

O cálculo de participação usa uma análise comum para:

- `Presencas_Gerais`;
- `Ensaios_Sopranos`;
- `Ensaios_Contraltos`;
- `Ensaios_Tenores`;
- `Ensaios_Baixos`.

As folhas de naipe são opcionais durante a migração e produzem avisos claros quando
estão ausentes. Uma folha existente com cabeçalhos obrigatórios em falta produz erro.

A mesma análise alimenta:

- `gerarDistribuicao`;
- `gerarResumo`;
- `diagnosticarAssiduidade`.

## Regras consolidadas

- Apenas membros ativos entram nos resultados.
- Concertos contam como atividades apenas desde 01/05/2026.
- Concertos nunca entram na assiduidade.
- Ensaios gerais e de naipe usam `Presente = 5`, `Atraso = 4` e `Falta = 0`.
- Respostas vazias em ensaios elegíveis entram no denominador e valem zero.
- Ensaios anteriores à entrada do membro são excluídos.
- O peso das atividades permanece em `0.75`.
- O fundo comum continua a resultar de receitas menos despesas em `Movimentos`.
- `Valor Individual` e `Valor Apoios` são preservados ao regenerar.
- `Valor Final` é fundo comum mais os dois valores manuais.

## Próximas melhorias seguras

1. Criar testes de caracterização executáveis no Apps Script.
2. Centralizar por completo a configuração atualmente ainda dividida entre
   `SHEETS` e `DISTRIBUICAO_CONFIG`.
3. Derivar todos os índices de saída dos cabeçalhos.
4. Avaliar `LockService` para proteger regenerações concorrentes.
5. Validar datas impossíveis e números financeiros inválidos.
6. Substituir a correspondência por nome por uma chave estável de membro.
7. Confirmar acionadores, botões e macros antes de remover qualquer ponto de entrada.

Não executar `clasp push` automaticamente e não alterar dados reais durante testes.
