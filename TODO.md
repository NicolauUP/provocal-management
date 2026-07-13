# 1. Folha de Distribuição

- Estruturar a folha, com várias colunas, para o calculo e divisão do dinheiro do coro pelos coristas. 

- Divisão do dinheiro (adicionar em colunas): 
    - Valor Caixa Antes WCG: que divisão soma (coluna (Receita) da folha Movimento_PREWCG - coluna (Despesa) da folha Movimento_PREWCG) pelos membros que estavam ativos antes de 01/05/2026  (na folha Membros tem a coluna: Entrada e a coluna Estado (Ativo ou Inativo))
    - Valor Fundo Comum: 
        - Calculo do fundo comum: SOMA da folha movimentos (coluna Receita - coluna Despesa) -> tens de fazer um loop ate encontrares um linha vazia para fazer soma
        - Calculo dos pontos: Folha Form_Responses, filtrar na coluna Tipo de Atividade == 'Concerto', e contar para cada membro (atencao que esta tabela na linha 1 o nome dos membros está nas colunas: ou seja, a coluna C linha 1 Registo de Presenças [Inês Regina], coluna D linha 1 Registo de Presenças [Catarina Lopes]), quantas do tipo Concerto tem "Presente". 
        - Soma de pontos de todos membros
        - Valor do fundo comum de cada membro: Nº de pontos * Soma da folha movimentos / nº total de pontos de todos os membros
    - Valor individual: coluna vazia
    - Valor Apoios: coluna vazia
    - Valor final = Antes WCG + Fundo Comum + Individual + Apoios

# 2. Folha despesas

    - Ter um folha Antes WCG (673.43€)
    - Ter uma folha com os Movimentos (concertos, despesas...) 
