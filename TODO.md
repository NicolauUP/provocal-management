# Próximos passos

## Migração manual das folhas

- [x] Renomear `Form_Responses` para `Presencas_Gerais`.
- [x] Criar `Ensaios_Sopranos`.
- [x] Criar `Ensaios_Contraltos`.
- [x] Criar `Ensaios_Tenores`.
- [x] Criar `Ensaios_Baixos`.
- [x] Confirmar os cabeçalhos `Timestamp` e `Data` nas folhas de naipe.
- [x] Confirmar também `Tipo de atividade` em `Presencas_Gerais`.
- [x] Validar a nova `Distribuição` e os valores manuais.
- [x] Eliminar manualmente a antiga folha de movimentos pré-WCG só depois da validação.

## Melhorias futuras

- [ ] Adicionar testes automatizados executáveis no ambiente Apps Script.
- [ ] Adicionar filtros ou gráficos aos mapas mensais de assiduidade.
- [ ] Avaliar uma chave estável de membro que não dependa do nome.
- [ ] Confirmar botões, macros e acionadores externos ao repositório.

## Consulta e publicação para o coro

- [x] Criar uma spreadsheet separada, destinada apenas à consulta dos membros.
- [x] Adicionar a ação manual `Publicar distribuição` ao menu da spreadsheet privada.
- [x] Restringir a publicação ao email configurado do responsável.
- [x] Pedir confirmação antes de cada publicação.
- [x] Copiar apenas valores e formatação, sem fórmulas ou ligações à spreadsheet privada.
- [x] Publicar a tabela de distribuição sem o resumo técnico interno.
- [x] Publicar os mapas mensais de assiduidade aprovados.
- [x] Registar no ficheiro público a data e hora da última atualização.
- [ ] Manter os membros do coro apenas com permissões de leitura.
- [x] Criar o protótipo inicial de Web App de consulta e validar os dados apresentados.
- [x] Criar um projeto Apps Script separado para o portal de consulta autenticada.
- [x] Criar uma fonte privada de dados do portal, separada da spreadsheet pública.
- [x] Adicionar a coluna `Email` e validar emails preenchidos antes de publicar no portal.
- [ ] Preencher o `Email` de todos os membros ativos que devem poder consultar o portal.
- [x] Exigir login Google e verificar a associação email → membro.
- [x] Limitar cada membro à consulta dos seus próprios dados.
- [x] Mostrar assiduidade acumulada desde 01/05/2026, atividades, valores e evolução mensal.
- [x] Garantir que o portal apresenta apenas a última versão publicada manualmente.
- [x] Não incluir no portal ações para editar dados, atualizar cálculos ou publicar.
- [x] Testar o portal com uma conta Google externa.
- [x] Implementar e partilhar o URL do portal autenticado.
- [x] Fixar a estrutura aprovada do portal; alterações estruturais exigem pedido explícito.
- [x] Remover o código do portal antigo por nome/mês do projeto administrativo.
- [ ] Arquivar a implantação antiga do portal de consulta por nome/mês no editor Apps Script.
- [ ] Confirmar que a spreadsheet pública é partilhada apenas como leitura com os membros.

## Registo privado de presenças por fotografia

> Iniciar apenas depois de o portal de consulta autenticada estar concluído e validado.

- [ ] Criar um projeto Web App separado e privado para registo de presenças.
- [ ] Restringir o acesso a responsáveis autorizados por email.
- [ ] Definir uma fonte de registos manuais sem escrever diretamente nas respostas dos Forms.
- [ ] Adicionar uma coluna de foto/URL de foto em `Membros`, com consentimento dos membros.
- [ ] Mostrar membros através de fotografias e permitir marcar `Presente`, `Atraso` ou `Falta`.
- [ ] Nos concertos, permitir apenas `Presente` ou `Falta`.
- [ ] Permitir escolher data, tipo de atividade e fonte geral/naipe.
- [ ] Unificar ensaios gerais, os quatro naipes e concertos numa única aplicação.
- [ ] Ao selecionar um naipe, apresentar apenas os membros aplicáveis.
- [ ] Incluir ações rápidas: marcar todos presentes/falta e pesquisa por nome.
- [ ] Mostrar pré-visualização e exigir confirmação antes de guardar.
- [ ] Impedir duplicados para a mesma atividade, data e fonte.
- [ ] Registar auditoria: responsável, data/hora e alterações.
- [ ] Criar uma fonte normalizada de registos manuais para não escrever diretamente nas respostas dos Forms.
- [ ] Combinar os registos manuais com os Forms no cálculo de assiduidade, sem duplicar presenças.
- [ ] Integrar os registos manuais no cálculo atual sem alterar regras de negócio.

## Portal privado de gestão do coro

> Evoluir depois de a aplicação de registo de presenças estar validada.

- [ ] Criar um portal separado, acessível apenas a responsáveis/"managers" autorizados.
- [ ] Definir funções e emails autorizados: publicação, presenças, movimentos e administração.
- [ ] Criar um módulo para adicionar movimentos do coro com validação de data, descrição, receita/despesa e valor.
- [ ] Gravar movimentos através de uma fonte controlada, preservando o cálculo financeiro atual.
- [ ] Mostrar histórico recente de movimentos e permitir correções auditáveis.
- [ ] Separar permissões de consulta, registo de presenças, registo financeiro e publicação.
- [ ] Centralizar ações de gestão: presenças, movimentos, atualização da distribuição e publicação.
- [ ] Registar auditoria de todas as alterações: responsável, data/hora, ação e valores antes/depois quando aplicável.
- [ ] Adicionar confirmação explícita a ações com impacto financeiro ou publicação pública.
