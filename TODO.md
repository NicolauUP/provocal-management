# Próximos passos

## Migração manual das folhas

- [ ] Renomear `Form_Responses` para `Presencas_Gerais`.
- [ ] Criar `Ensaios_Sopranos`.
- [ ] Criar `Ensaios_Contraltos`.
- [ ] Criar `Ensaios_Tenores`.
- [ ] Criar `Ensaios_Baixos`.
- [ ] Confirmar os cabeçalhos `Timestamp` e `Data` nas folhas de naipe.
- [ ] Confirmar também `Tipo de atividade` em `Presencas_Gerais`.
- [ ] Validar a nova `Distribuição` e os valores manuais.
- [ ] Eliminar manualmente a antiga folha de movimentos pré-WCG só depois da validação.

## Melhorias futuras

- [ ] Adicionar testes automatizados executáveis no ambiente Apps Script.
- [ ] Adicionar filtros ou gráficos aos mapas mensais de assiduidade.
- [ ] Avaliar uma chave estável de membro que não dependa do nome.
- [ ] Confirmar botões, macros e acionadores externos ao repositório.

## Área pública do coro

- [x] Criar uma spreadsheet separada, destinada apenas à consulta dos membros.
- [x] Adicionar a ação manual `Publicar distribuição` ao menu da spreadsheet privada.
- [x] Restringir a publicação ao email configurado do responsável.
- [x] Pedir confirmação antes de cada publicação.
- [x] Copiar apenas valores e formatação, sem fórmulas ou ligações à spreadsheet privada.
- [x] Publicar a tabela de distribuição sem o resumo técnico interno.
- [x] Publicar os mapas mensais de assiduidade aprovados.
- [x] Registar no ficheiro público a data e hora da última atualização.
- [ ] Manter os membros do coro apenas com permissões de leitura.
- [x] Criar um Web App público e exclusivamente de leitura.
- [x] Permitir selecionar livremente o nome de qualquer membro.
- [x] Permitir selecionar o mês a consultar.
- [x] Mostrar concertos, ensaios gerais e ensaios de naipe no resumo individual.
- [x] Mostrar estados, pontos `5/4/0`, totais e percentagem de assiduidade.
- [x] Garantir que o portal apresenta apenas a última versão publicada manualmente.
- [x] Não incluir no portal ações para editar dados, atualizar cálculos ou publicar.
- [ ] Criar um projeto Apps Script separado para o portal de consulta.
- [ ] Criar uma fonte privada de dados do portal, separada da spreadsheet pública.
- [ ] Adicionar e validar a coluna `Email` na folha `Membros`.
- [ ] Exigir login Google e verificar a associação email → membro.
- [ ] Limitar cada membro à consulta dos seus próprios dados.
- [ ] Testar o portal com contas Google externas antes de o partilhar.
- [ ] Arquivar a implantação atual antes de disponibilizar o portal autenticado.
- [ ] Implantar e partilhar o URL do portal autenticado.

## Registo privado de presenças por fotografia

> Iniciar apenas depois de o portal de consulta autenticada estar concluído e validado.

- [ ] Criar um projeto Web App separado e privado para registo de presenças.
- [ ] Restringir o acesso a responsáveis autorizados por email.
- [ ] Definir uma fonte de registos manuais sem escrever diretamente nas respostas dos Forms.
- [ ] Adicionar uma coluna de foto/URL de foto em `Membros`, com consentimento dos membros.
- [ ] Mostrar membros através de fotografias e permitir marcar `Presente`, `Atraso` ou `Falta`.
- [ ] Nos concertos, permitir apenas `Presente` ou `Falta`.
- [ ] Permitir escolher data, tipo de atividade e fonte geral/naipe.
- [ ] Incluir ações rápidas: marcar todos presentes/falta e pesquisa por nome.
- [ ] Mostrar pré-visualização e exigir confirmação antes de guardar.
- [ ] Impedir duplicados para a mesma atividade, data e fonte.
- [ ] Registar auditoria: responsável, data/hora e alterações.
- [ ] Integrar os registos manuais no cálculo atual sem alterar regras de negócio.
