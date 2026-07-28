# Portal individual do coro

Este é um projeto Apps Script separado do projeto administrativo. O seu ID é
`1QM-4bYL-7N0gcf4kLgnNOBju57FEKLvp4OpIiW17XvKcSmy2xl8FXzWb`.

O portal irá ler apenas o snapshot privado na spreadsheet
`1UwlcSqP5xDMIaSGmxwvRlu8R3jsp3U2PYHAl1-Nl4ME`. Nunca deve receber funções
de cálculo, publicação ou edição de presenças.

## Antes de implementar

1. Adicionar e preencher a coluna `Email` na folha `Membros` do projeto
   administrativo, com valores únicos para membros ativos.
2. Criar um OAuth Client ID do tipo **Web application** para o login Google.
   Em **Authorized redirect URIs**, registar:
   o URL `/exec` da implementação web do portal.
3. Guardar o *client secret* apenas nas Script Properties deste projeto, com o
   nome `GOOGLE_OAUTH_CLIENT_SECRET`. Nunca o colocar no código, no clasp ou
   numa mensagem.

Enquanto esses passos não estiverem concluídos, o portal não autentica nem
expõe dados de membros.

## Sincronização manual

Quando chegar a altura, executar `clasp push` a partir desta pasta. Nunca
executar esse comando sem rever os ficheiros e criar uma implementação nova no
editor Apps Script.
