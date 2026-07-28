# Coro Assiduidade

Sistema de gestão de assiduidade para coros, desenvolvido com Google Forms, Google Sheets e Google Apps Script.

O objetivo é automatizar o registo de presenças, calcular assiduidade e pontuações, gerir membros e servir de base para futuras funcionalidades administrativas do coro.

---

## Funcionalidades

### Atualmente

- Geração automática de tabela resumo
- Conversão de presenças em pontos
- Cálculo automático da assiduidade
- Integração de ensaios gerais e de naipe
- Contagem de concertos desde 01/05/2026
- Distribuição proporcional do fundo comum
- Mapas mensais de assiduidade com pontos por ensaio
- Publicação manual da distribuição e dos mapas numa spreadsheet de consulta
- Portal Web de consulta individual por membro e mês

## Portal público

O portal lê apenas o último snapshot criado por `Publicar distribuição`.
Para o disponibilizar:

1. Enviar o código manualmente com `clasp push`.
2. No editor do Apps Script, escolher `Implementar` → `Nova implementação`.
3. Selecionar `Aplicação Web`.
4. Executar como o proprietário do projeto.
5. Escolher quem pode aceder e concluir a implementação.
6. Partilhar o URL terminado em `/exec`.

Depois de alterar o código do portal, é necessário criar uma nova versão da
implementação. O portal não contém operações de escrita ou publicação.
- Formatação automática da folha de resumo
- Conditional Formatting
- Desenvolvimento local com VS Code + clasp
- Controlo de versões com Git

---

## Roadmap

### Gestão de membros

- [ ] Base de dados de membros
- [ ] Data de entrada
- [ ] Data de saída
- [ ] Naipe
- [ ] Estado (Ativo / Suspenso / Antigo)

### Assiduidade

- [x] Ensaios gerais
- [x] Ensaios de naipe
- [x] Concertos
- [ ] Atividades extraordinárias
- [ ] Diferentes pesos por atividade

### Estatísticas

- [ ] Ranking de assiduidade
- [ ] Estatísticas por mês
- [ ] Estatísticas anuais
- [ ] Evolução temporal

### Financeiro

- [x] Distribuição automática de verbas
- [x] Sistema de pontos ponderados
- [ ] Exportação dos resultados

### Exportação

- [ ] PDF
- [ ] Excel
- [ ] Relatório mensal

### Backend

- [ ] Base de dados SQL
- [ ] API
- [ ] Dashboard Web

---

## Estrutura do projeto

```
.
├── Code.js
├── appsscript.json
├── .clasp.json
└── README.md
```

(Esta estrutura será reorganizada à medida que o projeto evoluir.)

---

## Tecnologias

- Google Forms
- Google Sheets
- Google Apps Script
- clasp
- Git
- GitHub
- Visual Studio Code

---


## Objetivos

- Automatizar tarefas repetitivas da comissão do coro.
- Garantir consistência e transparência no cálculo da assiduidade.
- Facilitar a gestão de membros.
- Criar uma plataforma extensível para futuras funcionalidades administrativas.

---

## Licença

Projeto privado.
