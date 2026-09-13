# 💰 Julius Finance System

<p align="center">
  Sistema de gestão financeira pessoal desenvolvido com Google Sheets e Google Apps Script.
</p>

<p align="center">
  <strong>Controle de hoje. Liberdade de amanhã.</strong>
</p>

---

## 📌 Sobre o projeto

O **Julius Finance System** é um sistema de gestão financeira pessoal criado sobre o ecossistema do Google Sheets, utilizando **Google Apps Script e JavaScript** para transformar uma planilha convencional em uma experiência mais próxima de uma aplicação financeira.

O projeto nasceu da necessidade de centralizar e automatizar o controle de:

- receitas;
- despesas fixas;
- despesas variáveis;
- cartões de crédito;
- parcelamentos;
- planejamento mensal;
- reembolsos;
- pagamentos de reembolsos;
- categorias;
- responsáveis;
- importação de faturas;
- indicadores financeiros.

Além dos cálculos financeiros, o sistema possui automações, modais, configurações centralizadas, filtros por competência e um Dashboard visual para acompanhamento mensal.

A identidade visual foi inspirada no conceito de economia e controle financeiro associado ao personagem Julius, de *Todo Mundo Odeia o Chris*, mas aplicada a uma interface moderna e voltada à experiência de produto.

> Este é um projeto independente e não possui vínculo oficial com a série, seus produtores, emissoras ou detentores de direitos.

---

# ✨ Funcionalidades

## 📊 Dashboard

O Dashboard concentra os principais indicadores da competência selecionada.

Atualmente apresenta:

- receita mensal;
- despesas mensais;
- saldo da competência;
- percentual de economia;
- fluxo financeiro;
- gastos por categoria;
- insights financeiros automáticos;
- seletor de competência.

O Dashboard utiliza os dados registrados nos demais módulos e desconsidera automaticamente valores classificados como reembolso de terceiros nos indicadores financeiros principais.

### 📈 Fluxo financeiro

Apresenta graficamente a evolução acumulada das:

- receitas;
- despesas.

Isso permite acompanhar como entradas e gastos evoluem durante a competência.

### 🍩 Gastos por categoria

Os gastos são agrupados automaticamente por categoria.

O gráfico considera todas as categorias com movimentação no período selecionado, permitindo visualizar onde o dinheiro está sendo utilizado.

---

## 👀 Julius Diz

O **Julius Diz** funciona como uma camada de interpretação dos dados financeiros.

Com base nos números da competência, o sistema pode identificar situações como:

- despesas superiores às receitas;
- saldo mensal positivo ou negativo;
- baixo percentual de economia;
- comprometimento da renda com parcelamentos;
- necessidade de atenção ao orçamento.

A proposta é não apresentar apenas números, mas transformar os dados em informações úteis para tomada de decisão.

---

## 💰 Receitas

Módulo destinado ao controle das entradas financeiras.

Permite registrar informações como:

- data;
- descrição;
- categoria;
- valor;
- tipo;
- recorrência.

Os lançamentos alimentam automaticamente os indicadores mensais do sistema.

---

## 📋 Despesas Fixas

Controle de despesas recorrentes e compromissos financeiros fixos.

Os registros podem ser associados a responsáveis e categorias, permitindo que o sistema diferencie despesas próprias de valores relacionados a terceiros.

---

## 🛒 Despesas Variáveis

Controle dos gastos realizados ao longo do mês.

Os registros podem conter:

- data;
- categoria;
- descrição;
- valor;
- forma de pagamento;
- cartão;
- responsável.

As despesas são consideradas automaticamente nos cálculos da competência correspondente.

---

## 📅 Planejamento Mensal

Área destinada ao acompanhamento financeiro da competência.

Centraliza informações utilizadas para comparar planejamento e realização financeira e serve de apoio para outros módulos do sistema.

O planejamento também reúne informações relacionadas ao uso dos cartões e aos compromissos financeiros do mês.

---

## 💳 Parcelamentos

O módulo de parcelamentos acompanha compras e compromissos divididos em múltiplas parcelas.

Entre os dados controlados estão:

- categoria;
- produto ou serviço;
- valor total;
- quantidade de parcelas;
- parcela atual;
- parcelas restantes;
- valor da parcela;
- cartão;
- responsável;
- observação;
- data da primeira parcela;
- última competência;
- vigência na competência selecionada;
- status.

O sistema calcula automaticamente a evolução do parcelamento de acordo com a competência selecionada.

Uma parcela continua ativa em sua última competência e deixa de impactar os meses seguintes após o encerramento.

---

## 💳 Cartões

Área destinada ao gerenciamento dos cartões utilizados no sistema.

Os cartões cadastrados podem ser utilizados nos demais módulos e fazem parte da estrutura de análise financeira.

A evolução do Dashboard prevê indicadores específicos para:

- participação de cada cartão nos gastos;
- ranking de utilização;
- comprometimento da renda;
- distribuição das despesas entre cartões.

---

## 💸 Reembolsos

O sistema possui uma estrutura própria para controle de valores relacionados a terceiros.

Responsáveis podem ser classificados como:

- **Principal**
- **Terceiro / Reembolso**

Quando um lançamento pertence a alguém classificado como **Terceiro / Reembolso**, o valor não é considerado como despesa pessoal nos indicadores principais.

Esses registros são direcionados ao controle de reembolsos.

O módulo acompanha:

- data;
- responsável;
- origem;
- descrição;
- valor devido;
- valor pago;
- saldo;
- status;
- observação;
- competência;
- origem técnica do lançamento.

Os registros históricos permanecem armazenados mesmo quando outra competência é selecionada.

---

## 💵 Pagamentos de Reembolsos

O sistema possui um fluxo específico para registrar valores recebidos de terceiros.

O usuário pode informar:

- data do recebimento;
- responsável;
- valor recebido;
- forma de pagamento;
- observação;
- competência que deseja quitar.

O pagamento é distribuído automaticamente entre as dívidas pendentes daquela competência.

São suportados:

- pagamentos parciais;
- pagamentos integrais;
- distribuição de um pagamento entre múltiplas pendências;
- histórico de recebimentos;
- rastreamento das alocações realizadas.

Uma estrutura técnica separada mantém a relação entre pagamentos e débitos, permitindo auditoria do histórico.

---

## 📄 Importação de faturas

O Julius Finance possui um módulo para importação de faturas em PDF.

O fluxo foi desenvolvido para permitir:

1. seleção do arquivo;
2. leitura do conteúdo;
3. processamento dos lançamentos;
4. identificação do cartão;
5. revisão dos dados;
6. edição antes da importação;
7. confirmação pelo usuário;
8. inclusão dos lançamentos.

O sistema também possui suporte a OCR para documentos que não disponibilizam o conteúdo diretamente como texto.

Nenhum lançamento deve ser importado sem confirmação do usuário.

Também existem mecanismos para reduzir importações duplicadas.

---

## ⚙️ Configurações

O sistema possui uma área centralizada de configurações.

Ela permite administrar:

### 👤 Responsáveis

Define quem pode ser associado aos lançamentos financeiros.

Cada responsável pode possuir um tipo que determina como seus valores serão tratados pelo sistema.

### 🏷️ Categorias

Centraliza as categorias disponíveis para classificação das movimentações.

### 💳 Cartões

Mantém a relação de cartões disponíveis nos módulos financeiros.

### 💵 Formas de pagamento

Centraliza as formas de pagamento utilizadas nos registros.

As configurações ficam armazenadas em uma estrutura técnica e são utilizadas pelos demais módulos do sistema.

---

# 🧠 Regras financeiras

Uma das principais características do Julius Finance é separar a interface das regras de negócio.

## Competência

A competência selecionada no Dashboard funciona como referência central para os módulos que dependem do período.

Ao alterar o mês, o sistema pode atualizar:

- indicadores;
- parcelamentos vigentes;
- despesas;
- reembolsos;
- gráficos;
- análises mensais.

---

## Responsáveis e reembolsos

O sistema não depende de nomes específicos para identificar terceiros.

A classificação é feita através da configuração do responsável.

Isso permite cadastrar novos responsáveis sem alterar manualmente as fórmulas ou regras principais.

---

## Parcelamentos

A vigência de cada parcelamento é calculada a partir da:

- data da primeira parcela;
- quantidade total de parcelas;
- competência selecionada.

Dessa forma, o sistema consegue determinar automaticamente se determinado parcelamento deve ou não impactar aquele mês.

---

# 🤖 Automações

O projeto utiliza Google Apps Script para automatizar diferentes comportamentos.

Entre eles:

- atualização por competência;
- sincronização de reembolsos;
- controle de parcelamentos;
- atualização de filtros;
- abertura de modais;
- importação de faturas;
- configurações;
- processamento de pagamentos;
- atualização visual.

Algumas rotinas utilizam gatilhos instaláveis do Google Apps Script para responder às alterações realizadas na planilha.

---

# 🎨 Interface e experiência

Embora utilize Google Sheets como base, o projeto procura reduzir a sensação de utilização de uma planilha tradicional.

A interface utiliza:

- Dashboard personalizado;
- cards;
- indicadores;
- gráficos;
- hierarquia visual;
- cores padronizadas;
- modais HTML;
- menus personalizados;
- validações;
- áreas técnicas ocultas;
- mensagens de feedback.

A identidade visual utiliza uma paleta inspirada no conceito do Julius Finance, com tons de:

- azul-petróleo;
- bege;
- caramelo;
- marrom;
- verde;
- vermelho para alertas.

O objetivo é manter a personalidade temática sem perder uma aparência séria e moderna.

---

# 🏗️ Arquitetura

O sistema utiliza o próprio Google Sheets como camada de dados e interface, enquanto o Google Apps Script concentra as regras de negócio e automações.

```text
Usuário
   │
   ▼
Google Sheets
   │
   ├── Dashboard
   ├── Receitas
   ├── Planejamento
   ├── Despesas
   ├── Parcelamentos
   ├── Cartões
   └── Reembolsos
   │
   ▼
Google Apps Script
   │
   ├── Regras de negócio
   ├── Automações
   ├── Sincronizações
   ├── Modais HTML
   ├── Importação de PDF
   └── Camada visual
   │
   ▼
Estruturas técnicas
```

Algumas informações auxiliares são armazenadas em abas técnicas que não fazem parte da navegação normal do usuário.

---

# 🗂️ Estrutura do projeto

A organização do código é dividida por responsabilidade.

```text
controle-financeiro-automatizado/
│
├── Código
├── Automacoes
├── Configuracoes
├── DashboardVisual
├── ReceitasVisual
├── PlanejamentoVisual
├── DespesasFixasVisual
├── DespesasVariaveis
├── DespesasVariaveisVisual
├── ParcelamentosVisual
├── CartoesVisual
├── Reembolsos
├── ReembolsosFiltro
├── ReembolsosVisual
├── PagamentosReembolsos
├── MetasVisual
├── ImportarFatura
│
├── NovoParcelamento.html
├── ImportarFaturaa.html
├── ConfiguracoesPainel.html
├── RegistrarPagamentoReembolso.html
│
├── appsscript.json
├── .clasp.json
├── .claspignore
├── .gitignore
└── README.md
```

> Os arquivos `.gs` do Google Apps Script podem aparecer localmente como `.js` ao utilizar `clasp`.

---

# 🛠️ Tecnologias

O projeto utiliza principalmente:

| Tecnologia | Utilização |
|---|---|
| Google Sheets | Interface e armazenamento dos dados |
| Google Apps Script | Regras de negócio e automações |
| JavaScript | Desenvolvimento das funcionalidades |
| HTML/CSS | Modais e interfaces auxiliares |
| Google Drive API | Processamento e integração de arquivos |
| clasp | Sincronização entre Apps Script e ambiente local |
| Git | Controle de versão |
| GitHub | Versionamento e documentação do projeto |
| VS Code | Ambiente local de desenvolvimento |

---

# 💻 Desenvolvimento local

O projeto pode ser sincronizado com o Google Apps Script utilizando `clasp`.

É necessário possuir:

- Node.js;
- npm;
- clasp;
- Git;
- acesso ao projeto correspondente no Google Apps Script.

Instalação global do clasp:

```bash
npm install -g @google/clasp
```

Autenticação:

```bash
clasp login
```

---

## ⬇️ Apps Script → ambiente local

Quando as alterações forem realizadas diretamente no editor do Google Apps Script:

```bash
clasp pull
```

Depois é possível revisar as mudanças localmente antes de versioná-las.

---

## ⬆️ Ambiente local → Apps Script

Quando as alterações forem realizadas no VS Code:

```bash
clasp push
```

Antes de executar o comando, é recomendado verificar as alterações para evitar sobrescrever uma versão mais recente existente no Apps Script.

---

# 🌿 Versionamento

Fluxo básico utilizado no projeto:

```bash
git status
git add .
git commit -m "descrição da alteração"
git push origin main
```

O histórico do Git registra a evolução das funcionalidades e da interface.

---

# 🔐 Segurança e privacidade

Este repositório deve armazenar **somente o código necessário para demonstrar e desenvolver o sistema**.

Não devem ser versionados:

- dados financeiros reais;
- nomes e informações pessoais utilizados na planilha real;
- PDFs de faturas;
- documentos pessoais;
- tokens;
- credenciais;
- chaves privadas;
- arquivos de autenticação;
- IDs privados desnecessários;
- dados utilizados apenas em ambiente de produção.

O arquivo `.clasp.json` contém informações específicas da conexão com o projeto Apps Script e deve permanecer fora do repositório público através do `.gitignore`.

Antes de publicar exemplos ou screenshots, os dados devem ser fictícios ou anonimizados.

---

# 🚀 Roadmap

## Concluído

- [x] Estrutura financeira principal
- [x] Controle de receitas
- [x] Despesas fixas
- [x] Despesas variáveis
- [x] Planejamento mensal
- [x] Controle de parcelamentos
- [x] Gestão de cartões
- [x] Sistema de responsáveis
- [x] Separação de terceiros/reembolsos
- [x] Controle de reembolsos
- [x] Pagamentos parciais de reembolsos
- [x] Histórico de pagamentos
- [x] Configurações centralizadas
- [x] Importação de faturas
- [x] Leitura de PDF/OCR
- [x] Interface visual temática
- [x] Dashboard por competência
- [x] Insights automáticos — Julius Diz
- [x] Gráfico de fluxo financeiro
- [x] Gráfico de gastos por categoria
- [x] Integração com Git e GitHub
- [x] Ambiente local com clasp

## Em desenvolvimento

- [ ] Dashboard avançado
- [ ] Ranking e análise dos cartões
- [ ] Gráfico de participação dos cartões
- [ ] Planejado x realizado no Dashboard
- [ ] Indicadores de reembolsos no Dashboard
- [ ] Próximos compromissos
- [ ] Refinamento dos gráficos
- [ ] Refinamento responsivo do layout
- [ ] Área técnica dedicada para dados do Dashboard
- [ ] Padronização final da experiência visual

## Evoluções futuras

- [ ] Transformação da interface em experiência ainda mais próxima de uma aplicação
- [ ] Novos indicadores financeiros
- [ ] Relatórios financeiros
- [ ] Histórico comparativo entre competências
- [ ] Análises de tendência
- [ ] Maior personalização por usuário
- [ ] Preparação da arquitetura para possível distribuição/comercialização

---

# 🎯 Objetivo do projeto

Além de resolver uma necessidade real de organização financeira, o Julius Finance System funciona como projeto prático de desenvolvimento.

Durante sua construção são trabalhados conceitos de:

- JavaScript;
- automação;
- manipulação de dados;
- regras de negócio;
- arquitetura de software;
- integração com APIs;
- interfaces HTML;
- Git;
- GitHub;
- desenvolvimento iterativo;
- experiência do usuário;
- design de produto.

O projeto continuará evoluindo à medida que novas funcionalidades e melhorias forem implementadas.

---

# 👩‍💻 Autoria

Desenvolvido por **Sara Souza**.

GitHub: **@saracodevs**

Projeto desenvolvido como parte da evolução prática em programação, automação e construção de produtos digitais.

---

<p align="center">
  <strong>JULIUS FINANCE SYSTEM</strong><br>
  Controle de hoje. Liberdade de amanhã.
</p>