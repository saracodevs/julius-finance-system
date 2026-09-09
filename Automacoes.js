/**
 * ============================================================
 * ⚙️ AUTOMAÇÕES — JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * IMPORTANTE:
 * Existe apenas UM onEdit(e) em todo o projeto.
 *
 * Dashboard B7:
 * 1. Atualiza Despesas Variáveis
 * 2. Ordena Parcelamentos
 * 3. Sincroniza Reembolsos
 * 4. Aplica automaticamente a visão da competência atual
 *
 * Despesas Variáveis / Parcelamentos:
 * sincronizam Reembolsos quando envolverem terceiros.
 */


/**
 * ============================================================
 * 🚀 ÚNICO ONEDIT DO SISTEMA
 * ============================================================
 */

function onEdit(e) {

  if (
    !e ||
    !e.range
  ) {
    return;
  }


  const aba =
    e.range.getSheet();


  const nomeAba =
    aba.getName();


  /**
   * ==========================================================
   * 📅 DASHBOARD — MUDANÇA DE COMPETÊNCIA
   * ==========================================================
   */

  if (
    nomeAba === '📊 Dashboard' &&
    e.range.getA1Notation() === 'B7'
  ) {

    atualizarSistemaPorCompetencia_();

    return;

  }


  /**
   * ==========================================================
   * 🛒 DESPESAS VARIÁVEIS
   * ==========================================================
   */

  if (
    nomeAba === '🛒 Desp.Variáveis'
  ) {

    processarEdicaoDespesasVariaveis_(e);

    return;

  }


  /**
   * ==========================================================
   * 💳 PARCELAMENTOS
   * ==========================================================
   */

  if (
    nomeAba === '💳 Parcelamentos'
  ) {

    processarEdicaoParcelamentos_(e);

  }

}


/**
 * ============================================================
 * 📅 ATUALIZA TODO O SISTEMA PELA COMPETÊNCIA
 * ============================================================
 */

function atualizarSistemaPorCompetencia_() {

  SpreadsheetApp.flush();


  /**
   * ----------------------------------------------------------
   * 🛒 DESPESAS VARIÁVEIS
   * ----------------------------------------------------------
   */

  try {

    if (
      typeof filtrarDespesasVariaveisPorMes ===
      'function'
    ) {

      filtrarDespesasVariaveisPorMes();

    }

  } catch (erro) {

    console.log(
      'Erro ao filtrar Despesas Variáveis: ' +
      erro
    );

  }


  /**
   * ----------------------------------------------------------
   * 💳 PARCELAMENTOS
   * ----------------------------------------------------------
   */

  try {

    if (
      typeof ordenarParcelamentos ===
      'function'
    ) {

      ordenarParcelamentos();

    }

  } catch (erro) {

    console.log(
      'Erro ao ordenar Parcelamentos: ' +
      erro
    );

  }


  SpreadsheetApp.flush();


  /**
   * ----------------------------------------------------------
   * 💸 REEMBOLSOS
   * ----------------------------------------------------------
   *
   * Primeiro sincroniza os dados.
   * Depois aplica a visão do mês selecionado.
   */

  tentarSincronizarReembolsos_();


  SpreadsheetApp.flush();

}


/**
 * ============================================================
 * 🛒 EDIÇÃO EM DESPESAS VARIÁVEIS
 * ============================================================
 *
 * Estrutura:
 *
 * A Data
 * B Categoria
 * C Descrição
 * D Valor
 * E Pagamento
 * F Cartão
 * G Responsável
 *
 * Dados começam na linha 7.
 */

function processarEdicaoDespesasVariaveis_(e) {

  const range =
    e.range;


  /**
   * Cabeçalho / área superior.
   */

  if (
    range.getLastRow() < 7
  ) {
    return;
  }


  /**
   * Só nos interessam A:G.
   */

  if (
    !intervaloCruzaColunas_(
      range,
      1,
      7
    )
  ) {
    return;
  }


  const aba =
    range.getSheet();


  /**
   * ==========================================================
   * RESPONSÁVEL FOI ALTERADO?
   * ==========================================================
   *
   * G = coluna 7
   *
   * Precisamos sincronizar tanto quando:
   * - vira terceiro
   * - deixa de ser terceiro
   */

  if (
    intervaloCruzaAlgumaColuna_(
      range,
      [7]
    )
  ) {

    if (
      edicaoPodeEnvolverTerceiro_(
        e,
        aba,
        7
      )
    ) {

      tentarSincronizarReembolsos_();

    }


    return;

  }


  /**
   * ==========================================================
   * OUTROS CAMPOS
   * ==========================================================
   *
   * Só sincronizamos se a linha já pertencer
   * a um responsável terceiro.
   */

  if (
    intervaloPossuiTerceiro_(
      aba,
      range.getRow(),
      range.getNumRows(),
      7
    )
  ) {

    tentarSincronizarReembolsos_();

  }

}


/**
 * ============================================================
 * 💳 EDIÇÃO EM PARCELAMENTOS
 * ============================================================
 *
 * Dados começam na linha 4.
 *
 * Campos relevantes:
 *
 * A Categoria
 * B Produto
 * C Valor Total
 * D Qtd Parcelas
 * G Valor Parcela
 * H Cartão
 * I Responsável
 * J Observação
 * L Data 1ª Parcela
 */

function processarEdicaoParcelamentos_(e) {

  const range =
    e.range;


  if (
    range.getLastRow() < 4
  ) {
    return;
  }


  const colunasRelevantes = [
    1,
    2,
    3,
    4,
    7,
    8,
    9,
    10,
    12
  ];


  if (
    !intervaloCruzaAlgumaColuna_(
      range,
      colunasRelevantes
    )
  ) {
    return;
  }


  const aba =
    range.getSheet();


  /**
   * ==========================================================
   * RESPONSÁVEL — COLUNA I
   * ==========================================================
   */

  if (
    intervaloCruzaAlgumaColuna_(
      range,
      [9]
    )
  ) {

    if (
      edicaoPodeEnvolverTerceiro_(
        e,
        aba,
        9
      )
    ) {

      tentarSincronizarReembolsos_();

    }


    return;

  }


  /**
   * ==========================================================
   * OUTROS CAMPOS
   * ==========================================================
   */

  if (
    intervaloPossuiTerceiro_(
      aba,
      range.getRow(),
      range.getNumRows(),
      9
    )
  ) {

    tentarSincronizarReembolsos_();

  }

}


/**
 * ============================================================
 * 👤 EDIÇÃO PODE ENVOLVER TERCEIRO?
 * ============================================================
 *
 * Verifica:
 * - valor anterior
 * - valor atual
 * - linhas atuais afetadas
 */

function edicaoPodeEnvolverTerceiro_(
  e,
  aba,
  colunaResponsavel
) {

  /**
   * Valor antigo só existe de forma confiável
   * para edição de uma única célula.
   */

  if (
    e.oldValue &&
    responsavelEhTerceiroSeguro_(
      e.oldValue
    )
  ) {

    return true;

  }


  /**
   * Valor novo.
   */

  if (
    e.value &&
    responsavelEhTerceiroSeguro_(
      e.value
    )
  ) {

    return true;

  }


  /**
   * Verifica o estado atual das linhas afetadas.
   */

  return intervaloPossuiTerceiro_(
    aba,
    e.range.getRow(),
    e.range.getNumRows(),
    colunaResponsavel
  );

}


/**
 * ============================================================
 * 👥 INTERVALO POSSUI TERCEIRO?
 * ============================================================
 */

function intervaloPossuiTerceiro_(
  aba,
  linhaInicial,
  quantidadeLinhas,
  colunaResponsavel
) {

  if (
    quantidadeLinhas <= 0
  ) {
    return false;
  }


  const valores =
    aba
      .getRange(
        linhaInicial,
        colunaResponsavel,
        quantidadeLinhas,
        1
      )
      .getDisplayValues();


  for (
    let i = 0;
    i < valores.length;
    i++
  ) {

    const responsavel =
      valores[i][0];


    if (
      responsavelEhTerceiroSeguro_(
        responsavel
      )
    ) {

      return true;

    }

  }


  return false;

}


/**
 * ============================================================
 * 👤 RESPONSÁVEL É TERCEIRO?
 * ============================================================
 *
 * Usa primeiro Configuracoes.gs.
 *
 * Caso ocorra algum problema,
 * consulta diretamente _CONFIG.
 */

function responsavelEhTerceiroSeguro_(
  responsavel
) {

  const nome =
    String(
      responsavel || ''
    ).trim();


  if (!nome) {
    return false;
  }


  /**
   * ==========================================================
   * CONFIGURAÇÕES.GS
   * ==========================================================
   */

  try {

    if (
      typeof responsavelEhTerceiro_ ===
      'function'
    ) {

      return Boolean(
        responsavelEhTerceiro_(
          nome
        )
      );

    }

  } catch (erro) {

    console.log(
      'Fallback responsável terceiro: ' +
      erro
    );

  }


  /**
   * ==========================================================
   * FALLBACK _CONFIG
   * ==========================================================
   */

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const config =
    ss.getSheetByName(
      '_CONFIG'
    );


  if (
    !config ||
    config.getLastRow() < 2
  ) {

    return false;

  }


  const dados =
    config
      .getRange(
        2,
        1,
        config.getLastRow() - 1,
        2
      )
      .getDisplayValues();


  const nomeNormalizado =
    nome
      .trim()
      .toLowerCase();


  for (
    let i = 0;
    i < dados.length;
    i++
  ) {

    const responsavelConfig =
      String(
        dados[i][0] || ''
      )
        .trim()
        .toLowerCase();


    const tipo =
      String(
        dados[i][1] || ''
      ).trim();


    if (
      responsavelConfig ===
        nomeNormalizado &&
      tipo ===
        'Terceiro / Reembolso'
    ) {

      return true;

    }

  }


  return false;

}


/**
 * ============================================================
 * 💸 SINCRONIZA REEMBOLSOS COM SEGURANÇA
 * ============================================================
 *
 * AQUI está o novo comportamento:
 *
 * 1. sincronizarReembolsos()
 * 2. aplicarVisaoCompetenciaAtualReembolsos()
 */

function tentarSincronizarReembolsos_() {

  try {

    /**
     * ========================================================
     * SINCRONIZA DADOS
     * ========================================================
     */

    if (
      typeof sincronizarReembolsos ===
      'function'
    ) {

      sincronizarReembolsos();

    }


    SpreadsheetApp.flush();


    /**
     * ========================================================
     * APLICA VISÃO DO MÊS
     * ========================================================
     */

    if (
      typeof aplicarVisaoCompetenciaAtualReembolsos ===
      'function'
    ) {

      aplicarVisaoCompetenciaAtualReembolsos();

    }

  } catch (erro) {

    console.log(
      'Erro ao sincronizar Reembolsos: ' +
      erro
    );

  }

}


/**
 * ============================================================
 * 🔎 INTERVALO CRUZA FAIXA DE COLUNAS?
 * ============================================================
 */

function intervaloCruzaColunas_(
  range,
  colunaInicial,
  colunaFinal
) {

  const inicio =
    range.getColumn();


  const fim =
    range.getLastColumn();


  return (
    inicio <= colunaFinal &&
    fim >= colunaInicial
  );

}


/**
 * ============================================================
 * 🔎 INTERVALO CRUZA ALGUMA COLUNA?
 * ============================================================
 */

function intervaloCruzaAlgumaColuna_(
  range,
  colunas
) {

  const inicio =
    range.getColumn();


  const fim =
    range.getLastColumn();


  return colunas.some(
    coluna =>
      coluna >= inicio &&
      coluna <= fim
  );

}