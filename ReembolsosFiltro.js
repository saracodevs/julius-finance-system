/**
 * ============================================================
 * 👁️ FILTRO DE COMPETÊNCIA — REEMBOLSOS
 * JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * A aba 💸 Reembolsos continua armazenando todas as dívidas.
 *
 * Esta rotina apenas OCULTA as linhas que não pertencem
 * à competência selecionada em 📊 Dashboard!B7.
 *
 * Nada é apagado.
 *
 * Execute manualmente:
 * aplicarVisaoCompetenciaAtualReembolsos()
 *
 * Depois vamos integrar ao sistema automaticamente.
 */


const REEMB_FILTRO_CONFIG = {

  ABA:
    '💸 Reembolsos',

  DASHBOARD:
    '📊 Dashboard',

  CELULA_COMPETENCIA:
    'B7',

  PRIMEIRA_LINHA:
    5,

  COLUNA_COMPETENCIA:
    10 // J

};


/**
 * ============================================================
 * 👁️ VISÃO DA COMPETÊNCIA ATUAL
 * ============================================================
 */

function aplicarVisaoCompetenciaAtualReembolsos() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const aba =
    ss.getSheetByName(
      REEMB_FILTRO_CONFIG.ABA
    );


  const dashboard =
    ss.getSheetByName(
      REEMB_FILTRO_CONFIG.DASHBOARD
    );


  if (
    !aba ||
    !dashboard
  ) {

    throw new Error(
      'Não foi possível encontrar Dashboard ou Reembolsos.'
    );

  }


  const competencia =
    dashboard
      .getRange(
        REEMB_FILTRO_CONFIG.CELULA_COMPETENCIA
      )
      .getValue();


  if (
    !(competencia instanceof Date) ||
    isNaN(
      competencia.getTime()
    )
  ) {

    throw new Error(
      '📊 Dashboard!B7 não contém uma competência válida.'
    );

  }


  const ultimaLinha =
    aba.getLastRow();


  /**
   * Primeiro mostra todas as linhas.
   *
   * Isso é importante quando mudamos:
   * agosto -> julho -> agosto.
   */

  if (
    ultimaLinha >=
    REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA
  ) {

    aba.showRows(
      REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA,
      ultimaLinha -
        REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA +
        1
    );

  }


  if (
    ultimaLinha <
    REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA
  ) {

    return;

  }


  /**
   * Lê somente a coluna técnica J.
   */

  const competencias =
    aba
      .getRange(
        REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA,
        REEMB_FILTRO_CONFIG.COLUNA_COMPETENCIA,
        ultimaLinha -
          REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA +
          1,
        1
      )
      .getValues();


  const anoSelecionado =
    competencia.getFullYear();


  const mesSelecionado =
    competencia.getMonth();


  /**
   * ==========================================================
   * ENCONTRA BLOCOS QUE DEVEM SER OCULTADOS
   * ==========================================================
   *
   * Em vez de chamar hideRows linha por linha,
   * agrupamos sequências para ficar rápido.
   */

  let inicioBloco =
    null;


  let tamanhoBloco =
    0;


  for (
    let i = 0;
    i < competencias.length;
    i++
  ) {

    const valor =
      competencias[i][0];


    let pertenceAoMes =
      false;


    if (
      valor instanceof Date &&
      !isNaN(
        valor.getTime()
      )
    ) {

      pertenceAoMes =
        (
          valor.getFullYear() ===
          anoSelecionado
        ) &&
        (
          valor.getMonth() ===
          mesSelecionado
        );

    }


    /**
     * Linha física na planilha.
     */

    const linhaPlanilha =
      REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA +
      i;


    if (
      !pertenceAoMes
    ) {

      if (
        inicioBloco === null
      ) {

        inicioBloco =
          linhaPlanilha;

        tamanhoBloco =
          1;

      } else {

        tamanhoBloco++;

      }

    } else {

      /**
       * Encontramos uma linha visível.
       * Fecha o bloco anterior.
       */

      if (
        inicioBloco !== null
      ) {

        aba.hideRows(
          inicioBloco,
          tamanhoBloco
        );


        inicioBloco =
          null;

        tamanhoBloco =
          0;

      }

    }

  }


  /**
   * Pode ter sobrado bloco no final.
   */

  if (
    inicioBloco !== null
  ) {

    aba.hideRows(
      inicioBloco,
      tamanhoBloco
    );

  }


  SpreadsheetApp.flush();


  ss.toast(
    'Reembolsos exibindo ' +
      formatarMesAnoReembolsosFiltro_(
        competencia
      ) +
      ' 👁️',
    'Julius Finance',
    3
  );

}


/**
 * ============================================================
 * 👁️ MOSTRAR TODAS AS DÍVIDAS
 * ============================================================
 *
 * Vamos usar esta função mais tarde no seletor:
 *
 * Competência atual
 * Todos os pendentes
 * Histórico completo
 */

function mostrarTodosReembolsos() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const aba =
    ss.getSheetByName(
      REEMB_FILTRO_CONFIG.ABA
    );


  if (!aba) {
    return;
  }


  const ultimaLinha =
    aba.getLastRow();


  if (
    ultimaLinha >=
    REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA
  ) {

    aba.showRows(
      REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA,
      ultimaLinha -
        REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA +
        1
    );

  }


  SpreadsheetApp.flush();

}


/**
 * ============================================================
 * 📅 TEXTO DA COMPETÊNCIA
 * ============================================================
 */

function formatarMesAnoReembolsosFiltro_(
  data
) {

  const meses = [
    'janeiro',
    'fevereiro',
    'março',
    'abril',
    'maio',
    'junho',
    'julho',
    'agosto',
    'setembro',
    'outubro',
    'novembro',
    'dezembro'
  ];


  return (
    meses[
      data.getMonth()
    ] +
    '/' +
    data.getFullYear()
  );

}