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
 */

const REEMB_FILTRO_CONFIG = {
  ABA: '💸 Reembolsos',
  DASHBOARD: '📊 Dashboard',
  CELULA_COMPETENCIA: 'B7',
  PRIMEIRA_LINHA: 5,
  COLUNA_COMPETENCIA: 10 // J
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

  if (!aba || !dashboard) {
    throw new Error(
      'Não foi possível encontrar Dashboard ou Reembolsos.'
    );
  }


  /**
   * ==========================================================
   * COMPETÊNCIA SELECIONADA NO DASHBOARD
   * ==========================================================
   */

  const rangeCompetencia =
    dashboard.getRange(
      REEMB_FILTRO_CONFIG.CELULA_COMPETENCIA
    );

  const competenciaValor =
    rangeCompetencia.getValue();

  const competenciaTexto =
    rangeCompetencia.getDisplayValue();

  const competenciaSelecionada =
    extrairAnoMesReembolsoFiltro_(
      competenciaValor,
      competenciaTexto
    );

  if (!competenciaSelecionada) {
    throw new Error(
      '📊 Dashboard!B7 não contém uma competência válida.'
    );
  }


  /**
   * ==========================================================
   * LINHAS DA ABA
   * ==========================================================
   */

  const ultimaLinha =
    aba.getLastRow();

  if (
    ultimaLinha <
    REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA
  ) {
    return;
  }


  const quantidadeLinhas =
    ultimaLinha -
    REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA +
    1;


  /**
   * Primeiro mostra todas.
   *
   * Isso garante funcionamento correto ao navegar:
   * agosto -> setembro -> agosto.
   */

  aba.showRows(
    REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA,
    quantidadeLinhas
  );


  /**
   * ==========================================================
   * LÊ A COMPETÊNCIA TÉCNICA — COLUNA J
   * ==========================================================
   *
   * Pegamos tanto o valor real quanto o texto exibido.
   *
   * Assim funciona mesmo se J estiver:
   *
   * Date
   * 08/2026
   * 01/08/2026
   * agosto/2026
   */

  const rangeCompetencias =
    aba.getRange(
      REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA,
      REEMB_FILTRO_CONFIG.COLUNA_COMPETENCIA,
      quantidadeLinhas,
      1
    );

  const valores =
    rangeCompetencias.getValues();

  const textos =
    rangeCompetencias.getDisplayValues();


  /**
   * ==========================================================
   * IDENTIFICA BLOCOS A OCULTAR
   * ==========================================================
   */

  let inicioBloco = null;
  let tamanhoBloco = 0;

  for (
    let i = 0;
    i < quantidadeLinhas;
    i++
  ) {

    const competenciaLinha =
      extrairAnoMesReembolsoFiltro_(
        valores[i][0],
        textos[i][0]
      );

    const pertenceAoMes =
      competenciaLinha !== null &&
      competenciaLinha.ano ===
        competenciaSelecionada.ano &&
      competenciaLinha.mes ===
        competenciaSelecionada.mes;


    const linhaPlanilha =
      REEMB_FILTRO_CONFIG.PRIMEIRA_LINHA +
      i;


    if (!pertenceAoMes) {

      if (inicioBloco === null) {

        inicioBloco =
          linhaPlanilha;

        tamanhoBloco = 1;

      } else {

        tamanhoBloco++;

      }

    } else {

      if (inicioBloco !== null) {

        aba.hideRows(
          inicioBloco,
          tamanhoBloco
        );

        inicioBloco = null;
        tamanhoBloco = 0;

      }

    }

  }


  /**
   * Fecha eventual bloco final.
   */

  if (inicioBloco !== null) {

    aba.hideRows(
      inicioBloco,
      tamanhoBloco
    );

  }


  SpreadsheetApp.flush();


  ss.toast(
    'Reembolsos exibindo ' +
      formatarMesAnoReembolsosFiltroPorPartes_(
        competenciaSelecionada.ano,
        competenciaSelecionada.mes
      ) +
      ' 👁️',
    'Julius Finance',
    3
  );

}


/**
 * ============================================================
 * 🧠 INTERPRETA QUALQUER FORMATO DE COMPETÊNCIA
 * ============================================================
 *
 * Retorna:
 *
 * {
 *   ano: 2026,
 *   mes: 8
 * }
 *
 * O mês aqui vai de 1 a 12.
 */

function extrairAnoMesReembolsoFiltro_(
  valor,
  texto
) {

  /**
   * ----------------------------------------------------------
   * 1. DATA REAL
   * ----------------------------------------------------------
   */

  if (
    valor instanceof Date &&
    !isNaN(valor.getTime())
  ) {

    return {
      ano: valor.getFullYear(),
      mes: valor.getMonth() + 1
    };

  }


  /**
   * ----------------------------------------------------------
   * 2. TEXTO
   * ----------------------------------------------------------
   */

  const bruto =
    String(
      texto ||
      valor ||
      ''
    )
      .trim()
      .toLowerCase();


  if (!bruto) {
    return null;
  }


  /**
   * ----------------------------------------------------------
   * mm/yyyy
   * ----------------------------------------------------------
   */

  let match =
    bruto.match(
      /^(\d{1,2})\/(\d{4})$/
    );

  if (match) {

    const mes =
      Number(match[1]);

    const ano =
      Number(match[2]);

    if (
      mes >= 1 &&
      mes <= 12
    ) {

      return {
        ano: ano,
        mes: mes
      };

    }

  }


  /**
   * ----------------------------------------------------------
   * dd/mm/yyyy
   * ----------------------------------------------------------
   */

  match =
    bruto.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
    );

  if (match) {

    const mes =
      Number(match[2]);

    const ano =
      Number(match[3]);

    if (
      mes >= 1 &&
      mes <= 12
    ) {

      return {
        ano: ano,
        mes: mes
      };

    }

  }


  /**
   * ----------------------------------------------------------
   * yyyy-mm-dd
   * ----------------------------------------------------------
   */

  match =
    bruto.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})/
    );

  if (match) {

    const ano =
      Number(match[1]);

    const mes =
      Number(match[2]);

    if (
      mes >= 1 &&
      mes <= 12
    ) {

      return {
        ano: ano,
        mes: mes
      };

    }

  }


  /**
   * ----------------------------------------------------------
   * nome-do-mês/ano
   * ----------------------------------------------------------
   */

  const meses = {
    janeiro: 1,
    fevereiro: 2,
    marco: 3,
    março: 3,
    abril: 4,
    maio: 5,
    junho: 6,
    julho: 7,
    agosto: 8,
    setembro: 9,
    outubro: 10,
    novembro: 11,
    dezembro: 12
  };


  match =
    bruto.match(
      /^([a-záéíóúâêôãõç]+)\s*\/\s*(\d{4})$/i
    );

  if (match) {

    const nomeMes =
      match[1];

    const ano =
      Number(match[2]);

    const mes =
      meses[nomeMes];

    if (mes) {

      return {
        ano: ano,
        mes: mes
      };

    }

  }


  return null;

}


/**
 * ============================================================
 * 👁️ MOSTRAR TODAS AS DÍVIDAS
 * ============================================================
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

  if (
    !(data instanceof Date) ||
    isNaN(data.getTime())
  ) {
    return '';
  }


  return formatarMesAnoReembolsosFiltroPorPartes_(
    data.getFullYear(),
    data.getMonth() + 1
  );

}


/**
 * ============================================================
 * 📅 TEXTO DA COMPETÊNCIA POR ANO/MÊS
 * ============================================================
 */

function formatarMesAnoReembolsosFiltroPorPartes_(
  ano,
  mes
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
    meses[mes - 1] +
    '/' +
    ano
  );

}