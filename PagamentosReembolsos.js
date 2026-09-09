/**
 * ============================================================
 * 💵 PAGAMENTOS DE REEMBOLSOS
 * JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * PASSO 1
 *
 * Cria:
 *
 * 1. 💵 Pagamentos Reembolsos
 *    Histórico de cada valor realmente recebido.
 *
 * 2. _REEMB_ALOCACOES
 *    Controle técnico de como cada pagamento foi distribuído
 *    entre as dívidas.
 *
 * Execute somente:
 * prepararSistemaPagamentosReembolsos()
 */


const PAG_REEMB_CONFIG = {

  ABA_PAGAMENTOS:
    '💵 Pagamentos Reembolsos',

  ABA_ALOCACOES:
    '_REEMB_ALOCACOES',

  TIMEZONE:
    'America/Sao_Paulo'
};


/**
 * ============================================================
 * 🚀 FUNÇÃO PRINCIPAL
 * ============================================================
 */

function prepararSistemaPagamentosReembolsos() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  /**
   * Mantém timezone correto.
   */

  try {

    if (
      ss.getSpreadsheetTimeZone() !==
      PAG_REEMB_CONFIG.TIMEZONE
    ) {

      ss.setSpreadsheetTimeZone(
        PAG_REEMB_CONFIG.TIMEZONE
      );

    }

  } catch (erro) {

    console.log(
      'Timezone não alterado: ' +
      erro
    );

  }


  prepararAbaHistoricoPagamentos_(ss);

  prepararAbaAlocacoesReembolsos_(ss);


  SpreadsheetApp.flush();


  ss.toast(
    'Sistema de pagamentos preparado 💵',
    'Julius Finance',
    4
  );

}


/**
 * ============================================================
 * 💵 HISTÓRICO DE PAGAMENTOS
 * ============================================================
 *
 * Uma linha = um pagamento recebido.
 *
 * Exemplo:
 *
 * 09/09/2026 | Renata | 200,00 | PIX | Pagamento parcial
 *
 * O valor NÃO é ligado manualmente a uma compra.
 * A distribuição será feita automaticamente depois.
 */

function prepararAbaHistoricoPagamentos_(ss) {

  let aba =
    ss.getSheetByName(
      PAG_REEMB_CONFIG.ABA_PAGAMENTOS
    );


  if (!aba) {

    aba =
      ss.insertSheet(
        PAG_REEMB_CONFIG.ABA_PAGAMENTOS
      );

  }


  /**
   * ==========================================================
   * CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A1:G1')
    .setValues([[
      'ID Pagamento',
      'Data Recebimento',
      'Responsável',
      'Valor Recebido',
      'Forma de Pagamento',
      'Observação',
      'Registrado em'
    ]]);


  /**
   * ==========================================================
   * VISUAL
   * ==========================================================
   */

  const C = {

    azul:
      '#286173',

    branco:
      '#FFF9F0',

    fundo:
      '#FFF8EC',

    borda:
      '#C3A989',

    texto:
      '#33291F'
  };


  aba.setHiddenGridlines(true);

  aba.setFrozenRows(1);


  aba
    .getRange('A1:G1')
    .setBackground(C.azul)
    .setFontColor(C.branco)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');


  aba
    .getRange(
      2,
      1,
      Math.max(
        aba.getMaxRows() - 1,
        1
      ),
      7
    )
    .setBackground(C.fundo)
    .setFontColor(C.texto);


  aba.setColumnWidth(1, 170);
  aba.setColumnWidth(2, 130);
  aba.setColumnWidth(3, 160);
  aba.setColumnWidth(4, 140);
  aba.setColumnWidth(5, 160);
  aba.setColumnWidth(6, 260);
  aba.setColumnWidth(7, 170);


  aba
    .getRange(
      'B2:B' +
      aba.getMaxRows()
    )
    .setNumberFormat(
      'dd/mm/yyyy'
    );


  aba
    .getRange(
      'D2:D' +
      aba.getMaxRows()
    )
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  aba
    .getRange(
      'G2:G' +
      aba.getMaxRows()
    )
    .setNumberFormat(
      'dd/mm/yyyy hh:mm'
    );


  /**
   * Bordas somente no cabeçalho.
   */

  aba
    .getRange('A1:G1')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      C.borda,
      SpreadsheetApp.BorderStyle.SOLID
    );

}


/**
 * ============================================================
 * ⚙️ ALOCAÇÕES TÉCNICAS
 * ============================================================
 *
 * Essa aba registra:
 *
 * "Dos R$ 200 recebidos, quanto foi usado em cada dívida?"
 *
 * Exemplo:
 *
 * PAG-001 | CHAVE-DIVIDA-1 | 114,99
 * PAG-001 | CHAVE-DIVIDA-2 | 34,62
 * PAG-001 | CHAVE-DIVIDA-3 | 50,39
 *
 * Assim conseguimos reconstruir tudo depois.
 */

function prepararAbaAlocacoesReembolsos_(ss) {

  let aba =
    ss.getSheetByName(
      PAG_REEMB_CONFIG.ABA_ALOCACOES
    );


  if (!aba) {

    aba =
      ss.insertSheet(
        PAG_REEMB_CONFIG.ABA_ALOCACOES
      );

  }


  aba
    .getRange('A1:F1')
    .setValues([[
      'ID Pagamento',
      'Chave Reembolso',
      'Valor Alocado',
      'Data Pagamento',
      'Responsável',
      'Competência Dívida'
    ]]);


  aba
    .getRange('A1:F1')
    .setFontWeight('bold');


  aba
    .getRange(
      'C2:C' +
      aba.getMaxRows()
    )
    .setNumberFormat(
      '"R$ " #,##0.00'
    );


  aba
    .getRange(
      'D2:D' +
      aba.getMaxRows()
    )
    .setNumberFormat(
      'dd/mm/yyyy'
    );


  aba
    .getRange(
      'F2:F' +
      aba.getMaxRows()
    )
    .setNumberFormat(
      'mm/yyyy'
    );


  /**
   * Esta aba é puramente técnica.
   */

  if (!aba.isSheetHidden()) {

    aba.hideSheet();

  }

}