/**
 * ============================================================
 * 💳 CARTÕES VISUAL V2 — JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * Execute somente:
 * montarCartoesVisual()
 *
 * V2:
 * - Percentuais com mais peso visual.
 * - Faixa < 50% fica verde também no percentual.
 * - Cartões sem limite ficam mais discretos.
 * - Mantém fórmulas e estrutura existentes.
 */

function montarCartoesVisual() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('💳 Cartões');

  if (!aba) {
    throw new Error(
      'A aba 💳 Cartões não foi encontrada.'
    );
  }


  /**
   * ==========================================================
   * 🎨 PALETA JULIUS
   * ==========================================================
   */

  const C = {

    fundo: '#F1E4CF',
    fundoClaro: '#F8F0E3',
    card: '#FFF8EC',

    azulPetroleo: '#286173',
    azulEscuro: '#194754',
    azulClaro: '#D8E8EB',

    caramelo: '#C78342',
    carameloClaro: '#E7B678',

    marrom: '#654229',
    marromEscuro: '#3C2A1E',

    texto: '#33291F',
    textoSecundario: '#786554',
    branco: '#FFF9F0',

    verde: '#4F7A3B',
    verdeClaro: '#DDE8D4',

    amarelo: '#A77E18',
    amareloClaro: '#F3E7BD',

    laranja: '#C87326',
    laranjaClaro: '#FAE2C3',

    vermelho: '#B54332',
    vermelhoClaro: '#F2D5CC',

    cinza: '#8B8178',
    cinzaClaro: '#E8E1D8',

    borda: '#C3A989'
  };


  /**
   * ==========================================================
   * 🧱 BASE
   * ==========================================================
   */

  aba.setHiddenGridlines(true);
  aba.setFrozenRows(3);

  aba
    .getRange('A1:H25')
    .setFontFamily('Arial')
    .setFontColor(C.texto)
    .setVerticalAlignment('middle');


  /**
   * ==========================================================
   * 📐 LARGURAS
   * ==========================================================
   */

  aba.setColumnWidth(1, 160);
  aba.setColumnWidth(2, 135);
  aba.setColumnWidth(3, 135);
  aba.setColumnWidth(4, 135);
  aba.setColumnWidth(5, 115);
  aba.setColumnWidth(6, 100);
  aba.setColumnWidth(7, 100);
  aba.setColumnWidth(8, 145);


  /**
   * ==========================================================
   * 📏 ALTURAS
   * ==========================================================
   */

  aba.setRowHeight(1, 48);
  aba.setRowHeight(2, 52);
  aba.setRowHeight(3, 40);

  for (let linha = 4; linha <= 13; linha++) {
    aba.setRowHeight(linha, 36);
  }

  aba.setRowHeight(14, 12);
  aba.setRowHeight(15, 30);

  aba.setRowHeight(16, 36);
  aba.setRowHeight(17, 36);
  aba.setRowHeight(18, 36);

  aba.setRowHeight(19, 12);
  aba.setRowHeight(20, 34);


  /**
   * ==========================================================
   * 🔝 TÍTULO
   * ==========================================================
   */

  aba
    .getRange('A1:H1')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco);

  aba
    .getRange('A1')
    .setValue('💳 CONTROLE DE CARTÕES')
    .setFontSize(21)
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 📊 CARDS DO TOPO
   * ==========================================================
   */

  aba
    .getRange('A2:H2')
    .breakApart()
    .clearContent()
    .setBackground(C.fundo);


  /**
   * LIMITE TOTAL
   */

  cartV2Card_(
    aba,
    'A2:B2',
    C.card,
    C.borda
  );

  aba
    .getRange('A2:B2')
    .merge();

  aba
    .getRange('A2')
    .setFormula(
      '="💰 LIMITE TOTAL  •  "&TEXT(SUM(B4:B13);"R$ #,##0.00")'
    )
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(C.azulPetroleo)
    .setHorizontalAlignment('center');


  /**
   * UTILIZADO
   */

  cartV2Card_(
    aba,
    'C2:D2',
    C.card,
    C.borda
  );

  aba
    .getRange('C2:D2')
    .merge();

  aba
    .getRange('C2')
    .setFormula(
      '="💸 UTILIZADO  •  "&TEXT(SUM(C4:C13);"R$ #,##0.00")'
    )
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(C.vermelho)
    .setHorizontalAlignment('center');


  /**
   * DISPONÍVEL
   */

  cartV2Card_(
    aba,
    'E2:F2',
    C.card,
    C.borda
  );

  aba
    .getRange('E2:F2')
    .merge();

  aba
    .getRange('E2')
    .setFormula(
      '="💵 DISPONÍVEL  •  "&TEXT(SUM(D4:D13);"R$ #,##0.00")'
    )
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(C.verde)
    .setHorizontalAlignment('center');


  /**
   * USO GERAL
   */

  cartV2Card_(
    aba,
    'G2:H2',
    C.card,
    C.borda
  );

  aba
    .getRange('G2:H2')
    .merge();

  aba
    .getRange('G2')
    .setFormula(
      '="📊 USO GERAL  •  "&TEXT(IFERROR(SUM(C4:C13)/SUM(B4:B13);0);"0.0%")'
    )
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(C.caramelo)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 📋 CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A3:H3')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco)
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);


  /**
   * ==========================================================
   * 💳 CORPO
   * ==========================================================
   */

  for (let linha = 4; linha <= 13; linha++) {

    const nome =
      String(
        aba
          .getRange(linha, 1)
          .getDisplayValue()
      ).trim();


    if (!nome) {
      continue;
    }


    const limite =
      Number(
        aba
          .getRange(linha, 2)
          .getValue()
      ) || 0;


    const percentual =
      Number(
        aba
          .getRange(linha, 5)
          .getValue()
      ) || 0;


    let fundo =
      linha % 2 === 0
        ? C.card
        : C.fundoClaro;


    /**
     * Sem limite configurado
     */

    if (limite <= 0) {

      fundo = C.cinzaClaro;

    } else if (percentual >= 0.90) {

      fundo = C.vermelhoClaro;

    } else if (percentual >= 0.75) {

      fundo = C.laranjaClaro;

    } else if (percentual >= 0.50) {

      fundo = C.amareloClaro;

    } else {

      fundo = C.verdeClaro;
    }


    const faixa =
      aba.getRange(
        linha,
        1,
        1,
        8
      );


    faixa
      .setBackground(fundo)
      .setFontSize(10)
      .setFontColor(
        limite <= 0
          ? C.cinza
          : C.texto
      );


    cartV2Borda_(
      faixa,
      C.borda
    );


    /**
     * Nome do cartão
     */

    aba
      .getRange(linha, 1)
      .setFontWeight('bold')
      .setFontColor(
        limite <= 0
          ? C.cinza
          : C.azulEscuro
      );


    /**
     * ========================================================
     * 📊 PERCENTUAL — MAIS FORTE
     * ========================================================
     */

    aba
      .getRange(linha, 5)
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');


    if (limite <= 0) {

      aba
        .getRange(linha, 5)
        .setFontColor(C.cinza);

    } else if (percentual >= 0.90) {

      aba
        .getRange(linha, 5)
        .setBackground('#F05B4F')
        .setFontColor('#7A160D');

    } else if (percentual >= 0.75) {

      aba
        .getRange(linha, 5)
        .setBackground('#ED9C4B')
        .setFontColor('#743D08');

    } else if (percentual >= 0.50) {

      aba
        .getRange(linha, 5)
        .setBackground('#E2C95B')
        .setFontColor('#604F00');

    } else {

      aba
        .getRange(linha, 5)
        .setBackground('#7DBB6A')
        .setFontColor('#204A18');
    }

  }


  /**
   * ==========================================================
   * 💵 FORMATOS
   * ==========================================================
   */

  aba
    .getRange('B4:D13')
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  aba
    .getRange('E4:E13')
    .setNumberFormat('0.0%');


  /**
   * ==========================================================
   * 🎯 ALINHAMENTO
   * ==========================================================
   */

  aba
    .getRange('A4:A13')
    .setHorizontalAlignment('left');


  aba
    .getRange('B4:H13')
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * ✏️ CAMPOS EDITÁVEIS
   * ==========================================================
   */

  aba
    .getRange('B4:B13')
    .setFontWeight('bold');


  aba
    .getRange('B3')
    .setNote(
      'Informe o limite total do cartão. Se ficar vazio, o status será ⚪ Configurar.'
    );


  aba
    .getRange('F3')
    .setNote(
      'Dia de fechamento da fatura.'
    );


  aba
    .getRange('G3')
    .setNote(
      'Dia de vencimento da fatura.'
    );


  /**
   * ==========================================================
   * 🚦 STATUS
   * ==========================================================
   */

  cartV2FormatacaoCondicional_(
    aba,
    C
  );


  /**
   * ==========================================================
   * 👀 RESUMO INFERIOR
   * ==========================================================
   */

  aba
    .getRange('A15:H15')
    .setBackground(C.marrom)
    .setFontColor(C.branco);


  aba
    .getRange('A15')
    .setValue('👀 JULIUS CONFERE O CRÉDITO')
    .setFontWeight('bold')
    .setFontSize(11);


  for (let linha = 16; linha <= 18; linha++) {

    const faixa =
      aba.getRange(
        linha,
        1,
        1,
        8
      );


    faixa
      .setBackground(C.card)
      .setFontSize(10);


    cartV2Borda_(
      faixa,
      C.borda
    );


    aba
      .getRange(linha, 1)
      .setFontWeight('bold')
      .setFontColor(C.marromEscuro);


    aba
      .getRange(linha, 3)
      .setFontWeight('bold')
      .setFontSize(12)
      .setNumberFormat(
        '"R$ " #,##0.00;-"R$ " #,##0.00'
      );

  }


  aba
    .getRange('C16')
    .setFontColor(C.azulPetroleo);

  aba
    .getRange('C17')
    .setFontColor(C.vermelho);

  aba
    .getRange('C18')
    .setFontColor(C.verde);


  /**
   * ==========================================================
   * 🧾 LEGENDA
   * ==========================================================
   */

  aba
    .getRange('A20:H20')
    .setBackground(C.fundoClaro)
    .setFontColor(C.textoSecundario);


  aba
    .getRange('A20')
    .setValue(
      '🟢 abaixo de 50%   •   🟡 50–75%   •   🟠 75–90%   •   🔴 acima de 90%   •   ⚪ limite não configurado'
    )
    .setFontSize(9)
    .setFontStyle('italic');


  /**
   * ==========================================================
   * ✅ FINAL
   * ==========================================================
   */

  SpreadsheetApp.flush();


  ss.toast(
    'Cartões Julius V2 aplicados 💳👀',
    'Controle Financeiro',
    4
  );

}


/**
 * ============================================================
 * 🧱 CARD
 * ============================================================
 */

function cartV2Card_(
  aba,
  intervalo,
  fundo,
  borda
) {

  const range =
    aba.getRange(intervalo);


  range
    .setBackground(fundo)
    .setVerticalAlignment('middle');


  cartV2Borda_(
    range,
    borda
  );

}


/**
 * ============================================================
 * 🚦 STATUS
 * ============================================================
 */

function cartV2FormatacaoCondicional_(
  aba,
  C
) {

  const regrasAtuais =
    aba.getConditionalFormatRules();


  const preservadas =
    regrasAtuais.filter(regra => {

      return !regra
        .getRanges()
        .some(range => {

          return (
            range.getColumn() === 8 &&
            range.getRow() >= 4 &&
            range.getRow() <= 13
          );

        });

    });


  const novas = [];


  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('🔴')
      .setBackground(C.vermelhoClaro)
      .setFontColor(C.vermelho)
      .setBold(true)
      .setRanges([
        aba.getRange('H4:H13')
      ])
      .build()

  );


  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('🟠')
      .setBackground(C.laranjaClaro)
      .setFontColor(C.laranja)
      .setBold(true)
      .setRanges([
        aba.getRange('H4:H13')
      ])
      .build()

  );


  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('🟡')
      .setBackground(C.amareloClaro)
      .setFontColor(C.marromEscuro)
      .setBold(true)
      .setRanges([
        aba.getRange('H4:H13')
      ])
      .build()

  );


  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('🟢')
      .setBackground(C.verdeClaro)
      .setFontColor(C.verde)
      .setBold(true)
      .setRanges([
        aba.getRange('H4:H13')
      ])
      .build()

  );


  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('⚪')
      .setBackground(C.cinzaClaro)
      .setFontColor(C.cinza)
      .setBold(true)
      .setRanges([
        aba.getRange('H4:H13')
      ])
      .build()

  );


  aba.setConditionalFormatRules(
    preservadas.concat(novas)
  );

}


/**
 * ============================================================
 * 🔲 BORDA
 * ============================================================
 */

function cartV2Borda_(
  range,
  cor
) {

  range.setBorder(
    true,
    true,
    true,
    true,
    false,
    false,
    cor,
    SpreadsheetApp.BorderStyle.SOLID
  );

}