/**
 * ============================================================
 * 💸 REEMBOLSOS VISUAL V2 — JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * Execute somente:
 * montarReembolsosVisual()
 *
 * - Não altera Reembolsos.gs
 * - Não altera dados A:I
 * - Não altera fórmulas G/H
 * - J:L permanecem ocultas
 * - F continua sendo Valor Pago manual
 */

function montarReembolsosVisual() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('💸 Reembolsos');

  if (!aba) {
    throw new Error(
      'A aba 💸 Reembolsos não foi encontrada.'
    );
  }


  const C = {

    fundo: '#F1E4CF',
    fundoClaro: '#F8F0E3',
    card: '#FFF8EC',

    azulPetroleo: '#286173',
    azulEscuro: '#194754',

    caramelo: '#C78342',

    marrom: '#654229',
    marromEscuro: '#3C2A1E',

    texto: '#33291F',
    textoSecundario: '#786554',
    branco: '#FFF9F0',

    verde: '#4F7A3B',
    verdeClaro: '#DDE8D4',

    amareloClaro: '#F3E7BD',

    vermelho: '#B54332',
    vermelhoClaro: '#F2D5CC',

    cinza: '#8B8178',

    borda: '#C3A989',

    editavel: '#FFF1C9'
  };


  /**
   * ==========================================================
   * 🔎 TAMANHO DA BASE
   * ==========================================================
   */

  const ultimaLinha =
    Math.max(
      aba.getLastRow(),
      5
    );

  const maxLinhas =
    aba.getMaxRows();


  /**
   * ==========================================================
   * 🧱 BASE
   * ==========================================================
   */

  aba.setHiddenGridlines(true);

  aba.setFrozenRows(4);


  aba
    .getRange(
      1,
      1,
      maxLinhas,
      12
    )
    .setFontFamily('Arial')
    .setVerticalAlignment('middle');


  /**
   * ==========================================================
   * 📐 COLUNAS
   * ==========================================================
   */

  aba.setColumnWidth(1, 110);
  aba.setColumnWidth(2, 150);
  aba.setColumnWidth(3, 135);
  aba.setColumnWidth(4, 255);

  aba.setColumnWidth(5, 125);
  aba.setColumnWidth(6, 125);
  aba.setColumnWidth(7, 125);

  aba.setColumnWidth(8, 125);
  aba.setColumnWidth(9, 180);


  /**
   * J:L = técnicas
   */

  aba.hideColumns(10, 3);


  /**
   * ==========================================================
   * 📏 LINHAS
   * ==========================================================
   */

  aba.setRowHeight(1, 48);
  aba.setRowHeight(2, 34);
  aba.setRowHeight(3, 34);
  aba.setRowHeight(4, 40);


  /**
   * ==========================================================
   * 🔝 TÍTULO
   * ==========================================================
   */

  aba
    .getRange('A1:I1')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco);


  aba
    .getRange('A1')
    .setValue('💸 CONTROLE DE REEMBOLSOS')
    .setFontSize(21)
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 🧹 PREPARAÇÃO SEGURA DOS CARDS
   * ==========================================================
   *
   * CORREÇÃO DO ERRO:
   *
   * Em vez de chamar breakApart() diretamente em A2:I3,
   * primeiro localizamos cada mesclagem existente e
   * desfazemos o intervalo completo dela.
   */

  reembVDesfazerMesclasNaArea_(
    aba,
    'A2:I3'
  );


  aba
    .getRange('A2:I3')
    .clearContent()
    .setBackground(C.fundo)
    .setFontColor(C.texto);


  /**
   * ==========================================================
   * 💰 TOTAL A RECEBER
   * ==========================================================
   */

  aba
    .getRange('A2:B3')
    .merge();


  reembVCard_(
    aba,
    'A2:B3',
    C.card,
    C.borda
  );


  aba
    .getRange('A2')
    .setFormula(
      '="💰 TOTAL A RECEBER"&CHAR(10)&' +
      'TEXT(IFERROR(SUM(E5:E' +
      ultimaLinha +
      ');0);"R$ #,##0.00")'
    )
    .setFontSize(11)
    .setFontWeight('bold')
    .setFontColor(C.azulPetroleo)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);


  /**
   * ==========================================================
   * ✅ JÁ RECEBIDO
   * ==========================================================
   */

  aba
    .getRange('C2:D3')
    .merge();


  reembVCard_(
    aba,
    'C2:D3',
    C.verdeClaro,
    C.borda
  );


  aba
    .getRange('C2')
    .setFormula(
      '="✅ JÁ RECEBIDO"&CHAR(10)&' +
      'TEXT(IFERROR(SUM(F5:F' +
      ultimaLinha +
      ');0);"R$ #,##0.00")'
    )
    .setFontSize(11)
    .setFontWeight('bold')
    .setFontColor(C.verde)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);


  /**
   * ==========================================================
   * 🔴 SALDO PENDENTE
   * ==========================================================
   */

  aba
    .getRange('E2:F3')
    .merge();


  reembVCard_(
    aba,
    'E2:F3',
    C.vermelhoClaro,
    C.borda
  );


  aba
    .getRange('E2')
    .setFormula(
      '="🔴 SALDO PENDENTE"&CHAR(10)&' +
      'TEXT(IFERROR(SUM(G5:G' +
      ultimaLinha +
      ');0);"R$ #,##0.00")'
    )
    .setFontSize(11)
    .setFontWeight('bold')
    .setFontColor(C.vermelho)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);


  /**
   * ==========================================================
   * 👀 PENDÊNCIAS
   * ==========================================================
   */

  aba
    .getRange('G2:I3')
    .merge();


  reembVCard_(
    aba,
    'G2:I3',
    C.amareloClaro,
    C.borda
  );


  aba
    .getRange('G2')
    .setFormula(
      '="👀 PENDÊNCIAS"&CHAR(10)&' +
      'COUNTIF(G5:G' +
      ultimaLinha +
      ';">0")&" lançamentos"'
    )
    .setFontSize(11)
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);


  /**
   * ==========================================================
   * 📋 CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A4:I4')
    .setValues([[
      '📅 Data',
      '👤 Responsável',
      '💳 Origem',
      '📝 Descrição',
      '💰 Valor Devido',
      '✅ Valor Pago',
      '💸 Saldo',
      '🚦 Status',
      '📌 Observação'
    ]]);


  aba
    .getRange('A4:I4')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco)
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);


  /**
   * ==========================================================
   * 📋 CORPO
   * ==========================================================
   */

  const quantidadeLinhas =
    Math.max(
      ultimaLinha - 4,
      1
    );


  for (
    let linha = 5;
    linha <= ultimaLinha;
    linha++
  ) {

    const responsavel =
      String(
        aba
          .getRange(linha, 2)
          .getDisplayValue()
      ).trim();


    const descricao =
      String(
        aba
          .getRange(linha, 4)
          .getDisplayValue()
      ).trim();


    if (
      !responsavel &&
      !descricao
    ) {
      continue;
    }


    const saldo =
      Number(
        aba
          .getRange(linha, 7)
          .getValue()
      ) || 0;


    const pago =
      Number(
        aba
          .getRange(linha, 6)
          .getValue()
      ) || 0;


    let fundo =
      linha % 2 === 0
        ? C.card
        : C.fundoClaro;


    if (saldo === 0) {

      fundo =
        C.verdeClaro;

    } else if (pago > 0) {

      fundo =
        C.amareloClaro;

    }


    const faixa =
      aba.getRange(
        linha,
        1,
        1,
        9
      );


    faixa
      .setBackground(fundo)
      .setFontColor(C.texto)
      .setFontSize(9);


    reembVBorda_(
      faixa,
      C.borda
    );


    aba.setRowHeight(
      linha,
      32
    );

  }


  /**
   * ==========================================================
   * ✏️ VALOR PAGO
   * ==========================================================
   */

  aba
    .getRange(
      5,
      6,
      quantidadeLinhas,
      1
    )
    .setBackground(C.editavel)
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro);


  /**
   * ==========================================================
   * 💵 FORMATAÇÃO
   * ==========================================================
   */

  aba
    .getRange(
      5,
      1,
      quantidadeLinhas,
      1
    )
    .setNumberFormat('dd/mm/yyyy');


  aba
    .getRange(
      5,
      5,
      quantidadeLinhas,
      3
    )
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  /**
   * ==========================================================
   * 🎯 ALINHAMENTO
   * ==========================================================
   */

  aba
    .getRange(
      5,
      1,
      quantidadeLinhas,
      1
    )
    .setHorizontalAlignment('center');


  aba
    .getRange(
      5,
      2,
      quantidadeLinhas,
      3
    )
    .setHorizontalAlignment('left');


  aba
    .getRange(
      5,
      5,
      quantidadeLinhas,
      4
    )
    .setHorizontalAlignment('center');


  aba
    .getRange(
      5,
      9,
      quantidadeLinhas,
      1
    )
    .setHorizontalAlignment('left');


  /**
   * ==========================================================
   * DESTAQUES
   * ==========================================================
   */

  aba
    .getRange(
      5,
      5,
      quantidadeLinhas,
      1
    )
    .setFontWeight('bold');


  aba
    .getRange(
      5,
      7,
      quantidadeLinhas,
      1
    )
    .setFontWeight('bold');


  aba
    .getRange(
      5,
      8,
      quantidadeLinhas,
      1
    )
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 🚦 STATUS
   * ==========================================================
   */

  reembVFormatacaoCondicional_(
    aba,
    C,
    ultimaLinha
  );


  /**
   * ==========================================================
   * 📝 NOTAS
   * ==========================================================
   */

  aba
    .getRange('F4')
    .setNote(
      'Campo manual: informe aqui quanto já foi recebido referente a este lançamento.'
    );


  aba
    .getRange('G4')
    .setNote(
      'Saldo automático: Valor Devido menos Valor Pago.'
    );


  aba
    .getRange('H4')
    .setNote(
      'Status automático: Pendente, Parcial ou Pago.'
    );


  /**
   * ==========================================================
   * 🔎 FILTRO
   * ==========================================================
   */

  if (!aba.getFilter()) {

    aba
      .getRange(
        4,
        1,
        Math.max(
          ultimaLinha - 3,
          2
        ),
        9
      )
      .createFilter();

  }


  /**
   * ==========================================================
   * ✅ FINAL
   * ==========================================================
   */

  SpreadsheetApp.flush();


  ss.toast(
    'Reembolsos Julius V2 aplicados 💸👀',
    'Controle Financeiro',
    4
  );

}


/**
 * ============================================================
 * 🔧 DESFAZ MESCLAGENS COM SEGURANÇA
 * ============================================================
 */

function reembVDesfazerMesclasNaArea_(
  aba,
  intervalo
) {

  const area =
    aba.getRange(intervalo);


  const mesclagens =
    area.getMergedRanges();


  mesclagens.forEach(range => {

    range.breakApart();

  });

}


/**
 * ============================================================
 * 🧱 CARD
 * ============================================================
 */

function reembVCard_(
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


  reembVBorda_(
    range,
    borda
  );

}


/**
 * ============================================================
 * 🚦 FORMATAÇÃO CONDICIONAL
 * ============================================================
 */

function reembVFormatacaoCondicional_(
  aba,
  C,
  ultimaLinha
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
            range.getRow() >= 5
          );

        });

    });


  const novas = [];


  /**
   * 🔴 PENDENTE
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('🔴')
      .setBackground(C.vermelhoClaro)
      .setFontColor(C.vermelho)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'H5:H' + ultimaLinha
        )
      ])
      .build()

  );


  /**
   * 🟡 PARCIAL
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('🟡')
      .setBackground(C.amareloClaro)
      .setFontColor(C.marromEscuro)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'H5:H' + ultimaLinha
        )
      ])
      .build()

  );


  /**
   * ✅ PAGO
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('✅')
      .setBackground(C.verdeClaro)
      .setFontColor(C.verde)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'H5:H' + ultimaLinha
        )
      ])
      .build()

  );


  /**
   * 💚 SALDO ZERO
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setFontColor(C.verde)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'G5:G' + ultimaLinha
        )
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

function reembVBorda_(
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