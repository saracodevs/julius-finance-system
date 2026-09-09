/**
 * ============================================================
 * 💰 RECEITAS VISUAL — JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * Execute somente:
 * montarReceitasVisual()
 *
 * - Não altera os dados existentes.
 * - Não altera fórmulas.
 * - Não cria onEdit.
 * - Mantém a tabela original.
 */

function montarReceitasVisual() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    ss.getSheetByName('💰 Receitas');


  if (!aba) {

    throw new Error(
      'A aba 💰 Receitas não foi encontrada.'
    );

  }


  const C = {

    fundo: '#F1E4CF',
    fundoClaro: '#F8F0E3',
    card: '#FFF8EC',

    azulPetroleo: '#286173',
    azulEscuro: '#194754',

    caramelo: '#C78342',

    marromEscuro: '#3C2A1E',

    texto: '#33291F',
    textoSecundario: '#786554',
    branco: '#FFF9F0',

    verde: '#4F7A3B',
    verdeClaro: '#DDE8D4',

    amareloClaro: '#F3E7BD',

    borda: '#C3A989'
  };


  /**
   * ==========================================================
   * BASE
   * ==========================================================
   */

  const ultimaLinha =
    Math.max(
      aba.getLastRow(),
      4
    );


  const qtdLinhas =
    Math.max(
      ultimaLinha - 3,
      1
    );


  aba.setHiddenGridlines(true);

  aba.setFrozenRows(3);


  aba
    .getRange(
      1,
      1,
      ultimaLinha,
      6
    )
    .setFontFamily('Arial')
    .setFontColor(C.texto)
    .setVerticalAlignment('middle');


  /**
   * ==========================================================
   * COLUNAS
   * ==========================================================
   */

  aba.setColumnWidth(1, 115);
  aba.setColumnWidth(2, 230);
  aba.setColumnWidth(3, 140);
  aba.setColumnWidth(4, 135);
  aba.setColumnWidth(5, 135);
  aba.setColumnWidth(6, 115);


  /**
   * ==========================================================
   * ALTURAS
   * ==========================================================
   */

  aba.setRowHeight(1, 48);
  aba.setRowHeight(2, 54);
  aba.setRowHeight(3, 40);

  aba.setRowHeights(
    4,
    qtdLinhas,
    34
  );


  /**
   * ==========================================================
   * TÍTULO
   * ==========================================================
   */

  receitasDesfazerMesclas_(
    aba,
    'A1:F2'
  );


  aba
    .getRange('A1:F1')
    .merge()
    .setValue('💰 RECEITAS')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco)
    .setFontWeight('bold')
    .setFontSize(21)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * CARDS
   * ==========================================================
   */

  aba
    .getRange('A2:F2')
    .clearContent()
    .setBackground(C.fundo);


  /**
   * RECEITA DO MÊS
   */

  aba
    .getRange('A2:B2')
    .merge()
    .setBackground(C.card);


  aba
    .getRange('A2')
    .setFormula(
      '="💰 RECEITA DO MÊS  •  "&' +
      'TEXT(' +
      'SUMIFS(' +
      'D4:D' + ultimaLinha + ';' +
      'A4:A' + ultimaLinha + ';">="&EOMONTH(\'📊 Dashboard\'!B7;-1)+1;' +
      'A4:A' + ultimaLinha + ';"<="&EOMONTH(\'📊 Dashboard\'!B7;0)' +
      ');' +
      '"R$ #,##0.00"' +
      ')'
    )
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor(C.verde)
    .setHorizontalAlignment('center');


  /**
   * RECORRENTES
   */

  aba
    .getRange('C2:D2')
    .merge()
    .setBackground(C.verdeClaro);


  aba
    .getRange('C2')
    .setFormula(
      '="🔁 RECORRENTES  •  "&' +
      'TEXT(' +
      'SUMIFS(' +
      'D4:D' + ultimaLinha + ';' +
      'F4:F' + ultimaLinha + ';"Sim";' +
      'A4:A' + ultimaLinha + ';">="&EOMONTH(\'📊 Dashboard\'!B7;-1)+1;' +
      'A4:A' + ultimaLinha + ';"<="&EOMONTH(\'📊 Dashboard\'!B7;0)' +
      ');' +
      '"R$ #,##0.00"' +
      ')'
    )
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor(C.azulPetroleo)
    .setHorizontalAlignment('center');


  /**
   * QUANTIDADE
   */

  aba
    .getRange('E2:F2')
    .merge()
    .setBackground(C.amareloClaro);


  aba
    .getRange('E2')
    .setFormula(
      '="🧾 ENTRADAS  •  "&' +
      'COUNTIFS(' +
      'A4:A' + ultimaLinha + ';">="&EOMONTH(\'📊 Dashboard\'!B7;-1)+1;' +
      'A4:A' + ultimaLinha + ';"<="&EOMONTH(\'📊 Dashboard\'!B7;0)' +
      ')'
    )
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center');


  receitasBorda_(
    aba.getRange('A2:F2'),
    C.borda
  );


  /**
   * ==========================================================
   * CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A3:F3')
    .setValues([[
      '📅 Data',
      '📝 Descrição',
      '🏷️ Categoria',
      '💰 Valor',
      '📌 Tipo',
      '🔁 Recorrente?'
    ]]);


  aba
    .getRange('A3:F3')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco)
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * CORPO
   * ==========================================================
   */

  const corpo =
    aba.getRange(
      4,
      1,
      qtdLinhas,
      6
    );


  corpo
    .setBackground(C.card)
    .setFontSize(10);


  receitasBorda_(
    corpo,
    C.borda
  );


  /**
   * ==========================================================
   * FORMATOS
   * ==========================================================
   */

  aba
    .getRange(
      4,
      1,
      qtdLinhas,
      1
    )
    .setNumberFormat('dd/mm/yyyy')
    .setHorizontalAlignment('center');


  aba
    .getRange(
      4,
      4,
      qtdLinhas,
      1
    )
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    )
    .setFontWeight('bold')
    .setHorizontalAlignment('right');


  aba
    .getRange(
      4,
      2,
      qtdLinhas,
      2
    )
    .setHorizontalAlignment('left');


  aba
    .getRange(
      4,
      5,
      qtdLinhas,
      2
    )
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * DESTAQUE RECORRENTE
   * ==========================================================
   */

  receitasCondicionais_(
    aba,
    C,
    ultimaLinha
  );


  /**
   * ==========================================================
   * FINAL
   * ==========================================================
   */

  SpreadsheetApp.flush();


  ss.toast(
    'Receitas Julius aplicadas 💰👀',
    'Julius Finance',
    4
  );

}


/**
 * ============================================================
 * DESFAZER MESCLAGENS
 * ============================================================
 */

function receitasDesfazerMesclas_(
  aba,
  intervalo
) {

  aba
    .getRange(intervalo)
    .getMergedRanges()
    .forEach(
      range => range.breakApart()
    );

}


/**
 * ============================================================
 * CONDICIONAIS
 * ============================================================
 */

function receitasCondicionais_(
  aba,
  C,
  ultimaLinha
) {

  const regras =
    aba.getConditionalFormatRules();


  const preservadas =
    regras.filter(
      regra => {

        return !regra
          .getRanges()
          .some(
            range => {

              return (
                range.getColumn() === 6 &&
                range.getRow() >= 4
              );

            }
          );

      }
    );


  const recorrente =
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo('Sim')
      .setBackground(C.verdeClaro)
      .setFontColor(C.verde)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'F4:F' + ultimaLinha
        )
      ])
      .build();


  aba.setConditionalFormatRules(
    preservadas.concat([
      recorrente
    ])
  );

}


/**
 * ============================================================
 * BORDA
 * ============================================================
 */

function receitasBorda_(
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