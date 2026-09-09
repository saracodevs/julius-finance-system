/**
 * ============================================================
 * 📋 DESPESAS FIXAS VISUAL V3 — JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * Execute somente:
 * montarDespesasFixasVisual()
 *
 * IMPORTANTE:
 * - Checkbox marcado retorna "Pago"
 * - Checkbox desmarcado retorna "Devendo"
 * - Cards atualizam automaticamente
 * - Não cria onEdit
 * - Preserva dropdowns e checkboxes
 * - Não altera os dados da tabela
 */

function montarDespesasFixasVisual() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    ss.getSheetByName('📋 Desp.Fixas');


  if (!aba) {

    throw new Error(
      'A aba 📋 Desp.Fixas não foi encontrada.'
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

    caramelo: '#C78342',

    marromEscuro: '#3C2A1E',

    texto: '#33291F',
    textoSecundario: '#786554',
    branco: '#FFF9F0',

    verde: '#4F7A3B',
    verdeClaro: '#DDE8D4',

    amareloClaro: '#F3E7BD',

    vermelho: '#B54332',
    vermelhoClaro: '#F2D5CC',

    borda: '#C3A989'
  };


  /**
   * ==========================================================
   * 🔎 TAMANHO REAL DA TABELA
   * ==========================================================
   */

  const ultimaLinha =
    Math.max(
      aba.getLastRow(),
      6
    );


  const qtdLinhas =
    Math.max(
      ultimaLinha - 5,
      1
    );


  /**
   * ==========================================================
   * 🔧 DESFAZ MESCLAGENS DO TOPO
   * ==========================================================
   */

  despesasFixasV3DesfazerMesclas_(
    aba,
    'A1:H4'
  );


  /**
   * ==========================================================
   * 🧱 CONFIGURAÇÃO GERAL
   * ==========================================================
   */

  aba.setHiddenGridlines(true);

  aba.setFrozenRows(5);


  aba
    .getRange(
      1,
      1,
      ultimaLinha,
      8
    )
    .setFontFamily('Arial')
    .setFontColor(C.texto)
    .setVerticalAlignment('middle');


  /**
   * ==========================================================
   * 📐 COLUNAS
   * ==========================================================
   */

  aba.setColumnWidth(1, 180);
  aba.setColumnWidth(2, 145);
  aba.setColumnWidth(3, 210);
  aba.setColumnWidth(4, 105);
  aba.setColumnWidth(5, 135);
  aba.setColumnWidth(6, 180);
  aba.setColumnWidth(7, 220);
  aba.setColumnWidth(8, 90);


  /**
   * ==========================================================
   * 📏 ALTURAS
   * ==========================================================
   */

  aba.setRowHeight(1, 48);
  aba.setRowHeight(2, 22);
  aba.setRowHeight(3, 46);
  aba.setRowHeight(4, 58);
  aba.setRowHeight(5, 38);


  aba.setRowHeights(
    6,
    qtdLinhas,
    34
  );


  /**
   * ==========================================================
   * 🔝 TÍTULO
   * ==========================================================
   */

  aba
    .getRange('A1:H1')
    .merge()
    .setValue('📋 DESPESAS FIXAS')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco)
    .setFontWeight('bold')
    .setFontSize(21)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');


  /**
   * ==========================================================
   * 🟧 JULIUS FINANCE
   * ==========================================================
   */

  aba
    .getRange('A2:H2')
    .setBackground(C.caramelo);


  aba
    .getRange('A2')
    .setValue(
      'JULIUS FINANCE SYSTEM  •  conta fixa também merece fiscalização 👀'
    )
    .setFontSize(8)
    .setFontStyle('italic')
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro);


  /**
   * ==========================================================
   * 📅 COMPETÊNCIA
   * ==========================================================
   */

  aba
    .getRange('A3:B3')
    .merge()
    .setValue('📅 COMPETÊNCIA')
    .setBackground(C.card)
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center');


  aba
    .getRange('C3:D3')
    .merge();


  aba
    .getRange('C3')
    .setFormula(
      "='📊 Dashboard'!B7"
    )
    .setNumberFormat('mmmm/yyyy')
    .setBackground(C.card)
    .setFontColor(C.azulEscuro)
    .setFontSize(15)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');


  aba
    .getRange('E3:H3')
    .merge()
    .setValue(
      '🔄 Sincronizado com o Dashboard'
    )
    .setBackground(C.card)
    .setFontColor(C.textoSecundario)
    .setFontStyle('italic')
    .setFontSize(9)
    .setHorizontalAlignment('center');


  despesasFixasV3Borda_(
    aba.getRange('A3:H3'),
    C.borda
  );


  /**
   * ==========================================================
   * 💰 CARD — TOTAL DO MÊS
   * ==========================================================
   */

  aba
    .getRange('A4:B4')
    .merge()
    .setBackground(C.card);


  aba
    .getRange('A4')
    .setFormula(
      '="💰 TOTAL DO MÊS  •  "&' +
      'TEXT(' +
      'SUM(E6:E' + ultimaLinha + ');' +
      '"R$ #,##0.00"' +
      ')'
    )
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor(C.azulPetroleo)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * ✅ CARD — JÁ PAGO
   * ==========================================================
   */

  aba
    .getRange('C4:D4')
    .merge()
    .setBackground(C.verdeClaro);


  aba
    .getRange('C4')
    .setFormula(
      '="✅ JÁ PAGO  •  "&' +
      'TEXT(' +
      'SUMIF(' +
      'H6:H' + ultimaLinha + ';' +
      '"Pago";' +
      'E6:E' + ultimaLinha +
      ');' +
      '"R$ #,##0.00"' +
      ')'
    )
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor(C.verde)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 🔴 CARD — PENDENTE
   * ==========================================================
   */

  aba
    .getRange('E4:F4')
    .merge()
    .setBackground(C.vermelhoClaro);


  aba
    .getRange('E4')
    .setFormula(
      '="🔴 PENDENTE  •  "&' +
      'TEXT(' +
      'SUMIF(' +
      'H6:H' + ultimaLinha + ';' +
      '"Devendo";' +
      'E6:E' + ultimaLinha +
      ');' +
      '"R$ #,##0.00"' +
      ')'
    )
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor(C.vermelho)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 🧾 CARD — CONTAS
   * ==========================================================
   */

  aba
    .getRange('G4:H4')
    .merge()
    .setBackground(C.amareloClaro);


  aba
    .getRange('G4')
    .setFormula(
      '="🧾 CONTAS  •  "&' +
      'COUNTIF(' +
      'C6:C' + ultimaLinha + ';' +
      '"<>")'
    )
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center');


  despesasFixasV3Borda_(
    aba.getRange('A4:H4'),
    C.borda
  );


  /**
   * ==========================================================
   * 📋 CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A5:H5')
    .setValues([[
      '🏷️ Categoria',
      '👤 Responsável',
      '📝 Descrição',
      '📅 Vencimento',
      '💰 Valor',
      '💳 Forma de Pagamento',
      '📌 Observação',
      '✅ Pago'
    ]]);


  aba
    .getRange('A5:H5')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco)
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);


  /**
   * ==========================================================
   * 📋 CORPO DA TABELA
   * ==========================================================
   */

  const corpo =
    aba.getRange(
      6,
      1,
      qtdLinhas,
      8
    );


  corpo
    .setBackground(C.card)
    .setFontSize(10)
    .setFontColor(C.texto);


  despesasFixasV3Borda_(
    corpo,
    C.borda
  );


  /**
   * ==========================================================
   * 💰 VALOR
   * ==========================================================
   */

  aba
    .getRange(
      6,
      5,
      qtdLinhas,
      1
    )
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    )
    .setFontWeight('bold')
    .setHorizontalAlignment('right');


  /**
   * ==========================================================
   * 🎯 ALINHAMENTOS
   * ==========================================================
   */

  aba
    .getRange(
      6,
      1,
      qtdLinhas,
      3
    )
    .setHorizontalAlignment('left');


  aba
    .getRange(
      6,
      4,
      qtdLinhas,
      1
    )
    .setHorizontalAlignment('center');


  aba
    .getRange(
      6,
      6,
      qtdLinhas,
      1
    )
    .setHorizontalAlignment('center');


  aba
    .getRange(
      6,
      8,
      qtdLinhas,
      1
    )
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 🚦 PAGO / DEVENDO
   * ==========================================================
   */

  despesasFixasV3Condicionais_(
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
    .getRange('H5')
    .setNote(
      'Marque o checkbox para alterar de Devendo para Pago. Os cards do topo são atualizados automaticamente.'
    );


  aba
    .getRange('D5')
    .setNote(
      'Dia do vencimento. Valores como "x" são mantidos normalmente.'
    );


  /**
   * ==========================================================
   * ✅ FINAL
   * ==========================================================
   */

  SpreadsheetApp.flush();


  ss.toast(
    'Despesas Fixas V3 aplicadas 📋👀',
    'Julius Finance',
    4
  );

}


/**
 * ============================================================
 * 🔧 DESFAZER MESCLAGENS
 * ============================================================
 */

function despesasFixasV3DesfazerMesclas_(
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
 * 🚦 FORMATAÇÃO CONDICIONAL
 * ============================================================
 */

function despesasFixasV3Condicionais_(
  aba,
  C,
  ultimaLinha
) {

  const regrasAtuais =
    aba.getConditionalFormatRules();


  /**
   * Remove apenas regras anteriores aplicadas
   * pelo nosso visual nesta tabela.
   */

  const preservadas =
    regrasAtuais.filter(
      regra => {

        return !regra
          .getRanges()
          .some(
            range => {

              return (
                range.getRow() >= 6 &&
                range.getColumn() <= 8
              );

            }
          );

      }
    );


  const faixaCompleta =
    aba.getRange(
      'A6:H' + ultimaLinha
    );


  const colunaPago =
    aba.getRange(
      'H6:H' + ultimaLinha
    );


  /**
   * ==========================================================
   * ✅ PAGO
   * ==========================================================
   *
   * O checkbox marcado possui valor "Pago".
   */

  const regraPago =
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=$H6="Pago"'
      )
      .setBackground(
        C.verdeClaro
      )
      .setFontColor(
        C.verde
      )
      .setRanges([
        faixaCompleta
      ])
      .build();


  /**
   * ==========================================================
   * 🔴 DEVENDO
   * ==========================================================
   *
   * Só colore a célula do checkbox.
   */

  const regraDevendo =
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($C6<>"";$H6="Devendo")'
      )
      .setBackground(
        C.vermelhoClaro
      )
      .setRanges([
        colunaPago
      ])
      .build();


  aba.setConditionalFormatRules(
    preservadas.concat([
      regraPago,
      regraDevendo
    ])
  );

}


/**
 * ============================================================
 * 🔲 BORDA
 * ============================================================
 */

function despesasFixasV3Borda_(
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