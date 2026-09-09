/**
 * ============================================================
 * 🛒 DESPESAS VARIÁVEIS VISUAL V2
 * JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * Execute somente:
 * montarDespesasVariaveisVisual()
 *
 * CORREÇÕES V2:
 * - Remove limite fixo da linha 1000.
 * - Detecta automaticamente a última linha.
 * - Cards passam a considerar todos os lançamentos.
 * - Formatação de terceiro cobre toda a base atual.
 * - Valores altos cobrem toda a base atual.
 * - Mantém H:J ocultas e intactas.
 */

function montarDespesasVariaveisVisual() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('🛒 Desp.Variáveis');

  if (!aba) {
    throw new Error(
      'A aba 🛒 Desp.Variáveis não foi encontrada.'
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

    verde: '#668348',
    verdeClaro: '#E1E8D2',

    amarelo: '#C99A37',
    amareloClaro: '#F3E7BD',

    laranja: '#D98B3A',
    laranjaClaro: '#FAE2C3',

    vermelho: '#B9553D',
    vermelhoClaro: '#F2D5CC',

    borda: '#C3A989'
  };


  /**
   * ==========================================================
   * 🔎 TAMANHO REAL DA BASE
   * ==========================================================
   */

  const ultimaLinha =
    Math.max(
      aba.getLastRow(),
      7
    );

  const ultimaLinhaFormatacao =
    Math.max(
      aba.getMaxRows(),
      ultimaLinha
    );


  /**
   * ==========================================================
   * 🧱 BASE
   * ==========================================================
   */

  aba.setHiddenGridlines(true);

  aba.setFrozenRows(6);

  aba
    .getRange(
      1,
      1,
      aba.getMaxRows(),
      10
    )
    .setFontFamily('Arial');


  /**
   * ==========================================================
   * 📐 COLUNAS
   * ==========================================================
   */

  aba.setColumnWidth(1, 105); // Data
  aba.setColumnWidth(2, 155); // Categoria
  aba.setColumnWidth(3, 245); // Descrição
  aba.setColumnWidth(4, 125); // Valor
  aba.setColumnWidth(5, 145); // Pagamento
  aba.setColumnWidth(6, 135); // Cartão
  aba.setColumnWidth(7, 165); // Responsável


  /**
   * Colunas técnicas.
   * Não apagamos absolutamente nada delas.
   */

  aba.hideColumns(8, 3);


  /**
   * ==========================================================
   * 📏 ALTURAS
   * ==========================================================
   */

  aba.setRowHeight(1, 48);
  aba.setRowHeight(2, 24);
  aba.setRowHeight(3, 42);
  aba.setRowHeight(4, 26);
  aba.setRowHeight(5, 46);
  aba.setRowHeight(6, 38);


  /**
   * ==========================================================
   * 🔝 CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A1:G1')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco);


  aba
    .getRange('A1')
    .setValue('🛒 DESPESAS VARIÁVEIS')
    .setFontSize(21)
    .setFontWeight('bold');


  aba
    .getRange('A2:G2')
    .setBackground(C.caramelo)
    .setFontColor(C.marromEscuro);


  aba
    .getRange('A2')
    .setValue(
      'JULIUS FINANCE SYSTEM  •  cada comprinha também conta 👀'
    )
    .setFontSize(9)
    .setFontStyle('italic')
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 📅 COMPETÊNCIA
   * ==========================================================
   */

  dvV2Card_(
    aba,
    'A3:G3',
    C.card,
    C.borda
  );


  aba
    .getRange('A3')
    .setValue('📅 COMPETÊNCIA')
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(C.textoSecundario);


  /**
   * C3 continua com a fórmula original
   * ligada ao Dashboard.
   */

  aba
    .getRange('C3')
    .setFontSize(17)
    .setFontWeight('bold')
    .setFontColor(C.azulEscuro)
    .setHorizontalAlignment('center')
    .setNumberFormat('mmmm/yyyy');


  aba
    .getRange('F3')
    .setValue('🔄 Sincronizado com o Dashboard')
    .setFontSize(9)
    .setFontStyle('italic')
    .setFontColor(C.textoSecundario);


  /**
   * ==========================================================
   * 📊 RESUMOS
   * ==========================================================
   */


  /**
   * ----------------------------------------------------------
   * 💸 TOTAL DO MÊS
   * ----------------------------------------------------------
   *
   * Exclui automaticamente responsáveis configurados
   * como Terceiro / Reembolso.
   */

  dvV2Card_(
    aba,
    'A4:B5',
    C.card,
    C.borda
  );


  aba
    .getRange('A4')
    .setValue('💸 TOTAL DO MÊS')
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(C.vermelho);


  const formulaTotal =
    '=IFERROR(' +
      'SUMPRODUCT(' +

        '(A7:A' + ultimaLinha +
        '>=EOMONTH(C3;-1)+1)*' +

        '(A7:A' + ultimaLinha +
        '<=EOMONTH(C3;0))*' +

        'D7:D' + ultimaLinha + '*' +

        'N(' +
          'IFERROR(' +
            'MATCH(' +
              'G7:G' + ultimaLinha + ';' +
              'FILTER(' +
                '\'_CONFIG\'!A2:A;' +
                '\'_CONFIG\'!B2:B="Terceiro / Reembolso"' +
              ');' +
              '0' +
            ');' +
            '0' +
          ')=0' +
        ')' +

      ');' +
      '0' +
    ')';


  aba
    .getRange('B5')
    .setFormula(formulaTotal)
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    )
    .setFontSize(17)
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center');


  /**
   * ----------------------------------------------------------
   * 💳 NO CRÉDITO
   * ----------------------------------------------------------
   */

  dvV2Card_(
    aba,
    'C4:D5',
    C.card,
    C.borda
  );


  aba
    .getRange('C4')
    .setValue('💳 NO CRÉDITO')
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(C.caramelo);


  const formulaCredito =
    '=IFERROR(' +
      'SUMPRODUCT(' +

        '(A7:A' + ultimaLinha +
        '>=EOMONTH(C3;-1)+1)*' +

        '(A7:A' + ultimaLinha +
        '<=EOMONTH(C3;0))*' +

        '(E7:E' + ultimaLinha +
        '="Crédito")*' +

        'D7:D' + ultimaLinha + '*' +

        'N(' +
          'IFERROR(' +
            'MATCH(' +
              'G7:G' + ultimaLinha + ';' +
              'FILTER(' +
                '\'_CONFIG\'!A2:A;' +
                '\'_CONFIG\'!B2:B="Terceiro / Reembolso"' +
              ');' +
              '0' +
            ');' +
            '0' +
          ')=0' +
        ')' +

      ');' +
      '0' +
    ')';


  aba
    .getRange('D5')
    .setFormula(formulaCredito)
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    )
    .setFontSize(17)
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center');


  /**
   * ----------------------------------------------------------
   * 🧾 QUANTIDADE DE LANÇAMENTOS
   * ----------------------------------------------------------
   */

  dvV2Card_(
    aba,
    'E4:G5',
    C.card,
    C.borda
  );


  aba
    .getRange('E4')
    .setValue('🧾 LANÇAMENTOS')
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(C.azulPetroleo);


  const formulaQuantidade =
    '=COUNTIFS(' +

      'A7:A' + ultimaLinha + ';' +
      '">="&EOMONTH(C3;-1)+1;' +

      'A7:A' + ultimaLinha + ';' +
      '"<="&EOMONTH(C3;0)' +

    ')';


  aba
    .getRange('F5')
    .setFormula(formulaQuantidade)
    .setNumberFormat('0')
    .setFontSize(17)
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 📋 CABEÇALHO DA TABELA
   * ==========================================================
   */

  aba
    .getRange('A6:G6')
    .setValues([[
      '📅 Data',
      '🏷️ Categoria',
      '📝 Descrição',
      '💰 Valor',
      '💵 Pagamento',
      '💳 Cartão',
      '👤 Responsável'
    ]]);


  aba
    .getRange('A6:G6')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco)
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);


  /**
   * ==========================================================
   * 📋 LINHAS COM DADOS
   * ==========================================================
   */

  if (ultimaLinha >= 7) {

    for (
      let linha = 7;
      linha <= ultimaLinha;
      linha++
    ) {

      const fundo =
        linha % 2 === 0
          ? C.card
          : C.fundoClaro;


      const faixa =
        aba.getRange(
          linha,
          1,
          1,
          7
        );


      faixa
        .setBackground(fundo)
        .setFontColor(C.texto)
        .setFontSize(9)
        .setVerticalAlignment('middle');


      dvV2Borda_(
        faixa,
        C.borda
      );


      aba.setRowHeight(
        linha,
        30
      );

    }

  }


  /**
   * ==========================================================
   * 💵 FORMATAÇÃO
   * ==========================================================
   */

  const quantidadeLinhasDados =
    Math.max(
      ultimaLinha - 6,
      1
    );


  aba
    .getRange(
      7,
      1,
      quantidadeLinhasDados,
      1
    )
    .setNumberFormat('dd/mm/yyyy');


  aba
    .getRange(
      7,
      4,
      quantidadeLinhasDados,
      1
    )
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    )
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 🎯 ALINHAMENTO
   * ==========================================================
   */

  aba
    .getRange(
      7,
      1,
      quantidadeLinhasDados,
      1
    )
    .setHorizontalAlignment('center');


  aba
    .getRange(
      7,
      2,
      quantidadeLinhasDados,
      2
    )
    .setHorizontalAlignment('left');


  aba
    .getRange(
      7,
      4,
      quantidadeLinhasDados,
      4
    )
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 👤 RESPONSÁVEL
   * ==========================================================
   */

  aba
    .getRange(
      7,
      7,
      quantidadeLinhasDados,
      1
    )
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 📝 NOTAS
   * ==========================================================
   */

  aba
    .getRange('A6')
    .setNote(
      'Data da despesa.'
    );


  aba
    .getRange('B6')
    .setNote(
      'Categoria da despesa.'
    );


  aba
    .getRange('C6')
    .setNote(
      'Descrição livre do lançamento.'
    );


  aba
    .getRange('D6')
    .setNote(
      'Valor da despesa.'
    );


  aba
    .getRange('G6')
    .setNote(
      'O responsável define se o valor pertence ao orçamento principal ou a Terceiro / Reembolso.'
    );


  /**
   * ==========================================================
   * 🚦 FORMATAÇÃO CONDICIONAL
   * ==========================================================
   */

  dvV2FormatacaoCondicional_(
    aba,
    C,
    ultimaLinhaFormatacao
  );


  /**
   * ==========================================================
   * 🔎 FILTRO
   * ==========================================================
   */

  if (!aba.getFilter()) {

    aba
      .getRange(
        6,
        1,
        Math.max(
          ultimaLinha - 5,
          2
        ),
        7
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
    'Despesas Variáveis V2 atualizadas 🛒👀',
    'Controle Financeiro',
    4
  );

}


/**
 * ============================================================
 * 🧱 CARD
 * ============================================================
 */

function dvV2Card_(
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


  dvV2Borda_(
    range,
    borda
  );

}


/**
 * ============================================================
 * 🚦 FORMATAÇÃO CONDICIONAL
 * ============================================================
 */

function dvV2FormatacaoCondicional_(
  aba,
  C,
  ultimaLinha
) {

  const regrasExistentes =
    aba.getConditionalFormatRules();


  /**
   * Mantemos regras não relacionadas
   * aos nossos dois destaques.
   */

  const preservadas =
    regrasExistentes.filter(regra => {

      return !regra
        .getRanges()
        .some(range => {

          const coluna =
            range.getColumn();

          const linha =
            range.getRow();

          /**
           * Remove regras antigas que começavam
           * nas colunas D ou G a partir da linha 7.
           */

          return (
            linha >= 7 &&
            (
              coluna === 4 ||
              coluna === 7
            )
          );

        });

    });


  /**
   * ==========================================================
   * 👤 TERCEIRO / REEMBOLSO
   * ==========================================================
   */

  const regraTerceiro =
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND(' +
          '$G7<>"";' +
          'IFERROR(' +
            'VLOOKUP(' +
              '$G7;' +
              'INDIRECT("_CONFIG!A:B");' +
              '2;' +
              'FALSE' +
            ')="Terceiro / Reembolso";' +
            'FALSE' +
          ')' +
        ')'
      )
      .setBackground(C.amareloClaro)
      .setFontColor(C.marromEscuro)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'G7:G' + ultimaLinha
        )
      ])
      .build();


  /**
   * ==========================================================
   * 💸 VALOR ALTO
   * ==========================================================
   */

  const regraValorAlto =
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(1000)
      .setBackground(C.vermelhoClaro)
      .setFontColor(C.vermelho)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'D7:D' + ultimaLinha
        )
      ])
      .build();


  aba.setConditionalFormatRules(
    preservadas.concat([
      regraTerceiro,
      regraValorAlto
    ])
  );

}


/**
 * ============================================================
 * 🔲 BORDA
 * ============================================================
 */

function dvV2Borda_(
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