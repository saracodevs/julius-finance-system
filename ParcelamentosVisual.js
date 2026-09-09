/**
 * ============================================================
 * 💳 PARCELAMENTOS VISUAL V2 — JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * Execute somente:
 * montarParcelamentosVisual()
 *
 * V2:
 * - Cards usam somente A:N.
 * - TERCEIROS ganha mais espaço.
 * - M3 passa a se chamar STATUS.
 * - O:P continuam ocultas e intactas.
 * - Não altera fórmulas dos parcelamentos.
 * - Não altera Novo Parcelamento.
 * - Não cria onEdit.
 * - K "Pago" continua apenas informativa.
 */

function montarParcelamentosVisual() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('💳 Parcelamentos');

  if (!aba) {
    throw new Error(
      'A aba 💳 Parcelamentos não foi encontrada.'
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

    cinza: '#8B8178',
    cinzaClaro: '#E8E1D8',

    borda: '#C3A989'
  };


  /**
   * ==========================================================
   * 🔎 BASE
   * ==========================================================
   */

  const ultimaLinha =
    Math.max(
      aba.getLastRow(),
      4
    );

  const maxLinhas =
    aba.getMaxRows();


  aba.setHiddenGridlines(true);
  aba.setFrozenRows(3);


  aba
    .getRange(
      1,
      1,
      maxLinhas,
      16
    )
    .setFontFamily('Arial');


  /**
   * ==========================================================
   * 📐 LARGURAS
   * ==========================================================
   */

  aba.setColumnWidth(1, 145);   // A Categoria
  aba.setColumnWidth(2, 220);   // B Produto
  aba.setColumnWidth(3, 120);   // C Valor Total

  aba.setColumnWidth(4, 90);    // D Qtd
  aba.setColumnWidth(5, 125);   // E Parcela Atual
  aba.setColumnWidth(6, 125);   // F Restantes

  aba.setColumnWidth(7, 120);   // G Valor Parcela
  aba.setColumnWidth(8, 125);   // H Cartão
  aba.setColumnWidth(9, 150);   // I Responsável
  aba.setColumnWidth(10, 175);  // J Observação

  aba.setColumnWidth(11, 100);  // K Pago
  aba.setColumnWidth(12, 115);  // L Data
  aba.setColumnWidth(13, 115);  // M Status
  aba.setColumnWidth(14, 115);  // N Vigente


  /**
   * O:P são colunas técnicas.
   */

  aba.hideColumns(15, 2);


  /**
   * ==========================================================
   * 📏 ALTURAS
   * ==========================================================
   */

  aba.setRowHeight(1, 48);
  aba.setRowHeight(2, 52);
  aba.setRowHeight(3, 42);


  /**
   * ==========================================================
   * 🔝 TÍTULO
   * ==========================================================
   */

  aba
    .getRange('A1:N1')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco);


  aba
    .getRange('A1')
    .setValue('💳 CONTROLE DE PARCELAMENTOS')
    .setFontSize(21)
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 🧹 LIMPA SOMENTE O PAINEL DA LINHA 2
   * ==========================================================
   */

  aba
    .getRange('A2:P2')
    .breakApart()
    .clearContent()
    .setBackground(C.fundo);


  /**
   * ==========================================================
   * 💸 CARD 1 — PARCELAS DO MÊS
   * A:C
   * ==========================================================
   */

  parcelV2Card_(
    aba,
    'A2:C2',
    C.card,
    C.borda
  );

  aba
    .getRange('A2:C2')
    .merge();


  const formulaMes =
    '="💸  PARCELAS DO MÊS   •   "&' +
    'TEXT(' +
      'IFERROR(' +
        'SUMPRODUCT(' +
          '($N$4:$N$' + ultimaLinha + '="Sim")*' +
          '$G$4:$G$' + ultimaLinha + '*' +
          'N(' +
            'IFERROR(' +
              'MATCH(' +
                '$I$4:$I$' + ultimaLinha + ';' +
                'FILTER(' +
                  '\'_CONFIG\'!$A$2:$A;' +
                  '\'_CONFIG\'!$B$2:$B="Terceiro / Reembolso"' +
                ');' +
                '0' +
              ');' +
              '0' +
            ')=0' +
          ')' +
        ');' +
        '0' +
      ');' +
      '"R$ #,##0.00"' +
    ')';


  aba
    .getRange('A2')
    .setFormula(formulaMes)
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(C.vermelho)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 🧾 CARD 2 — ATIVOS
   * D:F
   * ==========================================================
   */

  parcelV2Card_(
    aba,
    'D2:F2',
    C.card,
    C.borda
  );

  aba
    .getRange('D2:F2')
    .merge();


  const formulaAtivos =
    '="🧾  ATIVOS   •   "&' +
    'IFERROR(' +
      'SUMPRODUCT(' +
        '($N$4:$N$' + ultimaLinha + '="Sim")*' +
        'N(' +
          'IFERROR(' +
            'MATCH(' +
              '$I$4:$I$' + ultimaLinha + ';' +
              'FILTER(' +
                '\'_CONFIG\'!$A$2:$A;' +
                '\'_CONFIG\'!$B$2:$B="Terceiro / Reembolso"' +
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
    .getRange('D2')
    .setFormula(formulaAtivos)
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(C.azulPetroleo)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * ⏳ CARD 3 — RESTANTES
   * G:J
   * ==========================================================
   */

  parcelV2Card_(
    aba,
    'G2:J2',
    C.card,
    C.borda
  );

  aba
    .getRange('G2:J2')
    .merge();


  const formulaRestantes =
    '="⏳  RESTANTES   •   "&' +
    'IFERROR(' +
      'SUMPRODUCT(' +
        '($N$4:$N$' + ultimaLinha + '="Sim")*' +
        '$F$4:$F$' + ultimaLinha + '*' +
        'N(' +
          'IFERROR(' +
            'MATCH(' +
              '$I$4:$I$' + ultimaLinha + ';' +
              'FILTER(' +
                '\'_CONFIG\'!$A$2:$A;' +
                '\'_CONFIG\'!$B$2:$B="Terceiro / Reembolso"' +
              ');' +
              '0' +
            ');' +
            '0' +
          ')=0' +
        ')' +
      ');' +
      '0' +
    ')&" parcelas"';


  aba
    .getRange('G2')
    .setFormula(formulaRestantes)
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(C.caramelo)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 👤 CARD 4 — TERCEIROS
   * K:N
   * ==========================================================
   */

  parcelV2Card_(
    aba,
    'K2:N2',
    C.amareloClaro,
    C.borda
  );

  aba
    .getRange('K2:N2')
    .merge();


  const formulaTerceiros =
    '="👤  TERCEIROS   •   "&' +
    'TEXT(' +
      'IFERROR(' +
        'SUMPRODUCT(' +
          '($N$4:$N$' + ultimaLinha + '="Sim")*' +
          '$G$4:$G$' + ultimaLinha + '*' +
          'N(' +
            'IFERROR(' +
              'MATCH(' +
                '$I$4:$I$' + ultimaLinha + ';' +
                'FILTER(' +
                  '\'_CONFIG\'!$A$2:$A;' +
                  '\'_CONFIG\'!$B$2:$B="Terceiro / Reembolso"' +
                ');' +
                '0' +
              ');' +
              '0' +
            ')>0' +
          ')' +
        ');' +
        '0' +
      ');' +
      '"R$ #,##0.00"' +
    ')';


  aba
    .getRange('K2')
    .setFormula(formulaTerceiros)
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 📋 CABEÇALHO
   * ==========================================================
   *
   * E3 e F3 possuem fórmulas dinâmicas.
   * Portanto não sobrescrevemos essas células.
   */

  aba
    .getRange('A3:N3')
    .setBackground(C.azulPetroleo)
    .setFontColor(C.branco)
    .setFontWeight('bold')
    .setFontSize(8)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);


  /**
   * Ajuste semântico:
   *
   * M contém Ativo / Finalizado.
   * Então "Status" representa melhor a coluna.
   */

  aba
    .getRange('M3')
    .setValue('Status');


  aba
    .getRange('N3')
    .setValue('Vigente no mês');


  /**
   * ==========================================================
   * 📋 CORPO DA TABELA
   * ==========================================================
   */

  const quantidadeLinhas =
    Math.max(
      ultimaLinha - 3,
      1
    );


  for (
    let linha = 4;
    linha <= ultimaLinha;
    linha++
  ) {

    const produto =
      String(
        aba
          .getRange(linha, 2)
          .getDisplayValue()
      ).trim();


    if (!produto) {

      aba
        .getRange(
          linha,
          1,
          1,
          14
        )
        .setBackground(C.fundo);

      continue;
    }


    const vigente =
      String(
        aba
          .getRange(linha, 14)
          .getDisplayValue()
      ).trim();


    const status =
      String(
        aba
          .getRange(linha, 13)
          .getDisplayValue()
      ).trim();


    const responsavel =
      String(
        aba
          .getRange(linha, 9)
          .getDisplayValue()
      ).trim();


    const terceiro =
      parcelV2ResponsavelTerceiro_(
        responsavel,
        ss
      );


    let fundo =
      linha % 2 === 0
        ? C.card
        : C.fundoClaro;


    /**
     * Terceiros ganham prioridade visual.
     */

    if (terceiro) {

      fundo =
        C.amareloClaro;

    } else if (
      status === 'Finalizado' ||
      vigente === 'Não'
    ) {

      fundo =
        C.cinzaClaro;

    }


    const faixa =
      aba.getRange(
        linha,
        1,
        1,
        14
      );


    faixa
      .setBackground(fundo)
      .setFontColor(
        status === 'Finalizado' &&
        !terceiro
          ? C.cinza
          : C.texto
      )
      .setFontSize(9)
      .setVerticalAlignment('middle');


    parcelV2Borda_(
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
   * 💵 FORMATOS
   * ==========================================================
   */

  aba
    .getRange(
      4,
      3,
      quantidadeLinhas,
      1
    )
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  aba
    .getRange(
      4,
      7,
      quantidadeLinhas,
      1
    )
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    )
    .setFontWeight('bold');


  aba
    .getRange(
      4,
      12,
      quantidadeLinhas,
      1
    )
    .setNumberFormat('dd/mm/yyyy');


  /**
   * ==========================================================
   * 🎯 ALINHAMENTOS
   * ==========================================================
   */

  aba
    .getRange(
      4,
      1,
      quantidadeLinhas,
      2
    )
    .setHorizontalAlignment('left');


  aba
    .getRange(
      4,
      10,
      quantidadeLinhas,
      1
    )
    .setHorizontalAlignment('left');


  aba
    .getRange(
      4,
      3,
      quantidadeLinhas,
      7
    )
    .setHorizontalAlignment('center');


  aba
    .getRange(
      4,
      11,
      quantidadeLinhas,
      4
    )
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 🔵 PARCELA ATUAL
   * ==========================================================
   */

  aba
    .getRange(
      4,
      5,
      quantidadeLinhas,
      1
    )
    .setFontWeight('bold')
    .setFontColor(C.azulPetroleo);


  /**
   * ==========================================================
   * ⏳ RESTANTES
   * ==========================================================
   */

  aba
    .getRange(
      4,
      6,
      quantidadeLinhas,
      1
    )
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 👤 RESPONSÁVEL
   * ==========================================================
   */

  aba
    .getRange(
      4,
      9,
      quantidadeLinhas,
      1
    )
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 🚦 FORMATAÇÃO CONDICIONAL
   * ==========================================================
   */

  parcelV2FormatacaoCondicional_(
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
    .getRange('E3')
    .setNote(
      'Parcela correspondente à competência selecionada no Dashboard.'
    );


  aba
    .getRange('F3')
    .setNote(
      'Quantidade de parcelas restantes após a competência atual.'
    );


  aba
    .getRange('K3')
    .setNote(
      'Campo manual e informativo. Pago/Devendo não determina se o parcelamento entra nos cálculos.'
    );


  aba
    .getRange('M3')
    .setNote(
      'Status automático do parcelamento na competência: Ativo ou Finalizado.'
    );


  aba
    .getRange('N3')
    .setNote(
      'Indica se o parcelamento pertence ao mês selecionado no Dashboard.'
    );


  /**
   * ==========================================================
   * 🔎 FILTRO
   * ==========================================================
   */

  if (!aba.getFilter()) {

    aba
      .getRange(
        3,
        1,
        Math.max(
          ultimaLinha - 2,
          2
        ),
        14
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
    'Parcelamentos Julius V2 aplicado 💳👀',
    'Controle Financeiro',
    4
  );

}


/**
 * ============================================================
 * 🧱 CARD
 * ============================================================
 */

function parcelV2Card_(
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


  parcelV2Borda_(
    range,
    borda
  );

}


/**
 * ============================================================
 * 👤 TERCEIRO?
 * ============================================================
 */

function parcelV2ResponsavelTerceiro_(
  responsavel,
  ss
) {

  if (!responsavel) {
    return false;
  }


  const config =
    ss.getSheetByName('_CONFIG');


  if (!config) {
    return false;
  }


  const ultima =
    config.getLastRow();


  if (ultima < 2) {
    return false;
  }


  const dados =
    config
      .getRange(
        2,
        1,
        ultima - 1,
        2
      )
      .getDisplayValues();


  for (
    let i = 0;
    i < dados.length;
    i++
  ) {

    const nome =
      String(
        dados[i][0]
      ).trim();


    const tipo =
      String(
        dados[i][1]
      ).trim();


    if (
      nome === responsavel &&
      tipo === 'Terceiro / Reembolso'
    ) {

      return true;
    }

  }


  return false;

}


/**
 * ============================================================
 * 🚦 FORMATAÇÃO CONDICIONAL
 * ============================================================
 */

function parcelV2FormatacaoCondicional_(
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

          const linha =
            range.getRow();

          const coluna =
            range.getColumn();


          return (
            linha >= 4 &&
            (
              coluna === 6 ||
              coluna === 9 ||
              coluna === 13 ||
              coluna === 14
            )
          );

        });

    });


  const novas = [];


  /**
   * 👤 TERCEIRO
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND(' +
          '$I4<>"";' +
          'IFERROR(' +
            'VLOOKUP(' +
              '$I4;' +
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
          'I4:I' + ultimaLinha
        )
      ])
      .build()

  );


  /**
   * 🟢 ATIVO
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo('Ativo')
      .setBackground(C.verdeClaro)
      .setFontColor(C.verde)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'M4:M' + ultimaLinha
        )
      ])
      .build()

  );


  /**
   * ⚪ FINALIZADO
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo('Finalizado')
      .setBackground(C.cinzaClaro)
      .setFontColor(C.cinza)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'M4:M' + ultimaLinha
        )
      ])
      .build()

  );


  /**
   * 🟢 VIGENTE
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextEqualTo('Sim')
      .setBackground(C.verdeClaro)
      .setFontColor(C.verde)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'N4:N' + ultimaLinha
        )
      ])
      .build()

  );


  /**
   * 🟠 ÚLTIMA PARCELA
   *
   * F = 0 e N = Sim:
   * estamos no mês da última parcela.
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($F4=0;$N4="Sim")'
      )
      .setBackground(C.laranjaClaro)
      .setFontColor(C.laranja)
      .setBold(true)
      .setRanges([
        aba.getRange(
          'F4:F' + ultimaLinha
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

function parcelV2Borda_(
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