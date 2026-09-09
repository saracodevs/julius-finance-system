/**
 * ============================================================
 * 🎯 METAS VISUAL V3 — JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * Execute somente:
 * montarMetasVisual()
 *
 * V3:
 * - Corrige conflito de células mescladas.
 * - Pode ser executado várias vezes.
 * - Mantém a tabela principal A3:H8 intacta.
 * - Reconstrói apenas o visual superior e progresso inferior.
 * - Não altera as fórmulas principais das metas.
 */

function montarMetasVisual() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('🎯 Metas');

  if (!aba) {
    throw new Error(
      'A aba 🎯 Metas não foi encontrada.'
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

    marrom: '#654229',
    marromEscuro: '#3C2A1E',

    texto: '#33291F',
    textoSecundario: '#786554',
    branco: '#FFF9F0',

    verde: '#4F7A3B',
    verdeClaro: '#DDE8D4',

    amarelo: '#A77E18',
    amareloClaro: '#F3E7BD',

    vermelho: '#B54332',
    vermelhoClaro: '#F2D5CC',

    cinza: '#8B8178',
    cinzaClaro: '#E8E1D8',

    borda: '#C3A989',

    editavel: '#FFF1C9'
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
    .setBackground(C.fundo)
    .setFontFamily('Arial')
    .setFontColor(C.texto)
    .setVerticalAlignment('middle');


  /**
   * ==========================================================
   * 📐 COLUNAS
   * ==========================================================
   */

  aba.setColumnWidth(1, 245);
  aba.setColumnWidth(2, 135);
  aba.setColumnWidth(3, 135);
  aba.setColumnWidth(4, 115);
  aba.setColumnWidth(5, 145);
  aba.setColumnWidth(6, 145);
  aba.setColumnWidth(7, 185);
  aba.setColumnWidth(8, 135);


  /**
   * ==========================================================
   * 📏 ALTURAS
   * ==========================================================
   */

  aba.setRowHeight(1, 48);
  aba.setRowHeight(2, 52);
  aba.setRowHeight(3, 40);


  for (
    let linha = 4;
    linha <= 8;
    linha++
  ) {

    aba.setRowHeight(
      linha,
      42
    );
  }


  aba.setRowHeight(9, 12);
  aba.setRowHeight(10, 10);
  aba.setRowHeight(11, 34);


  /**
   * Progresso compacto
   */

  for (
    let linha = 12;
    linha <= 16;
    linha++
  ) {

    aba.setRowHeight(
      linha,
      38
    );
  }


  for (
    let linha = 17;
    linha <= 25;
    linha++
  ) {

    aba.setRowHeight(
      linha,
      18
    );
  }


  /**
   * ==========================================================
   * 🔧 DESFAZ MESCLAGENS EXISTENTES
   * ==========================================================
   *
   * Isso torna o código seguro para executar novamente.
   */

  metasV3DesfazerMesclas_(
    aba,
    'A2:H2'
  );


  metasV3DesfazerMesclas_(
    aba,
    'A11:H25'
  );


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
    .setValue('🎯 METAS FINANCEIRAS')
    .setFontSize(21)
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 📊 CARDS SUPERIORES
   * ==========================================================
   */

  aba
    .getRange('A2:H2')
    .clearContent()
    .setBackground(C.fundo);


  /**
   * 💰 TOTAL
   */

  aba
    .getRange('A2:B2')
    .merge();


  metasV3Card_(
    aba,
    'A2:B2',
    C.card,
    C.borda
  );


  aba
    .getRange('A2')
    .setFormula(
      '="💰 TOTAL DAS METAS  •  "&' +
      'TEXT(SUM(B4:B8);"R$ #,##0.00")'
    )
    .setFontWeight('bold')
    .setFontColor(C.azulPetroleo)
    .setHorizontalAlignment('center');


  /**
   * ✅ ACUMULADO
   */

  aba
    .getRange('C2:D2')
    .merge();


  metasV3Card_(
    aba,
    'C2:D2',
    C.verdeClaro,
    C.borda
  );


  aba
    .getRange('C2')
    .setFormula(
      '="✅ JÁ ACUMULADO  •  "&' +
      'TEXT(SUM(C4:C8);"R$ #,##0.00")'
    )
    .setFontWeight('bold')
    .setFontColor(C.verde)
    .setHorizontalAlignment('center');


  /**
   * 💸 FALTA
   */

  aba
    .getRange('E2:F2')
    .merge();


  metasV3Card_(
    aba,
    'E2:F2',
    C.vermelhoClaro,
    C.borda
  );


  aba
    .getRange('E2')
    .setFormula(
      '="💸 AINDA FALTA  •  "&' +
      'TEXT(' +
      'MAX(SUM(B4:B8)-SUM(C4:C8);0);' +
      '"R$ #,##0.00"' +
      ')'
    )
    .setFontWeight('bold')
    .setFontColor(C.vermelho)
    .setHorizontalAlignment('center');


  /**
   * 📥 APORTE
   */

  aba
    .getRange('G2:H2')
    .merge();


  metasV3Card_(
    aba,
    'G2:H2',
    C.amareloClaro,
    C.borda
  );


  aba
    .getRange('G2')
    .setFormula(
      '="📥 APORTE MENSAL  •  "&' +
      'TEXT(SUM(H4:H8);"R$ #,##0.00")'
    )
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 📋 CABEÇALHO PRINCIPAL
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
   * 🎯 TABELA DAS METAS
   * ==========================================================
   */

  for (
    let linha = 4;
    linha <= 8;
    linha++
  ) {

    const percentual =
      Number(
        aba
          .getRange(
            linha,
            4
          )
          .getValue()
      ) || 0;


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


    metasV3Borda_(
      faixa,
      C.borda
    );


    /**
     * Nome
     */

    aba
      .getRange(
        linha,
        1
      )
      .setFontWeight('bold')
      .setFontColor(C.azulEscuro);


    /**
     * % concluído
     */

    const celulaPercentual =
      aba.getRange(
        linha,
        4
      );


    celulaPercentual
      .setFontSize(11)
      .setFontWeight('bold');


    if (
      percentual >= 1
    ) {

      celulaPercentual
        .setBackground('#7DBB6A')
        .setFontColor('#204A18');

    } else if (
      percentual >= 0.75
    ) {

      celulaPercentual
        .setBackground(C.azulClaro)
        .setFontColor(C.azulEscuro);

    } else if (
      percentual >= 0.50
    ) {

      celulaPercentual
        .setBackground('#E2C95B')
        .setFontColor('#604F00');

    } else {

      celulaPercentual
        .setBackground(C.vermelhoClaro)
        .setFontColor(C.vermelho);

    }

  }


  /**
   * ==========================================================
   * ✏️ CAMPOS EDITÁVEIS
   * ==========================================================
   *
   * B = Valor Alvo
   * C = Valor Atual
   * H = Aporte Mensal
   */

  aba
    .getRange('B4:C8')
    .setBackground(C.editavel)
    .setFontWeight('bold');


  aba
    .getRange('H4:H8')
    .setBackground(C.editavel)
    .setFontWeight('bold');


  /**
   * ==========================================================
   * 💵 FORMATOS
   * ==========================================================
   */

  aba
    .getRange('B4:C8')
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  aba
    .getRange('D4:D8')
    .setNumberFormat(
      '0.0%'
    );


  aba
    .getRange('E4:E8')
    .setNumberFormat(
      'mmm/yyyy'
    );


  aba
    .getRange('H4:H8')
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  /**
   * ==========================================================
   * 🎯 ALINHAMENTO PRINCIPAL
   * ==========================================================
   */

  aba
    .getRange('A4:A8')
    .setHorizontalAlignment('left');


  aba
    .getRange('B4:F8')
    .setHorizontalAlignment('center');


  aba
    .getRange('G4:G8')
    .setHorizontalAlignment('left');


  aba
    .getRange('H4:H8')
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 🚦 STATUS
   * ==========================================================
   */

  metasV3Status_(
    aba,
    C
  );


  /**
   * ==========================================================
   * 📊 PROGRESSO DAS METAS
   * ==========================================================
   */

  aba
    .getRange('A11:H25')
    .clearContent()
    .setBackground(C.fundo);


  /**
   * Cabeçalho
   */

  aba
    .getRange('A11:H11')
    .setBackground(C.marrom)
    .setFontColor(C.branco);


  aba
    .getRange('A11')
    .setValue('📊 PROGRESSO DAS METAS')
    .setFontWeight('bold')
    .setFontSize(11);


  /**
   * Cada linha do painel usa diretamente
   * os dados da tabela principal.
   */

  const linhasOrigem =
    [4, 5, 6, 7, 8];


  linhasOrigem.forEach(
    (linhaOrigem, indice) => {

      const linhaDestino =
        12 + indice;


      const nome =
        aba
          .getRange(
            linhaOrigem,
            1
          )
          .getDisplayValue();


      const percentual =
        Number(
          aba
            .getRange(
              linhaOrigem,
              4
            )
            .getValue()
        ) || 0;


      /**
       * --------------------------------
       * NOME
       * --------------------------------
       */

      aba
        .getRange(
          linhaDestino,
          1,
          1,
          2
        )
        .merge();


      aba
        .getRange(
          linhaDestino,
          1
        )
        .setValue(nome)
        .setBackground(C.card)
        .setFontWeight('bold')
        .setFontColor(C.marromEscuro)
        .setHorizontalAlignment('left');


      /**
       * --------------------------------
       * BARRA
       * --------------------------------
       */

      aba
        .getRange(
          linhaDestino,
          3,
          1,
          4
        )
        .merge();


      aba
        .getRange(
          linhaDestino,
          3
        )
        .setValue(
          metasV3Barra_(
            percentual
          )
        )
        .setBackground(C.card)
        .setFontFamily('Courier New')
        .setFontWeight('bold')
        .setFontSize(11)
        .setFontColor(
          metasV3CorPercentual_(
            percentual,
            C
          )
        )
        .setHorizontalAlignment('center');


      /**
       * --------------------------------
       * PERCENTUAL
       * --------------------------------
       */

      aba
        .getRange(
          linhaDestino,
          7,
          1,
          2
        )
        .merge();


      aba
        .getRange(
          linhaDestino,
          7
        )
        .setValue(
          percentual
        )
        .setNumberFormat(
          '0.0%'
        )
        .setBackground(C.card)
        .setFontWeight('bold')
        .setFontSize(11)
        .setFontColor(
          metasV3CorPercentual_(
            percentual,
            C
          )
        )
        .setHorizontalAlignment('center');


      /**
       * Borda da linha completa
       */

      metasV3Borda_(
        aba.getRange(
          linhaDestino,
          1,
          1,
          8
        ),
        C.borda
      );

    }
  );


  /**
   * ==========================================================
   * 🧹 ÁREA ABAIXO
   * ==========================================================
   */

  aba
    .getRange('A17:H25')
    .setBackground(C.fundo)
    .clearContent();


  /**
   * ==========================================================
   * 📝 NOTAS
   * ==========================================================
   */

  aba
    .getRange('B3')
    .setNote(
      'Valor total necessário para concluir a meta.'
    );


  aba
    .getRange('C3')
    .setNote(
      'Valor já acumulado atualmente.'
    );


  aba
    .getRange('H3')
    .setNote(
      'Quanto pretende aportar por mês nesta meta.'
    );


  /**
   * ==========================================================
   * ✅ FINAL
   * ==========================================================
   */

  SpreadsheetApp.flush();


  ss.toast(
    'Metas Julius V3 aplicadas 🎯👀',
    'Controle Financeiro',
    4
  );

}


/**
 * ============================================================
 * 🔧 DESFAZER MESCLAGENS
 * ============================================================
 */

function metasV3DesfazerMesclas_(
  aba,
  intervalo
) {

  const area =
    aba.getRange(
      intervalo
    );


  const mesclagens =
    area.getMergedRanges();


  mesclagens.forEach(
    range => {

      range.breakApart();

    }
  );

}


/**
 * ============================================================
 * 📊 BARRA
 * ============================================================
 */

function metasV3Barra_(
  percentual
) {

  const tamanho =
    24;


  const valor =
    Math.max(
      0,
      Math.min(
        Number(
          percentual
        ) || 0,
        1
      )
    );


  const preenchidos =
    Math.round(
      valor *
      tamanho
    );


  const vazios =
    tamanho -
    preenchidos;


  return (
    '█'.repeat(
      preenchidos
    ) +
    '░'.repeat(
      vazios
    )
  );

}


/**
 * ============================================================
 * 🎨 COR DO PROGRESSO
 * ============================================================
 */

function metasV3CorPercentual_(
  percentual,
  C
) {

  if (
    percentual >= 1
  ) {
    return C.verde;
  }


  if (
    percentual >= 0.75
  ) {
    return C.azulPetroleo;
  }


  if (
    percentual >= 0.50
  ) {
    return C.amarelo;
  }


  return C.vermelho;

}


/**
 * ============================================================
 * 🧱 CARD
 * ============================================================
 */

function metasV3Card_(
  aba,
  intervalo,
  fundo,
  borda
) {

  const range =
    aba.getRange(
      intervalo
    );


  range
    .setBackground(
      fundo
    )
    .setVerticalAlignment(
      'middle'
    );


  metasV3Borda_(
    range,
    borda
  );

}


/**
 * ============================================================
 * 🚦 STATUS
 * ============================================================
 */

function metasV3Status_(
  aba,
  C
) {

  const atuais =
    aba.getConditionalFormatRules();


  const preservadas =
    atuais.filter(
      regra => {

        return !regra
          .getRanges()
          .some(
            range => {

              return (
                range.getColumn() === 6 &&
                range.getRow() >= 4 &&
                range.getRow() <= 8
              );

            }
          );

      }
    );


  const novas = [];


  /**
   * ✅ CONCLUÍDA
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('✅')
      .setBackground(C.verdeClaro)
      .setFontColor(C.verde)
      .setBold(true)
      .setRanges([
        aba.getRange('F4:F8')
      ])
      .build()

  );


  /**
   * 🟢 CAMINHO
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('🟢')
      .setBackground(C.azulClaro)
      .setFontColor(C.azulEscuro)
      .setBold(true)
      .setRanges([
        aba.getRange('F4:F8')
      ])
      .build()

  );


  /**
   * 🟡 PROGRESSO
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('🟡')
      .setBackground(C.amareloClaro)
      .setFontColor(C.marromEscuro)
      .setBold(true)
      .setRanges([
        aba.getRange('F4:F8')
      ])
      .build()

  );


  /**
   * 🔴 INÍCIO
   */

  novas.push(

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenTextContains('🔴')
      .setBackground(C.vermelhoClaro)
      .setFontColor(C.vermelho)
      .setBold(true)
      .setRanges([
        aba.getRange('F4:F8')
      ])
      .build()

  );


  aba.setConditionalFormatRules(
    preservadas.concat(
      novas
    )
  );

}


/**
 * ============================================================
 * 🔲 BORDA
 * ============================================================
 */

function metasV3Borda_(
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