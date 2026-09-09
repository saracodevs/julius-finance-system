/**
 * ============================================================
 * 📅 PLANEJAMENTO VISUAL V2 — JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * Execute somente:
 * montarPlanejamentoVisual()
 *
 * Esta versão:
 * - NÃO move fórmulas;
 * - NÃO apaga dados;
 * - NÃO altera os limites digitados;
 * - NÃO altera a lógica dos cartões;
 * - reorganiza somente o visual.
 */

function montarPlanejamentoVisual() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('📅 Planejamento Mensal');

  if (!aba) {
    throw new Error('A aba 📅 Planejamento Mensal não foi encontrada.');
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
   * 🧱 BASE
   * ==========================================================
   */

  aba.setHiddenGridlines(true);

  aba
    .getRange('A1:P45')
    .setBackground(C.fundo)
    .setFontFamily('Arial')
    .setFontColor(C.texto)
    .setVerticalAlignment('middle');


  /**
   * ==========================================================
   * 📐 COLUNAS
   * ==========================================================
   */

  const larguras = {

    A: 18,

    B: 185,
    C: 125,
    D: 125,
    E: 125,
    F: 100,
    G: 190,
    H: 135,

    I: 18,
    J: 24,

    K: 130,
    L: 110,
    M: 70,
    N: 70,
    O: 82,
    P: 90
  };

  Object.keys(larguras).forEach(coluna => {

    aba.setColumnWidth(
      pmV2ColunaParaNumero_(coluna),
      larguras[coluna]
    );

  });


  /**
   * ==========================================================
   * 📏 ALTURAS
   * ==========================================================
   */

  aba.setRowHeight(1, 48);
  aba.setRowHeight(2, 25);
  aba.setRowHeight(3, 10);

  aba.setRowHeight(4, 44);
  aba.setRowHeight(5, 14);

  aba.setRowHeight(6, 27);
  aba.setRowHeight(7, 48);

  aba.setRowHeight(8, 14);
  aba.setRowHeight(9, 10);

  aba.setRowHeight(10, 32);
  aba.setRowHeight(11, 32);
  aba.setRowHeight(12, 10);

  aba.setRowHeight(13, 36);

  for (let linha = 14; linha <= 19; linha++) {
    aba.setRowHeight(linha, 38);
  }

  aba.setRowHeight(20, 12);
  aba.setRowHeight(21, 12);
  aba.setRowHeight(22, 38);

  aba.setRowHeight(23, 14);
  aba.setRowHeight(24, 10);

  aba.setRowHeight(25, 34);
  aba.setRowHeight(26, 36);

  for (let linha = 27; linha <= 39; linha++) {
    aba.setRowHeight(linha, 34);
  }

  aba.setRowHeight(40, 12);
  aba.setRowHeight(41, 38);


  /**
   * ==========================================================
   * 🔝 CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A1:P1')
    .setBackground(C.azulPetroleo);

  aba
    .getRange('A1')
    .setValue('📅 PLANEJAMENTO FINANCEIRO')
    .setFontSize(22)
    .setFontWeight('bold')
    .setFontColor(C.branco)
    .setHorizontalAlignment('left');


  /**
   * ==========================================================
   * 🟤 SUBTÍTULO
   * ==========================================================
   */

  aba
    .getRange('A2:P2')
    .setBackground(C.caramelo);

  aba
    .getRange('A2')
    .setValue(
      'JULIUS FINANCE SYSTEM  •  planejar custa menos do que se arrepender 👀'
    )
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontStyle('italic')
    .setFontColor(C.marromEscuro);


  /**
   * ==========================================================
   * 📅 COMPETÊNCIA
   * ==========================================================
   */

  pmV2Card_(
    aba,
    'B4:H4',
    C.card,
    C.borda
  );

  aba
    .getRange('B4')
    .setValue('📅 COMPETÊNCIA')
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(C.textoSecundario);

  aba
    .getRange('D4')
    .setFontSize(18)
    .setFontWeight('bold')
    .setFontColor(C.azulEscuro)
    .setHorizontalAlignment('center')
    .setNumberFormat('mmmm/yyyy');

  aba
    .getRange('G4')
    .setValue('🔄 Sincronizado com o Dashboard')
    .setFontSize(9)
    .setFontStyle('italic')
    .setFontColor(C.textoSecundario);


  /**
   * ==========================================================
   * 💰 RESUMO GERAL
   * ==========================================================
   */

  pmV2Resumo_(
    aba,
    'B6:B7',
    'B6',
    'B7',
    '💰 ORÇAMENTO',
    C.azulPetroleo,
    C
  );

  pmV2Resumo_(
    aba,
    'D6:D7',
    'D6',
    'D7',
    '💸 GASTO REAL',
    C.vermelho,
    C
  );

  pmV2Resumo_(
    aba,
    'F6:F7',
    'F6',
    'F7',
    '💵 SALDO',
    C.verde,
    C
  );

  pmV2Resumo_(
    aba,
    'H6:H7',
    'H6',
    'H7',
    '📊 UTILIZADO',
    C.caramelo,
    C
  );

  ['B7', 'D7', 'F7'].forEach(celula => {

    aba
      .getRange(celula)
      .setNumberFormat(
        '"R$ " #,##0.00;-"R$ " #,##0.00'
      );

  });

  aba
    .getRange('H7')
    .setNumberFormat('0.0%');


  /**
   * ==========================================================
   * 🚦 CORES DO RESUMO
   * ==========================================================
   */

  const saldo =
    Number(
      aba.getRange('F7').getValue()
    ) || 0;

  const uso =
    Number(
      aba.getRange('H7').getValue()
    ) || 0;


  aba
    .getRange('F7')
    .setFontColor(
      saldo < 0
        ? C.vermelho
        : C.verde
    );


  aba
    .getRange('H7')
    .setFontColor(
      uso > 1
        ? C.vermelho
        : uso >= 0.85
          ? C.laranja
          : C.verde
    );


  /**
   * ==========================================================
   * 👀 BLOCO PRINCIPAL
   * ==========================================================
   */

  aba
    .getRange('B10:H10')
    .setBackground(C.marrom)
    .setFontColor(C.branco);

  aba
    .getRange('B10')
    .setValue('👀 JULIUS CONTROLA OS LIMITES')
    .setFontSize(12)
    .setFontWeight('bold');


  aba
    .getRange('B11:H11')
    .setBackground(C.fundoClaro);

  aba
    .getRange('B11')
    .setValue(
      '✏️ Ajuste somente o limite mensal. Gastos, saldo e status são automáticos.'
    )
    .setFontSize(9)
    .setFontStyle('italic')
    .setFontColor(C.textoSecundario);


  /**
   * ==========================================================
   * 📋 CABEÇALHO CATEGORIAS
   * ==========================================================
   */

  aba
    .getRange('B13:H13')
    .setValues([[
      'Categoria',
      'Limite',
      'Gasto Real',
      'Disponível',
      '% Usado',
      'Progresso',
      'Status'
    ]]);


  pmV2Cabecalho_(
    aba,
    'B13:H13',
    C.azulPetroleo,
    C.branco
  );


  /**
   * ==========================================================
   * 📊 CATEGORIAS
   * ==========================================================
   */

  for (let linha = 14; linha <= 19; linha++) {

    pmV2LinhaCategoria_(
      aba,
      linha,
      C
    );

  }


  /**
   * ==========================================================
   * ✏️ CAMPOS EDITÁVEIS
   * ==========================================================
   */

  aba
    .getRange('C14:C19')
    .setBackground('#FFF1C9')
    .setFontColor(C.marromEscuro)
    .setFontWeight('bold');


  /**
   * ==========================================================
   * TOTAL CATEGORIAS
   * ==========================================================
   */

  aba
    .getRange('B22:H22')
    .setBackground(C.azulClaro)
    .setFontColor(C.azulEscuro)
    .setFontWeight('bold')
    .setFontSize(10);

  pmV2Borda_(
    aba.getRange('B22:H22'),
    C.borda
  );


  /**
   * ==========================================================
   * 💳 CARTÕES — PAINEL LATERAL
   * ==========================================================
   */

  aba
    .getRange('K10:P10')
    .setBackground(C.marrom)
    .setFontColor(C.branco);

  aba
    .getRange('K10')
    .setValue('💳 COMPROMISSOS NOS CARTÕES')
    .setFontSize(11)
    .setFontWeight('bold');


  aba
    .getRange('K13:P13')
    .setValues([[
      'Cartão',
      'Valor',
      'Fecha',
      'Vence',
      'Pago?',
      'Anuidade'
    ]]);


  pmV2Cabecalho_(
    aba,
    'K13:P13',
    C.azulPetroleo,
    C.branco
  );


  for (let linha = 14; linha <= 21; linha++) {

    const faixa =
      aba.getRange(
        linha,
        11,
        1,
        6
      );

    faixa
      .setBackground(
        linha % 2 === 0
          ? C.card
          : C.fundoClaro
      )
      .setFontSize(9)
      .setFontColor(C.texto);

    pmV2Borda_(
      faixa,
      C.borda
    );

  }


  aba
    .getRange('L14:L21')
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  /**
   * ==========================================================
   * 🏠 DESPESAS FIXAS
   * ==========================================================
   */

  aba
    .getRange('B25:H25')
    .setBackground(C.marrom)
    .setFontColor(C.branco);

  aba
    .getRange('B25')
    .setValue('🏠 COMPROMISSOS FIXOS DO MÊS')
    .setFontSize(12)
    .setFontWeight('bold');


  aba
    .getRange('B26:H26')
    .setValues([[
      'Categoria',
      'Orçamento',
      'Comprometido',
      'Disponível',
      '% Receita',
      'Status',
      ''
    ]]);


  pmV2Cabecalho_(
    aba,
    'B26:H26',
    C.azulPetroleo,
    C.branco
  );


  /**
   * ==========================================================
   * 📋 LINHAS FIXAS
   * ==========================================================
   */

  for (let linha = 27; linha <= 39; linha++) {

    pmV2LinhaFixa_(
      aba,
      linha,
      C
    );

  }


  /**
   * ==========================================================
   * ✏️ ORÇAMENTO EDITÁVEL FIXAS
   * ==========================================================
   */

  aba
    .getRange('C27:C39')
    .setBackground('#FFF1C9')
    .setFontColor(C.marromEscuro)
    .setFontWeight('bold');


  /**
   * ==========================================================
   * TOTAL FIXAS
   * ==========================================================
   */

  aba
    .getRange('B41:H41')
    .setBackground(C.azulClaro)
    .setFontColor(C.azulEscuro)
    .setFontWeight('bold')
    .setFontSize(10);

  pmV2Borda_(
    aba.getRange('B41:H41'),
    C.borda
  );


  /**
   * ==========================================================
   * 💵 FORMATOS
   * ==========================================================
   */

  aba
    .getRange('C14:E22')
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );

  aba
    .getRange('F14:F22')
    .setNumberFormat('0.0%');


  aba
    .getRange('C27:E41')
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );

  aba
    .getRange('F27:F41')
    .setNumberFormat('0.0%');


  /**
   * ==========================================================
   * 🎯 ALINHAMENTOS
   * ==========================================================
   */

  aba
    .getRange('B14:B22')
    .setHorizontalAlignment('left');

  aba
    .getRange('C14:H22')
    .setHorizontalAlignment('center');


  aba
    .getRange('B27:B41')
    .setHorizontalAlignment('left');

  aba
    .getRange('C27:G41')
    .setHorizontalAlignment('center');


  aba
    .getRange('K14:K21')
    .setHorizontalAlignment('left');

  aba
    .getRange('L14:P21')
    .setHorizontalAlignment('center');


  /**
   * ==========================================================
   * 👀 DESTAQUE DO JULIUS — ALERTA PRINCIPAL
   * ==========================================================
   *
   * Usamos a área lateral inferior vazia sem mexer
   * nas tabelas nem nas fórmulas.
   */

  pmV2Card_(
    aba,
    'K25:P29',
    C.card,
    C.borda
  );

  aba
    .getRange('K25:P25')
    .setBackground(C.marrom);

  aba
    .getRange('K25')
    .setValue('👀 JULIUS ANALISA O MÊS')
    .setFontColor(C.branco)
    .setFontWeight('bold')
    .setFontSize(11);


  const totalOrcamento =
    Number(
      aba.getRange('B7').getValue()
    ) || 0;

  const gastoReal =
    Number(
      aba.getRange('D7').getValue()
    ) || 0;


  let mensagem1 = '';
  let mensagem2 = '';
  let mensagem3 = '';


  if (saldo < 0) {

    mensagem1 =
      '🔴 Planejamento ultrapassado em ' +
      pmV2Moeda_(Math.abs(saldo)) + '.';

  } else {

    mensagem1 =
      '🟢 Ainda restam ' +
      pmV2Moeda_(saldo) +
      ' no orçamento.';

  }


  if (uso > 1) {

    mensagem2 =
      '🔴 Você já utilizou ' +
      pmV2Percentual_(uso) +
      ' do orçamento.';

  } else if (uso >= 0.85) {

    mensagem2 =
      '🟠 Atenção: ' +
      pmV2Percentual_(uso) +
      ' do orçamento já foi usado.';

  } else {

    mensagem2 =
      '🟢 Uso atual: ' +
      pmV2Percentual_(uso) +
      '.';

  }


  let piorCategoria = '';
  let piorPercentual = -1;

  for (let linha = 14; linha <= 19; linha++) {

    const percentual =
      Number(
        aba.getRange(linha, 6).getValue()
      ) || 0;

    if (percentual > piorPercentual) {

      piorPercentual =
        percentual;

      piorCategoria =
        String(
          aba.getRange(linha, 2).getDisplayValue()
        );

    }

  }


  if (piorCategoria) {

    mensagem3 =
      '💸 Maior pressão: ' +
      piorCategoria +
      ' • ' +
      pmV2Percentual_(piorPercentual);

  }


  aba
    .getRange('K26:P26')
    .merge()
    .setValue(mensagem1)
    .setFontWeight('bold')
    .setFontSize(9)
    .setFontColor(
      saldo < 0
        ? C.vermelho
        : C.verde
    )
    .setWrap(true);


  aba
    .getRange('K27:P27')
    .merge()
    .setValue(mensagem2)
    .setFontWeight('bold')
    .setFontSize(9)
    .setFontColor(
      uso > 1
        ? C.vermelho
        : uso >= 0.85
          ? C.laranja
          : C.verde
    )
    .setWrap(true);


  aba
    .getRange('K28:P28')
    .merge()
    .setValue(mensagem3)
    .setFontWeight('bold')
    .setFontSize(9)
    .setFontColor(
      piorPercentual > 1
        ? C.vermelho
        : piorPercentual >= 0.90
          ? C.laranja
          : C.azulPetroleo
    )
    .setWrap(true);


  /**
   * ==========================================================
   * 📌 NOTAS
   * ==========================================================
   */

  aba
    .getRange('C13')
    .setNote(
      'Campo editável: informe o limite mensal desejado.'
    );

  aba
    .getRange('C26')
    .setNote(
      'Campo editável: informe o orçamento mensal desejado.'
    );


  /**
   * ==========================================================
   * 🚦 FORMATAÇÃO CONDICIONAL
   * ==========================================================
   */

  pmV2FormatacaoCondicional_(
    aba,
    C
  );


  /**
   * ==========================================================
   * 🧹 ESPAÇADORES
   * ==========================================================
   */

  aba
    .getRange('I1:J45')
    .setBackground(C.fundo);


  /**
   * ==========================================================
   * ✅ FINAL
   * ==========================================================
   */

  SpreadsheetApp.flush();

  ss.toast(
    'Planejamento Julius V2 aplicado 👀📅',
    'Controle Financeiro',
    4
  );
}


/**
 * ============================================================
 * 💰 CARD RESUMO
 * ============================================================
 */

function pmV2Resumo_(
  aba,
  intervalo,
  tituloCelula,
  valorCelula,
  titulo,
  destaque,
  C
) {

  pmV2Card_(
    aba,
    intervalo,
    C.card,
    C.borda
  );

  aba
    .getRange(tituloCelula)
    .setValue(titulo)
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(destaque)
    .setHorizontalAlignment('center');

  aba
    .getRange(valorCelula)
    .setFontSize(18)
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center');
}


/**
 * ============================================================
 * 🧱 CARD BASE
 * ============================================================
 */

function pmV2Card_(
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

  pmV2Borda_(
    range,
    borda
  );
}


/**
 * ============================================================
 * 📋 CABEÇALHO
 * ============================================================
 */

function pmV2Cabecalho_(
  aba,
  intervalo,
  fundo,
  texto
) {

  aba
    .getRange(intervalo)
    .setBackground(fundo)
    .setFontColor(texto)
    .setFontSize(9)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}


/**
 * ============================================================
 * 📊 LINHA CATEGORIA
 * ============================================================
 */

function pmV2LinhaCategoria_(
  aba,
  linha,
  C
) {

  const percentual =
    Number(
      aba.getRange(linha, 6).getValue()
    ) || 0;


  let fundo =
    linha % 2 === 0
      ? C.card
      : C.fundoClaro;


  if (percentual > 1) {

    fundo =
      C.vermelhoClaro;

  } else if (percentual >= 0.90) {

    fundo =
      C.laranjaClaro;

  } else if (percentual >= 0.75) {

    fundo =
      C.amareloClaro;

  }


  const range =
    aba.getRange(
      linha,
      2,
      1,
      7
    );

  range
    .setBackground(fundo)
    .setFontSize(10);


  pmV2Borda_(
    range,
    C.borda
  );


  aba
    .getRange(linha, 2)
    .setFontWeight('bold');


  /**
   * Barra de progresso
   */

  aba
    .getRange(linha, 7)
    .setFontFamily('Courier New')
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(
      percentual > 1
        ? C.vermelho
        : percentual >= 0.90
          ? C.laranja
          : percentual >= 0.75
            ? C.amarelo
            : C.verde
    );


  aba
    .getRange(linha, 8)
    .setFontWeight('bold');
}


/**
 * ============================================================
 * 🏠 LINHA FIXA
 * ============================================================
 */

function pmV2LinhaFixa_(
  aba,
  linha,
  C
) {

  const limite =
    Number(
      aba.getRange(linha, 3).getValue()
    ) || 0;

  const gasto =
    Number(
      aba.getRange(linha, 4).getValue()
    ) || 0;


  let fundo =
    linha % 2 === 0
      ? C.card
      : C.fundoClaro;


  if (
    limite > 0 &&
    gasto > limite
  ) {

    fundo =
      C.vermelhoClaro;
  }


  const range =
    aba.getRange(
      linha,
      2,
      1,
      6
    );


  range
    .setBackground(fundo)
    .setFontSize(10);


  pmV2Borda_(
    range,
    C.borda
  );


  aba
    .getRange(linha, 2)
    .setFontWeight('bold');


  aba
    .getRange(linha, 7)
    .setFontWeight('bold');

}


/**
 * ============================================================
 * 🚦 FORMATAÇÃO CONDICIONAL
 * ============================================================
 */

function pmV2FormatacaoCondicional_(
  aba,
  C
) {

  /**
   * Evita criar regras Julius duplicadas toda vez
   * que executar o script.
   */

  const existentes =
    aba.getConditionalFormatRules();


  const filtradas =
    existentes.filter(regra => {

      const ranges =
        regra.getRanges();

      return !ranges.some(range => {

        const a1 =
          range.getA1Notation();

        return (
          a1 === 'H14:H22' ||
          a1 === 'G27:G41'
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
        aba.getRange('H14:H22'),
        aba.getRange('G27:G41')
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
        aba.getRange('H14:H22'),
        aba.getRange('G27:G41')
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
        aba.getRange('H14:H22'),
        aba.getRange('G27:G41')
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
        aba.getRange('H14:H22'),
        aba.getRange('G27:G41')
      ])
      .build()

  );


  aba.setConditionalFormatRules(
    filtradas.concat(novas)
  );
}


/**
 * ============================================================
 * 🔲 BORDA
 * ============================================================
 */

function pmV2Borda_(
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


/**
 * ============================================================
 * 💵 MOEDA
 * ============================================================
 */

function pmV2Moeda_(valor) {

  const numero =
    Math.abs(
      Number(valor) || 0
    );

  return (
    'R$ ' +
    numero
      .toFixed(2)
      .replace('.', ',')
      .replace(
        /\B(?=(\d{3})+(?!\d))/g,
        '.'
      )
  );
}


/**
 * ============================================================
 * 📊 PERCENTUAL
 * ============================================================
 */

function pmV2Percentual_(valor) {

  return (
    (Number(valor || 0) * 100)
      .toFixed(1)
      .replace('.', ',') +
    '%'
  );
}


/**
 * ============================================================
 * 🔠 COLUNA → NÚMERO
 * ============================================================
 */

function pmV2ColunaParaNumero_(coluna) {

  let numero = 0;

  for (
    let i = 0;
    i < coluna.length;
    i++
  ) {

    numero =
      numero * 26 +
      coluna.charCodeAt(i) -
      64;

  }

  return numero;
}