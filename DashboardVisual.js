/**
 * ============================================================
 * 💰 DASHBOARD VISUAL V4 — JULIUS FINANCE SYSTEM
 * ============================================================
 *
 * Foco desta versão:
 * - remover textos duplicados;
 * - limpar o painel de alertas;
 * - deixar espaço real para o Julius;
 * - preservar fórmulas e lógica existente;
 * - manter a paleta aprovada.
 */

function montarDashboardVisual() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('📊 Dashboard');

  if (!aba) {
    throw new Error('A aba 📊 Dashboard não foi encontrada.');
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
    .getRange('A1:Z50')
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
    A: 22,
    B: 150,
    C: 72,
    D: 72,
    E: 18,
    F: 150,
    G: 72,
    H: 72,
    I: 18,
    J: 150,
    K: 72,
    L: 72,
    M: 22
  };

  Object.keys(larguras).forEach(coluna => {

    aba.setColumnWidth(
      colunaParaNumeroDashboard_(coluna),
      larguras[coluna]
    );

  });


  /**
   * ==========================================================
   * 📏 ALTURAS
   * ==========================================================
   */

  const alturas = {

    1: 10,
    2: 46,
    3: 10,
    4: 27,
    5: 12,

    6: 22,
    7: 42,

    8: 16,
    9: 30,

    10: 28,
    11: 44,
    12: 18,

    13: 10,
    14: 10,

    15: 28,
    16: 44,
    17: 18,

    18: 12,
    19: 12,

    20: 30,
    21: 34,
    22: 34,
    23: 34,

    24: 14,

    25: 30,

    26: 28,
    27: 28,
    28: 28,
    29: 28,
    30: 30,

    31: 12,
    32: 24
  };

  Object.keys(alturas).forEach(linha => {

    aba.setRowHeight(
      Number(linha),
      alturas[linha]
    );

  });


  /**
   * ==========================================================
   * 🔝 CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A1:M3')
    .setBackground(C.azulPetroleo);


  aba
    .getRange('B2')
    .setValue('💰 CONTROLE FINANCEIRO')
    .setFontSize(22)
    .setFontWeight('bold')
    .setFontColor(C.branco)
    .setHorizontalAlignment('left');


  aba
    .getRange('J2')
    .setValue('JULIUS FINANCE')
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(C.carameloClaro)
    .setHorizontalAlignment('right');


  aba
    .getRange('A4:M4')
    .setBackground(C.caramelo)
    .setFontColor(C.marromEscuro);


  aba
    .getRange('B4')
    .setValue('cada real tem endereço 📍')
    .setFontSize(10)
    .setFontStyle('italic')
    .setFontWeight('bold');


  aba
    .getRange('J4')
    .setValue('modo economia ativado 👀')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('right');


  /**
   * ==========================================================
   * 📅 COMPETÊNCIA
   * ==========================================================
   */

  cardBase_(
    aba,
    'B6:D7',
    C.card,
    C.borda
  );


  aba
    .getRange('B6')
    .setValue('📅 COMPETÊNCIA')
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(C.textoSecundario);


  aba
    .getRange('B7')
    .setBackground(C.fundoClaro)
    .setFontColor(C.azulEscuro)
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setNumberFormat('mmmm/yyyy');


  /**
   * ==========================================================
   * 💡 JULIUS LEMBRA
   * ==========================================================
   */

  cardBase_(
    aba,
    'F6:L7',
    C.card,
    C.borda
  );


  aba
    .getRange('F6')
    .setValue('💡 JULIUS LEMBRA')
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontColor(C.caramelo);


  aba
    .getRange('F7')
    .setValue('se não gastar, sobra. A matemática é implacável.')
    .setFontSize(10)
    .setFontStyle('italic')
    .setFontColor(C.textoSecundario);


  /**
   * ==========================================================
   * 📊 VISÃO GERAL
   * ==========================================================
   */

  aba
    .getRange('B9:L9')
    .setBackground(C.fundo);


  aba
    .getRange('B9')
    .setValue('VISÃO GERAL DO MÊS')
    .setFontSize(13)
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro);


  criarCardFinanceiro_(
    aba,
    'B10:D12',
    'B10',
    'B11',
    '💵 RECEITAS',
    C.verde,
    C
  );


  criarCardFinanceiro_(
    aba,
    'F10:H12',
    'F10',
    'F11',
    '🏠 DESPESAS FIXAS',
    C.vermelho,
    C
  );


  criarCardFinanceiro_(
    aba,
    'J10:L12',
    'J10',
    'J11',
    '🛒 DESPESAS VARIÁVEIS',
    C.laranja,
    C
  );


  criarCardFinanceiro_(
    aba,
    'B15:D17',
    'B15',
    'B16',
    '💳 PARCELAMENTOS',
    C.azulPetroleo,
    C
  );


  /**
   * ==========================================================
   * 💰 SALDO — LIMPO
   * ==========================================================
   */

  cardBase_(
    aba,
    'F15:H17',
    C.azulPetroleo,
    C.azulEscuro
  );


  /**
   * limpa textos antigos dentro do card
   * sem tocar no H16
   */

  aba
    .getRange('F15:G17')
    .clearContent();


  aba
    .getRange('F15')
    .setValue('💰 SALDO DO MÊS')
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(C.carameloClaro);


  aba
    .getRange('H16')
    .setFontSize(19)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setNumberFormat('"R$ " #,##0.00;-"R$ " #,##0.00');


  const saldo =
    Number(
      aba.getRange('H16').getValue()
    );


  if (saldo < 0) {

    aba
      .getRange('H16')
      .setFontColor('#FFD5C5');

  } else {

    aba
      .getRange('H16')
      .setFontColor('#E6F0D9');
  }


  /**
   * ==========================================================
   * 👔 ESPAÇO LIVRE PARA O JULIUS
   * ==========================================================
   *
   * Sem borda e sem card.
   * Apenas fundo para a ilustração flutuante.
   */

  aba
    .getRange('J15:L23')
    .clearContent()
    .setBackground(C.fundo);


  /**
   * ==========================================================
   * 👀 JULIUS ESTÁ DE OLHO
   * ==========================================================
   */

  cardBase_(
    aba,
    'B20:H23',
    C.card,
    C.borda
  );


  /**
   * limpa restos antigos
   */

  aba
    .getRange('B20:H23')
    .clearContent();


  aba
    .getRange('B20')
    .setValue('👀 JULIUS ESTÁ DE OLHO')
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontColor(C.marrom);


  /**
   * ==========================================================
   * ALERTA 1 — SALDO
   * ==========================================================
   */

  let textoSaldo;
  let corSaldo;


  if (saldo < 0) {

    textoSaldo =
      '🔴 O mês está fechando no negativo: ' +
      formatarMoedaDashboard_(saldo);

    corSaldo =
      C.vermelho;

  } else {

    textoSaldo =
      '🟢 Saldo positivo: ' +
      formatarMoedaDashboard_(saldo);

    corSaldo =
      C.verde;
  }


  aba
    .getRange('B21:H21')
    .merge()
    .setValue(textoSaldo)
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(corSaldo)
    .setWrap(true);


  /**
   * ==========================================================
   * ALERTA 2 — ECONOMIA
   * ==========================================================
   *
   * IMPORTANTE:
   * D21 tinha a fórmula antiga.
   *
   * Como acabamos de usar B21:H21 como linha visual,
   * precisamos buscar a taxa ANTES de substituir.
   *
   * Para preservar a lógica, calculamos a taxa a partir
   * dos valores do próprio Dashboard.
   */

  const receitas =
    Number(
      aba.getRange('B11').getValue()
    ) || 0;


  const fixas =
    Number(
      aba.getRange('F11').getValue()
    ) || 0;


  const variaveis =
    Number(
      aba.getRange('J11').getValue()
    ) || 0;


  const parcelas =
    Number(
      aba.getRange('B16').getValue()
    ) || 0;


  let taxa =
    0;


  if (receitas > 0) {

    taxa =
      1 -
      (
        fixas +
        variaveis +
        parcelas
      ) /
      receitas;
  }


  let textoTaxa;
  let corTaxa;


  if (taxa >= 0.20) {

    textoTaxa =
      '🟢 Economia aprovada pelo Julius: ' +
      formatarPercentualDashboard_(taxa);

    corTaxa =
      C.verde;

  } else if (taxa >= 0.10) {

    textoTaxa =
      '🟠 Economia em atenção: ' +
      formatarPercentualDashboard_(taxa);

    corTaxa =
      C.laranja;

  } else {

    textoTaxa =
      '🔴 Taxa de economia crítica: ' +
      formatarPercentualDashboard_(taxa);

    corTaxa =
      C.vermelho;
  }


  aba
    .getRange('B22:H22')
    .merge()
    .setValue(textoTaxa)
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(corTaxa)
    .setWrap(true);


  /**
   * ==========================================================
   * ALERTA 3 — PARCELAMENTOS
   * ==========================================================
   */

  let percentualParcelas =
    0;


  if (receitas > 0) {

    percentualParcelas =
      parcelas /
      receitas;
  }


  let textoParcelas;
  let corParcelas;


  if (percentualParcelas >= 0.50) {

    textoParcelas =
      '🟠 Parcelamentos consomem ' +
      formatarPercentualDashboard_(
        percentualParcelas
      ) +
      ' da receita.';

    corParcelas =
      C.laranja;

  } else {

    textoParcelas =
      '🔵 Parcelamentos do mês: ' +
      formatarMoedaDashboard_(
        parcelas
      );

    corParcelas =
      C.azulPetroleo;
  }


  aba
    .getRange('B23:H23')
    .merge()
    .setValue(textoParcelas)
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(corParcelas)
    .setWrap(true);


  /**
   * ==========================================================
   * 📊 FLUXO DE CAIXA
   * ==========================================================
   */

  aba
    .getRange('B25:L25')
    .setBackground(C.marrom)
    .setFontColor(C.branco);


  aba
    .getRange('B25')
    .setValue('📊 PARA ONDE FOI O DINHEIRO?')
    .setFontSize(12)
    .setFontWeight('bold');


  estilizarFluxoDashboard_(
    aba,
    26,
    C.verdeClaro,
    C.verde,
    C
  );


  estilizarFluxoDashboard_(
    aba,
    27,
    C.vermelhoClaro,
    C.vermelho,
    C
  );


  estilizarFluxoDashboard_(
    aba,
    28,
    C.vermelhoClaro,
    C.vermelho,
    C
  );


  estilizarFluxoDashboard_(
    aba,
    29,
    C.laranjaClaro,
    C.laranja,
    C
  );


  estilizarFluxoDashboard_(
    aba,
    30,
    C.azulClaro,
    C.azulEscuro,
    C
  );


  /**
   * ==========================================================
   * 💵 FORMATO MONETÁRIO
   * ==========================================================
   */

  [
    'B11',
    'F11',
    'J11',
    'B16',
    'H16',
    'F26',
    'F27',
    'F28',
    'F29',
    'F30'
  ].forEach(celula => {

    aba
      .getRange(celula)
      .setNumberFormat(
        '"R$ " #,##0.00;-"R$ " #,##0.00'
      );

  });


  /**
   * ==========================================================
   * 🧾 RODAPÉ
   * ==========================================================
   */

  aba
    .getRange('B32:L32')
    .setBackground(C.fundo)
    .setFontColor(C.textoSecundario);


  aba
    .getRange('B32')
    .setValue(
      '🟢 saudável   •   🟠 atenção   •   🔴 Julius quer explicações'
    )
    .setFontSize(8)
    .setFontStyle('italic');


  SpreadsheetApp.flush();


  ss.toast(
    'Dashboard Julius V4 aplicado 👀💰',
    'Controle Financeiro',
    4
  );
}


/**
 * ============================================================
 * CARD FINANCEIRO
 * ============================================================
 */

function criarCardFinanceiro_(
  aba,
  intervalo,
  tituloCelula,
  valorCelula,
  titulo,
  corDestaque,
  C
) {

  cardBase_(
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
    .setFontColor(corDestaque);


  aba
    .getRange(valorCelula)
    .setFontSize(18)
    .setFontWeight('bold')
    .setFontColor(C.marromEscuro)
    .setHorizontalAlignment('center')
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );
}


/**
 * ============================================================
 * CARD BASE
 * ============================================================
 */

function cardBase_(
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


  range.setBorder(
    true,
    true,
    true,
    true,
    false,
    false,
    borda,
    SpreadsheetApp.BorderStyle.SOLID
  );
}


/**
 * ============================================================
 * FLUXO
 * ============================================================
 */

function estilizarFluxoDashboard_(
  aba,
  linha,
  fundo,
  texto,
  C
) {

  const range =
    aba.getRange(
      linha,
      2,
      1,
      11
    );


  range
    .setBackground(fundo)
    .setFontColor(texto);


  range.setBorder(
    true,
    true,
    true,
    true,
    false,
    false,
    C.borda,
    SpreadsheetApp.BorderStyle.SOLID
  );


  aba
    .getRange(linha, 2)
    .setFontWeight(
      linha === 30
        ? 'bold'
        : 'normal'
    );


  aba
    .getRange(linha, 6)
    .setFontWeight('bold');
}


/**
 * ============================================================
 * FORMATAÇÃO MOEDA
 * ============================================================
 */

function formatarMoedaDashboard_(
  valor
) {

  const negativo =
    valor < 0;


  const absoluto =
    Math.abs(
      Number(valor) || 0
    );


  const texto =
    absoluto
      .toFixed(2)
      .replace('.', ',')
      .replace(
        /\B(?=(\d{3})+(?!\d))/g,
        '.'
      );


  return (
    negativo
      ? '-R$ '
      : 'R$ '
  ) + texto;
}


/**
 * ============================================================
 * FORMATAÇÃO PERCENTUAL
 * ============================================================
 */

function formatarPercentualDashboard_(
  valor
) {

  return (
    (Number(valor) * 100)
      .toFixed(1)
      .replace('.', ',') +
    '%'
  );
}


/**
 * ============================================================
 * COLUNA → NÚMERO
 * ============================================================
 */

function colunaParaNumeroDashboard_(
  coluna
) {

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