/**
 * ============================================================
 * 💰 DASHBOARD VISUAL V5.3 — JULIUS FINANCE SYSTEM
 * ============================================================
 * - B7 continua sendo a competência oficial.
 * - Dropdown de competência.
 * - KPIs: receitas, despesas, saldo e economia.
 * - Fluxo financeiro acumulado.
 * - Gastos por TODAS as categorias.
 * - Exclui responsáveis do tipo "Terceiro / Reembolso".
 * - Mantém Julius Diz e a caricatura existente.
 * - Não cria onEdit.
 */

function montarDashboardVisual() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    ss.getSheetByName('📊 Dashboard');

  if (!aba) {
    throw new Error(
      'A aba 📊 Dashboard não foi encontrada.'
    );
  }


  const C = {

    fundo: '#F1E4CF',
    fundoClaro: '#F8F0E3',
    card: '#FFF8EC',
    branco: '#FFF9F0',

    azul: '#286173',
    azulEscuro: '#194754',
    azulMuitoEscuro: '#103945',
    azulClaro: '#D8E8EB',

    caramelo: '#C78342',
    carameloClaro: '#E7B678',

    verde: '#668348',
    verdeEscuro: '#4F7A3B',
    verdeClaro: '#E1E8D2',

    vermelho: '#B9553D',
    vermelhoClaro: '#F2D5CC',

    marrom: '#654229',
    marromEscuro: '#3C2A1E',

    texto: '#33291F',
    textoSecundario: '#786554',

    borda: '#C3A989'
  };


  /**
   * ==========================================================
   * COMPETÊNCIA
   * ==========================================================
   */

  let competencia =
    aba.getRange('B7').getValue();


  if (
    !(competencia instanceof Date) ||
    isNaN(
      competencia.getTime()
    )
  ) {

    const hoje =
      new Date();


    competencia =
      new Date(
        hoje.getFullYear(),
        hoje.getMonth(),
        1,
        12,
        0,
        0
      );

  }


  /**
   * ==========================================================
   * REMOVE GRÁFICOS ANTIGOS
   * ==========================================================
   */

  aba
    .getCharts()
    .forEach(
      function (grafico) {

        aba.removeChart(
          grafico
        );

      }
    );


  /**
   * ==========================================================
   * GARANTE QUE BASE DOS GRÁFICOS FIQUE VISÍVEL
   * ==========================================================
   */

  try {

    aba.showColumns(
      27,
      6
    );

  } catch (erro) {

  }


  /**
   * ==========================================================
   * LIMPEZA
   * ==========================================================
   */

  dashboardV53DesfazerMesclas_(
    aba,
    'A1:S46'
  );


  aba
    .getRange('A1:S46')
    .clearContent()
    .clearFormat()
    .setBackground(
      C.fundo
    )
    .setFontFamily(
      'Arial'
    )
    .setFontColor(
      C.texto
    )
    .setVerticalAlignment(
      'middle'
    );


  aba
    .getRange('AH1:AM50')
    .clearContent()
    .clearFormat();


  aba.setHiddenGridlines(
    true
  );


  /**
   * ==========================================================
   * COLUNAS
   * ==========================================================
   */

  aba.setColumnWidth(
    1,
    18
  );


  for (
    let coluna = 2;
    coluna <= 19;
    coluna++
  ) {

    aba.setColumnWidth(
      coluna,
      76
    );

  }


  /**
   * ==========================================================
   * LINHAS
   * ==========================================================
   */

  const alturas = {

    1: 8,

    2: 34,
    3: 27,
    4: 26,

    5: 10,

    6: 24,
    7: 34,
    8: 28,

    9: 10,

    10: 25,
    11: 35,
    12: 38,
    13: 22,

    14: 10,

    15: 28,
    16: 24,

    17: 31,
    18: 31,
    19: 31,
    20: 31,
    21: 31,
    22: 31,
    23: 22,

    24: 10,

    25: 28,
    26: 24,

    27: 31,
    28: 31,
    29: 31,
    30: 31,
    31: 31,
    32: 31,
    33: 31,
    34: 22,

    35: 10,

    36: 28,
    37: 29,
    38: 29,
    39: 29,
    40: 22,

    41: 10,

    42: 22
  };


  Object
    .keys(
      alturas
    )
    .forEach(
      function (linha) {

        aba.setRowHeight(
          Number(
            linha
          ),
          alturas[linha]
        );

      }
    );


  /**
   * ==========================================================
   * CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A1:S4')
    .setBackground(
      C.azul
    );


  aba
    .getRange('B2:H2')
    .merge()
    .setValue(
      'JULIUS FINANCE SYSTEM'
    )
    .setFontSize(
      21
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.branco
    )
    .setHorizontalAlignment(
      'left'
    );


  aba
    .getRange('B3:H3')
    .merge()
    .setValue(
      'Controle de hoje. Liberdade de amanhã.'
    )
    .setFontSize(
      9
    )
    .setFontStyle(
      'italic'
    )
    .setFontColor(
      C.carameloClaro
    );


  aba
    .getRange('B4:H4')
    .merge()
    .setValue(
      'Seu dinheiro. Suas decisões. Seu controle.'
    )
    .setFontSize(
      8
    )
    .setFontColor(
      C.branco
    );


  aba
    .getRange('I2:K2')
    .merge()
    .setValue(
      '📅 COMPETÊNCIA'
    )
    .setFontSize(
      8
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.carameloClaro
    )
    .setHorizontalAlignment(
      'center'
    );


  aba
    .getRange('I3:K4')
    .merge();


  aba
    .getRange('I3')
    .setFormula(
      '=$B$7'
    )
    .setNumberFormat(
      'mmmm/yyyy'
    )
    .setFontSize(
      13
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.branco
    )
    .setHorizontalAlignment(
      'center'
    );


  aba
    .getRange('P2:S2')
    .merge()
    .setValue(
      'CADA REAL TEM ENDEREÇO'
    )
    .setFontSize(
      8
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.carameloClaro
    )
    .setHorizontalAlignment(
      'right'
    );


  aba
    .getRange('P3:S4')
    .merge()
    .setValue(
      '👀 modo economia ativado'
    )
    .setFontSize(
      8
    )
    .setFontStyle(
      'italic'
    )
    .setFontColor(
      C.branco
    )
    .setHorizontalAlignment(
      'right'
    );


  /**
   * ==========================================================
   * SELETOR DE COMPETÊNCIA
   * ==========================================================
   */

  dashboardV53Card_(
    aba,
    'B6:E8',
    C.card,
    C.borda
  );


  aba
    .getRange('B6:E6')
    .merge()
    .setValue(
      '📅 MÊS DE REFERÊNCIA'
    )
    .setFontSize(
      8
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.textoSecundario
    )
    .setHorizontalAlignment(
      'center'
    );


  aba
    .getRange('B7:E8')
    .merge();


  aba
    .getRange('B7')
    .setValue(
      competencia
    )
    .setNumberFormat(
      'mmmm/yyyy'
    )
    .setFontSize(
      17
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.azulEscuro
    )
    .setBackground(
      C.fundoClaro
    )
    .setHorizontalAlignment(
      'center'
    );


  const listaMeses = [];


  for (
    let i = -18;
    i <= 18;
    i++
  ) {

    listaMeses.push([

      new Date(
        competencia.getFullYear(),
        competencia.getMonth() + i,
        1,
        12,
        0,
        0
      )

    ]);

  }


  aba
    .getRange(
      2,
      24,
      listaMeses.length,
      1
    )
    .clearContent()
    .setValues(
      listaMeses
    )
    .setNumberFormat(
      'mmmm/yyyy'
    );


  const regraCompetencia =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInRange(
        aba.getRange(
          2,
          24,
          listaMeses.length,
          1
        ),
        true
      )
      .setAllowInvalid(
        false
      )
      .setHelpText(
        'Selecione a competência desejada.'
      )
      .build();


  aba
    .getRange('B7')
    .setDataValidation(
      regraCompetencia
    );


  /**
   * ==========================================================
   * JULIUS LEMBRA
   * ==========================================================
   */

  dashboardV53Card_(
    aba,
    'F6:S8',
    C.card,
    C.borda
  );


  aba
    .getRange('F6:S6')
    .merge()
    .setValue(
      'JULIUS LEMBRA'
    )
    .setFontSize(
      8
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.caramelo
    );


  aba
    .getRange('F7:S8')
    .merge()
    .setValue(
      '“Se eu não comprar nada, o desconto é maior.”'
    )
    .setFontSize(
      11
    )
    .setFontStyle(
      'italic'
    )
    .setFontColor(
      C.marrom
    )
    .setHorizontalAlignment(
      'center'
    );


  SpreadsheetApp.flush();


  /**
   * ==========================================================
   * COLETA DADOS
   * ==========================================================
   */

  const dados =
    dashboardV53ColetarDados_(
      ss,
      competencia
    );


  /**
   * ==========================================================
   * ENGINE
   * ==========================================================
   */

  aba
    .getRange('T1:AF100')
    .clearContent();


  aba
    .getRange('V1')
    .setValue(
      'JULIUS DASHBOARD ENGINE'
    );


  aba
    .getRange('U2:U8')
    .setValues([

      [
        dados.receitasTotal
      ],

      [
        dados.fixasTotal
      ],

      [
        dados.variaveisTotal
      ],

      [
        dados.parcelamentosTotal
      ],

      [
        dados.despesasTotal
      ],

      [
        dados.saldo
      ],

      [
        dados.economia
      ]

    ]);


  aba
    .getRange('U2:U7')
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  aba
    .getRange('U8')
    .setNumberFormat(
      '0.0%'
    );


  /**
   * ==========================================================
   * KPIs
   * ==========================================================
   */

  aba
    .getRange('B10:S10')
    .merge()
    .setValue(
      'VISÃO GERAL DO MÊS'
    )
    .setFontSize(
      11
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.marromEscuro
    );


  dashboardV53KpiMoeda_(
    aba,
    'B11:E13',
    '↗  RECEITAS',
    dados.receitasTotal,
    C.verdeClaro,
    C.verdeEscuro,
    'entradas da competência',
    C
  );


  dashboardV53KpiMoeda_(
    aba,
    'F11:I13',
    '↘  DESPESAS',
    dados.despesasTotal,
    C.vermelhoClaro,
    C.vermelho,
    'fixas + variáveis + parcelas',
    C
  );


  dashboardV53KpiMoeda_(
    aba,
    'J11:N13',
    '▣  SALDO DO MÊS',
    dados.saldo,
    C.azulClaro,
    C.azulEscuro,
    'resultado da competência',
    C
  );


  dashboardV53KpiPercentual_(
    aba,
    'O11:S13',
    '◎  ECONOMIA',
    dados.economia,
    C.carameloClaro,
    C.marromEscuro,
    'percentual da receita preservado',
    C
  );


  /**
   * ==========================================================
   * FLUXO
   * ==========================================================
   */

  dashboardV53Modulo_(
    aba,
    'B15:I23',
    '▥  FLUXO FINANCEIRO',
    'Receitas x despesas acumuladas na competência',
    C.azul,
    C
  );


  /**
   * ==========================================================
   * CATEGORIAS
   * ==========================================================
   */

  dashboardV53Modulo_(
    aba,
    'J15:N23',
    '◔  GASTOS POR CATEGORIA',
    'Todas as categorias com movimentação no mês',
    C.caramelo,
    C
  );


  /**
   * ==========================================================
   * JULIUS DIZ
   * ==========================================================
   */

  dashboardV53Card_(
    aba,
    'O15:S23',
    C.azulMuitoEscuro,
    C.azulEscuro
  );


  aba
    .getRange('O15:S15')
    .merge()
    .setValue(
      '👀  JULIUS DIZ...'
    )
    .setFontSize(
      11
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.branco
    );


  dashboardV53MontarJuliusDiz_(
    aba,
    dados,
    C
  );


  /**
   * ==========================================================
   * CARTÕES
   * ==========================================================
   */

  dashboardV53Modulo_(
    aba,
    'B25:I34',
    '▰  CARTÕES DE CRÉDITO',
    'Participação dos cartões nos gastos do mês',
    C.azul,
    C
  );


  aba
    .getRange('C28:H32')
    .merge()
    .setValue(
      'RANKING + GRÁFICO DE CARTÕES\nNA PRÓXIMA ETAPA'
    )
    .setFontSize(
      9
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.textoSecundario
    )
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    )
    .setWrap(
      true
    );


  /**
   * ==========================================================
   * ORÇAMENTO
   * ==========================================================
   */

  dashboardV53Modulo_(
    aba,
    'J25:N34',
    '◕  ORÇAMENTO DO MÊS',
    'Planejado x realizado',
    C.verde,
    C
  );


  aba
    .getRange('K28:M32')
    .merge()
    .setValue(
      'BARRAS DE PROGRESSO\nNA PRÓXIMA ETAPA'
    )
    .setFontSize(
      9
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.textoSecundario
    )
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    )
    .setWrap(
      true
    );


  /**
   * ==========================================================
   * REEMBOLSOS
   * ==========================================================
   */

  dashboardV53Modulo_(
    aba,
    'O25:S34',
    '↔  REEMBOLSOS',
    'Recebido • pendente • saldo',
    C.verde,
    C
  );


  aba
    .getRange('P28:R32')
    .merge()
    .setValue(
      'INDICADORES DE\nREEMBOLSO'
    )
    .setFontSize(
      9
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.textoSecundario
    )
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    );


  /**
   * ==========================================================
   * PRÓXIMOS COMPROMISSOS
   * ==========================================================
   */

  dashboardV53Modulo_(
    aba,
    'B36:S40',
    '▣  PRÓXIMOS COMPROMISSOS',
    'Parcelamentos e obrigações que merecem atenção',
    C.marrom,
    C
  );


  aba
    .getRange('C38:R39')
    .merge()
    .setValue(
      'PRÓXIMOS VENCIMENTOS ENTRAM NA ETAPA FINAL'
    )
    .setFontSize(
      9
    )
    .setFontColor(
      C.textoSecundario
    )
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    );


  aba
    .getRange('B42:S42')
    .merge()
    .setValue(
      'JULIUS FINANCE SYSTEM  •  Controle de hoje. Liberdade de amanhã.'
    )
    .setFontSize(
      8
    )
    .setFontStyle(
      'italic'
    )
    .setFontColor(
      C.textoSecundario
    )
    .setHorizontalAlignment(
      'center'
    );


  /**
   * ==========================================================
   * BASES DOS GRÁFICOS
   * ==========================================================
   */

  dashboardV53EscreverFluxo_(
    aba,
    dados
  );


  dashboardV53EscreverCategorias_(
    aba,
    dados
  );


  SpreadsheetApp.flush();


  /**
   * ==========================================================
   * GRÁFICOS
   * ==========================================================
   */

  dashboardV53CriarGraficos_(
    aba,
    dados,
    C
  );


  /**
   * ==========================================================
   * JULIUS
   * ==========================================================
   */

  dashboardV53PosicionarJulius_(
    aba
  );


  /**
   * ==========================================================
   * OCULTA SOMENTE T:Z
   * ==========================================================
   */

  try {

    aba.hideColumns(
      20,
      7
    );

  } catch (erro) {

  }


  SpreadsheetApp.flush();


  ss.toast(
    'Dashboard Julius V5.3 aplicado 📊👀',
    'Julius Finance',
    4
  );

}


/**
 * ============================================================
 * COLETA DOS DADOS
 * ============================================================
 */

function dashboardV53ColetarDados_(
  ss,
  competencia
) {

  const ano =
    competencia.getFullYear();


  const mes =
    competencia.getMonth();


  const diasMes =
    new Date(
      ano,
      mes + 1,
      0
    ).getDate();


  const receitasDia =
    Array(
      diasMes
    ).fill(
      0
    );


  const despesasDia =
    Array(
      diasMes
    ).fill(
      0
    );


  const categorias =
    new Map();


  const terceiros =
    dashboardV53ObterTerceiros_(
      ss
    );


  let receitasTotal =
    0;


  let fixasTotal =
    0;


  let variaveisTotal =
    0;


  let parcelamentosTotal =
    0;


  /**
   * RECEITAS
   */

  const receitasAba =
    ss.getSheetByName(
      '💰 Receitas'
    );


  if (
    receitasAba &&
    receitasAba.getLastRow() >= 4
  ) {

    const linhas =
      receitasAba
        .getRange(
          4,
          1,
          receitasAba.getLastRow() - 3,
          4
        )
        .getValues();


    linhas.forEach(
      function (linha) {

        const data =
          dashboardV53Data_(
            linha[0]
          );


        const valor =
          dashboardV53Numero_(
            linha[3]
          );


        if (
          !data ||
          !valor ||
          data.getFullYear() !== ano ||
          data.getMonth() !== mes
        ) {

          return;

        }


        receitasTotal +=
          valor;


        receitasDia[
          data.getDate() - 1
        ] +=
          valor;

      }
    );

  }


  /**
   * FIXAS
   */

  const fixasAba =
    ss.getSheetByName(
      '📋 Desp.Fixas'
    );


  if (
    fixasAba &&
    fixasAba.getLastRow() >= 6
  ) {

    const linhas =
      fixasAba
        .getRange(
          6,
          1,
          fixasAba.getLastRow() - 5,
          6
        )
        .getValues();


    linhas.forEach(
      function (linha) {

        const categoria =
          dashboardV53Texto_(
            linha[0]
          ) ||
          'Sem categoria';


        const responsavel =
          dashboardV53Texto_(
            linha[1]
          );


        const valor =
          dashboardV53Numero_(
            linha[4]
          );


        if (
          !valor ||
          dashboardV53EhTerceiro_(
            responsavel,
            terceiros
          )
        ) {

          return;

        }


        fixasTotal +=
          valor;


        dashboardV53SomarCategoria_(
          categorias,
          categoria,
          valor
        );


        let dia =
          parseInt(
            linha[3],
            10
          );


        if (
          !Number.isFinite(
            dia
          ) ||
          dia < 1 ||
          dia > diasMes
        ) {

          dia =
            diasMes;

        }


        despesasDia[
          dia - 1
        ] +=
          valor;

      }
    );

  }


  /**
   * VARIÁVEIS
   */

  const variaveisAba =
    ss.getSheetByName(
      '🛒 Desp.Variáveis'
    );


  if (
    variaveisAba &&
    variaveisAba.getLastRow() >= 7
  ) {

    const linhas =
      variaveisAba
        .getRange(
          7,
          1,
          variaveisAba.getLastRow() - 6,
          7
        )
        .getValues();


    linhas.forEach(
      function (linha) {

        const data =
          dashboardV53Data_(
            linha[0]
          );


        const categoria =
          dashboardV53Texto_(
            linha[1]
          ) ||
          'Sem categoria';


        const valor =
          dashboardV53Numero_(
            linha[3]
          );


        const responsavel =
          dashboardV53Texto_(
            linha[6]
          );


        if (
          !data ||
          !valor ||
          data.getFullYear() !== ano ||
          data.getMonth() !== mes
        ) {

          return;

        }


        if (
          dashboardV53EhTerceiro_(
            responsavel,
            terceiros
          )
        ) {

          return;

        }


        variaveisTotal +=
          valor;


        dashboardV53SomarCategoria_(
          categorias,
          categoria,
          valor
        );


        despesasDia[
          data.getDate() - 1
        ] +=
          valor;

      }
    );

  }


  /**
   * PARCELAMENTOS
   */

  const parcelasAba =
    ss.getSheetByName(
      '💳 Parcelamentos'
    );


  if (
    parcelasAba &&
    parcelasAba.getLastRow() >= 4
  ) {

    const linhas =
      parcelasAba
        .getRange(
          4,
          1,
          parcelasAba.getLastRow() - 3,
          16
        )
        .getValues();


    linhas.forEach(
      function (linha) {

        const categoria =
          dashboardV53Texto_(
            linha[0]
          ) ||
          'Sem categoria';


        const valor =
          dashboardV53Numero_(
            linha[6]
          );


        const responsavel =
          dashboardV53Texto_(
            linha[8]
          );


        const dataPrimeira =
          dashboardV53Data_(
            linha[11]
          );


        const vigente =
          dashboardV53Texto_(
            linha[13]
          ).toLowerCase();


        if (
          vigente !== 'sim' ||
          !valor ||
          dashboardV53EhTerceiro_(
            responsavel,
            terceiros
          )
        ) {

          return;

        }


        parcelamentosTotal +=
          valor;


        dashboardV53SomarCategoria_(
          categorias,
          categoria,
          valor
        );


        const dia =
          dataPrimeira
            ? Math.min(
                dataPrimeira.getDate(),
                diasMes
              )
            : diasMes;


        despesasDia[
          dia - 1
        ] +=
          valor;

      }
    );

  }


  /**
   * FLUXO ACUMULADO
   */

  const fluxo =
    [];


  let receitasAcumuladas =
    0;


  let despesasAcumuladas =
    0;


  for (
    let dia = 1;
    dia <= diasMes;
    dia++
  ) {

    receitasAcumuladas +=
      receitasDia[
        dia - 1
      ];


    despesasAcumuladas +=
      despesasDia[
        dia - 1
      ];


    fluxo.push([

      String(
        dia
      ).padStart(
        2,
        '0'
      ) +
      '/' +
      String(
        mes + 1
      ).padStart(
        2,
        '0'
      ),

      receitasAcumuladas,

      despesasAcumuladas

    ]);

  }


  /**
   * TODAS AS CATEGORIAS
   */

  const categoriasOrdenadas =
    Array
      .from(
        categorias.entries()
      )
      .filter(
        function (item) {

          return (
            Number(
              item[1]
            ) > 0
          );

        }
      )
      .sort(
        function (a, b) {

          return (
            b[1] -
            a[1]
          );

        }
      );


  const despesasTotal =
    fixasTotal +
    variaveisTotal +
    parcelamentosTotal;


  const saldo =
    receitasTotal -
    despesasTotal;


  const economia =
    receitasTotal > 0
      ? saldo /
        receitasTotal
      : 0;


  return {

    ano:
      ano,

    mes:
      mes,

    diasMes:
      diasMes,

    receitasTotal:
      receitasTotal,

    fixasTotal:
      fixasTotal,

    variaveisTotal:
      variaveisTotal,

    parcelamentosTotal:
      parcelamentosTotal,

    despesasTotal:
      despesasTotal,

    saldo:
      saldo,

    economia:
      economia,

    fluxo:
      fluxo,

    categorias:
      categoriasOrdenadas

  };

}


/**
 * ============================================================
 * TERCEIROS
 * ============================================================
 */

function dashboardV53ObterTerceiros_(
  ss
) {

  const resultado =
    new Set();


  const aba =
    ss.getSheetByName(
      '_CONFIG'
    );


  if (
    !aba ||
    aba.getLastRow() < 2
  ) {

    return resultado;

  }


  aba
    .getRange(
      2,
      1,
      aba.getLastRow() - 1,
      2
    )
    .getDisplayValues()
    .forEach(
      function (linha) {

        const nome =
          dashboardV53Texto_(
            linha[0]
          );


        const tipo =
          dashboardV53Texto_(
            linha[1]
          ).toLowerCase();


        if (
          nome &&
          tipo ===
            'terceiro / reembolso'
        ) {

          resultado.add(
            nome.toLowerCase()
          );

        }

      }
    );


  return resultado;

}


/**
 * ============================================================
 * RESPONSÁVEL É TERCEIRO?
 * ============================================================
 */

function dashboardV53EhTerceiro_(
  responsavel,
  terceiros
) {

  const texto =
    dashboardV53Texto_(
      responsavel
    ).toLowerCase();


  return (
    texto
      ? terceiros.has(
          texto
        )
      : false
  );

}


/**
 * ============================================================
 * SOMA CATEGORIA
 * ============================================================
 */

function dashboardV53SomarCategoria_(
  mapa,
  categoria,
  valor
) {

  const nome =
    dashboardV53Texto_(
      categoria
    ) ||
    'Sem categoria';


  mapa.set(
    nome,
    (
      mapa.get(
        nome
      ) || 0
    ) +
    Number(
      valor || 0
    )
  );

}


/**
 * ============================================================
 * BASE FLUXO
 * ============================================================
 */

function dashboardV53EscreverFluxo_(
  aba,
  dados
) {

  aba
    .getRange('AA1:AC40')
    .clearContent();


  aba
    .getRange('AA1:AC1')
    .setValues([[
      'Data',
      'Receitas',
      'Despesas'
    ]]);


  if (
    !dados.fluxo.length
  ) {

    return;

  }


  aba
    .getRange(
      2,
      27,
      dados.fluxo.length,
      3
    )
    .setValues(
      dados.fluxo
    );


  aba
    .getRange(
      2,
      28,
      dados.fluxo.length,
      2
    )
    .setNumberFormat(
      '"R$ " #,##0.00'
    );

}


/**
 * ============================================================
 * BASE CATEGORIAS
 * ============================================================
 */

function dashboardV53EscreverCategorias_(
  aba,
  dados
) {

  aba
    .getRange('AE1:AF100')
    .clearContent();


  aba
    .getRange('AE1:AF1')
    .setValues([[
      'Categoria',
      'Valor'
    ]]);


  if (
    !dados.categorias.length
  ) {

    return;

  }


  aba
    .getRange(
      2,
      31,
      dados.categorias.length,
      2
    )
    .setValues(
      dados.categorias
    );


  aba
    .getRange(
      2,
      32,
      dados.categorias.length,
      1
    )
    .setNumberFormat(
      '"R$ " #,##0.00'
    );

}


/**
 * ============================================================
 * GRÁFICOS
 * ============================================================
 */

function dashboardV53CriarGraficos_(
  aba,
  dados,
  C
) {

  /**
   * FLUXO
   */

  if (
    dados.fluxo.length
  ) {

    const graficoFluxo =
      aba
        .newChart()
        .setChartType(
          Charts.ChartType.LINE
        )
        .addRange(
          aba.getRange(
            'AA1:AC' +
            (
              dados.fluxo.length +
              1
            )
          )
        )
        .setNumHeaders(
          1
        )
        .setPosition(
          17,
          2,
          0,
          0
        )
        .setOption(
          'width',
          590
        )
        .setOption(
          'height',
          245
        )
        .build();


    aba.insertChart(
      graficoFluxo
    );

  }


  /**
   * CATEGORIAS
   */

  if (
    dados.categorias.length
  ) {

    const graficoCategorias =
      aba
        .newChart()
        .setChartType(
          Charts.ChartType.PIE
        )
        .addRange(
          aba.getRange(
            'AE1:AF' +
            (
              dados.categorias.length +
              1
            )
          )
        )
        .setNumHeaders(
          1
        )
        .setPosition(
          17,
          10,
          0,
          0
        )
        .setOption(
          'width',
          380
        )
        .setOption(
          'height',
          245
        )
        .setOption(
          'pieHole',
          0.55
        )
        .build();


    aba.insertChart(
      graficoCategorias
    );

  }

}


/**
 * ============================================================
 * JULIUS DIZ
 * ============================================================
 */

function dashboardV53MontarJuliusDiz_(
  aba,
  dados,
  C
) {

  const receitas =
    dados.receitasTotal;


  const despesas =
    dados.despesasTotal;


  const saldo =
    dados.saldo;


  const economia =
    dados.economia;


  const parcelas =
    dados.parcelamentosTotal;


  let alerta1;


  if (
    receitas > 0 &&
    despesas > receitas
  ) {

    alerta1 =
      '🔴 ATENÇÃO\n' +
      'As despesas ultrapassaram a receita em ' +
      dashboardV53Moeda_(
        despesas -
        receitas
      ) +
      '.';

  } else if (
    saldo > 0
  ) {

    alerta1 =
      '🟢 MÊS POSITIVO\n' +
      'Você está fechando com ' +
      dashboardV53Moeda_(
        saldo
      ) +
      ' de saldo.';

  } else {

    alerta1 =
      '🟠 ATENÇÃO\n' +
      'Acompanhe o fechamento da competência.';

  }


  let alerta2;


  if (
    economia >= 0.20
  ) {

    alerta2 =
      '💡 BOA ECONOMIA\n' +
      dashboardV53Percentual_(
        economia
      ) +
      ' da receita está sendo preservada.';

  } else if (
    economia >= 0
  ) {

    alerta2 =
      '⚠ OLHO NA ECONOMIA\n' +
      'A reserva do mês está em ' +
      dashboardV53Percentual_(
        economia
      ) +
      '.';

  } else {

    alerta2 =
      '⚠ GASTOS ACIMA DA RENDA\n' +
      'O mês está consumindo mais do que arrecada.';

  }


  const pesoParcelas =
    receitas > 0
      ? parcelas /
        receitas
      : 0;


  const alerta3 =
    pesoParcelas >= 0.40
      ? '💳 CRÉDITO COMPROMETIDO\n' +
        'Parcelamentos consomem ' +
        dashboardV53Percentual_(
          pesoParcelas
        ) +
        ' da receita.'
      : '💳 PARCELAMENTOS\n' +
        'Compromisso atual: ' +
        dashboardV53Moeda_(
          parcelas
        ) +
        '.';


  aba
    .getRange('O16:S23')
    .clearContent();


  dashboardV53Insight_(
    aba,
    'O16:S17',
    alerta1,
    C.card,
    saldo < 0
      ? C.vermelho
      : C.verdeEscuro
  );


  dashboardV53Insight_(
    aba,
    'O18:S19',
    alerta2,
    C.card,
    economia < 0
      ? C.vermelho
      : C.caramelo
  );


  dashboardV53Insight_(
    aba,
    'O20:S21',
    alerta3,
    C.card,
    C.azulEscuro
  );


  aba
    .getRange('O22:S23')
    .merge()
    .setValue(
      '“Disciplina é o que faz o mês fechar no azul.”'
    )
    .setFontSize(
      7
    )
    .setFontStyle(
      'italic'
    )
    .setFontColor(
      C.carameloClaro
    )
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    )
    .setWrap(
      true
    );

}


/**
 * ============================================================
 * INSIGHT
 * ============================================================
 */

function dashboardV53Insight_(
  aba,
  intervalo,
  texto,
  fundo,
  cor
) {

  aba
    .getRange(
      intervalo
    )
    .merge()
    .setValue(
      texto
    )
    .setBackground(
      fundo
    )
    .setFontSize(
      8
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      cor
    )
    .setVerticalAlignment(
      'middle'
    )
    .setWrap(
      true
    );

}


/**
 * ============================================================
 * KPI MOEDA
 * ============================================================
 */

function dashboardV53KpiMoeda_(
  aba,
  intervalo,
  titulo,
  valor,
  fundo,
  cor,
  subtitulo,
  C
) {

  const range =
    aba.getRange(
      intervalo
    );


  const linha =
    range.getRow();


  const coluna =
    range.getColumn();


  const colunas =
    range.getNumColumns();


  dashboardV53Card_(
    aba,
    intervalo,
    fundo,
    C.borda
  );


  aba
    .getRange(
      linha,
      coluna,
      1,
      colunas
    )
    .merge()
    .setValue(
      titulo
    )
    .setFontSize(
      9
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      cor
    )
    .setHorizontalAlignment(
      'center'
    );


  aba
    .getRange(
      linha + 1,
      coluna,
      1,
      colunas
    )
    .merge()
    .setValue(
      valor
    )
    .setFontSize(
      18
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.marromEscuro
    )
    .setHorizontalAlignment(
      'center'
    )
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  aba
    .getRange(
      linha + 2,
      coluna,
      1,
      colunas
    )
    .merge()
    .setValue(
      subtitulo
    )
    .setFontSize(
      7
    )
    .setFontStyle(
      'italic'
    )
    .setFontColor(
      C.textoSecundario
    )
    .setHorizontalAlignment(
      'center'
    );

}


/**
 * ============================================================
 * KPI PERCENTUAL
 * ============================================================
 */

function dashboardV53KpiPercentual_(
  aba,
  intervalo,
  titulo,
  valor,
  fundo,
  cor,
  subtitulo,
  C
) {

  const range =
    aba.getRange(
      intervalo
    );


  const linha =
    range.getRow();


  const coluna =
    range.getColumn();


  const colunas =
    range.getNumColumns();


  dashboardV53Card_(
    aba,
    intervalo,
    fundo,
    C.borda
  );


  aba
    .getRange(
      linha,
      coluna,
      1,
      colunas
    )
    .merge()
    .setValue(
      titulo
    )
    .setFontSize(
      9
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      cor
    )
    .setHorizontalAlignment(
      'center'
    );


  aba
    .getRange(
      linha + 1,
      coluna,
      1,
      colunas
    )
    .merge()
    .setValue(
      valor
    )
    .setFontSize(
      18
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      C.marromEscuro
    )
    .setHorizontalAlignment(
      'center'
    )
    .setNumberFormat(
      '0.0%'
    );


  aba
    .getRange(
      linha + 2,
      coluna,
      1,
      colunas
    )
    .merge()
    .setValue(
      subtitulo
    )
    .setFontSize(
      7
    )
    .setFontStyle(
      'italic'
    )
    .setFontColor(
      C.textoSecundario
    )
    .setHorizontalAlignment(
      'center'
    );

}


/**
 * ============================================================
 * MÓDULO
 * ============================================================
 */

function dashboardV53Modulo_(
  aba,
  intervalo,
  titulo,
  subtitulo,
  corTitulo,
  C
) {

  const range =
    aba.getRange(
      intervalo
    );


  const linha =
    range.getRow();


  const coluna =
    range.getColumn();


  const colunas =
    range.getNumColumns();


  dashboardV53Card_(
    aba,
    intervalo,
    C.card,
    C.borda
  );


  aba
    .getRange(
      linha,
      coluna,
      1,
      colunas
    )
    .merge()
    .setValue(
      titulo
    )
    .setFontSize(
      10
    )
    .setFontWeight(
      'bold'
    )
    .setFontColor(
      corTitulo
    );


  aba
    .getRange(
      linha + 1,
      coluna,
      1,
      colunas
    )
    .merge()
    .setValue(
      subtitulo
    )
    .setFontSize(
      7
    )
    .setFontColor(
      C.textoSecundario
    );

}


/**
 * ============================================================
 * CARD
 * ============================================================
 */

function dashboardV53Card_(
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


  range.setBorder(
    true,
    true,
    true,
    true,
    false,
    false,
    borda,
    SpreadsheetApp
      .BorderStyle
      .SOLID
  );

}


/**
 * ============================================================
 * POSICIONA JULIUS
 * ============================================================
 */

function dashboardV53PosicionarJulius_(
  aba
) {

  try {

    const imagens =
      aba.getImages();


    if (
      !imagens ||
      !imagens.length
    ) {

      return;

    }


    let julius =
      null;


    imagens.forEach(
      function (imagem) {

        if (julius) {

          return;

        }


        let titulo =
          '';


        let descricao =
          '';


        try {

          titulo =
            String(
              imagem.getAltTextTitle() ||
              ''
            ).toLowerCase();


          descricao =
            String(
              imagem.getAltTextDescription() ||
              ''
            ).toLowerCase();

        } catch (erro) {

        }


        if (
          titulo.includes(
            'julius'
          ) ||
          descricao.includes(
            'julius'
          )
        ) {

          julius =
            imagem;

        }

      }
    );


    if (!julius) {

      julius =
        imagens[0];

    }


    julius
      .setAnchorCell(
        aba.getRange(
          'L1'
        )
      )
      .setAnchorCellXOffset(
        8
      )
      .setAnchorCellYOffset(
        1
      )
      .setWidth(
        150
      )
      .setHeight(
        112
      );


  } catch (erro) {

    console.log(
      'Julius não reposicionado: ' +
      erro
    );

  }

}


/**
 * ============================================================
 * DESFAZ MESCLAS
 * ============================================================
 */

function dashboardV53DesfazerMesclas_(
  aba,
  intervalo
) {

  aba
    .getRange(
      intervalo
    )
    .getMergedRanges()
    .forEach(
      function (range) {

        range.breakApart();

      }
    );

}


/**
 * ============================================================
 * CONVERTE NÚMERO
 * ============================================================
 */

function dashboardV53Numero_(
  valor
) {

  if (
    typeof valor ===
    'number'
  ) {

    return (
      Number.isFinite(
        valor
      )
        ? valor
        : 0
    );

  }


  if (
    valor === null ||
    valor === undefined ||
    valor === ''
  ) {

    return 0;

  }


  let texto =
    String(
      valor
    )
      .trim()
      .replace(
        /R\$/gi,
        ''
      )
      .replace(
        /\s/g,
        ''
      );


  if (
    texto.includes(
      ','
    )
  ) {

    texto =
      texto
        .replace(
          /\./g,
          ''
        )
        .replace(
          ',',
          '.'
        );

  }


  const numero =
    Number(
      texto
    );


  return (
    Number.isFinite(
      numero
    )
      ? numero
      : 0
  );

}


/**
 * ============================================================
 * CONVERTE DATA
 * ============================================================
 */

function dashboardV53Data_(
  valor
) {

  if (
    valor instanceof Date &&
    !isNaN(
      valor.getTime()
    )
  ) {

    return valor;

  }


  if (!valor) {

    return null;

  }


  if (
    typeof valor ===
    'number'
  ) {

    const data =
      new Date(
        1899,
        11,
        30,
        12,
        0,
        0
      );


    data.setDate(
      data.getDate() +
      valor
    );


    return (
      isNaN(
        data.getTime()
      )
        ? null
        : data
    );

  }


  const texto =
    String(
      valor
    ).trim();


  const br =
    texto.match(
      /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+.*)?$/
    );


  if (br) {

    const data =
      new Date(
        Number(
          br[3]
        ),
        Number(
          br[2]
        ) - 1,
        Number(
          br[1]
        ),
        12,
        0,
        0
      );


    return (
      isNaN(
        data.getTime()
      )
        ? null
        : data
    );

  }


  const data =
    new Date(
      texto
    );


  return (
    isNaN(
      data.getTime()
    )
      ? null
      : data
  );

}


/**
 * ============================================================
 * TEXTO
 * ============================================================
 */

function dashboardV53Texto_(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  ).trim();

}


/**
 * ============================================================
 * MOEDA
 * ============================================================
 */

function dashboardV53Moeda_(
  valor
) {

  const numero =
    Number(
      valor
    ) || 0;


  const negativo =
    numero < 0;


  const texto =
    Math
      .abs(
        numero
      )
      .toFixed(
        2
      )
      .replace(
        '.',
        ','
      )
      .replace(
        /\B(?=(\d{3})+(?!\d))/g,
        '.'
      );


  return (
    negativo
      ? '-R$ '
      : 'R$ '
  ) +
  texto;

}


/**
 * ============================================================
 * PERCENTUAL
 * ============================================================
 */

function dashboardV53Percentual_(
  valor
) {

  return (
    (
      Number(
        valor
      ) *
      100
    )
      .toFixed(
        1
      )
      .replace(
        '.',
        ','
      ) +
    '%'
  );

}