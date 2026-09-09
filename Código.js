const CONFIG = {
  ABA_PARCELAMENTOS: '💳 Parcelamentos',
  ABA_DASHBOARD: '📊 Dashboard',
  MES_REFERENCIA: 'B7',

  PRIMEIRA_LINHA: 4,
  ULTIMA_COLUNA: 16, // P

  CATEGORIA: 1,        // A
  PRODUTO: 2,          // B
  VALOR_TOTAL: 3,      // C
  QTD_PARCELAS: 4,     // D
  PARCELA_ATUAL: 5,    // E
  RESTANTES: 6,        // F
  VALOR_PARCELA: 7,    // G
  CARTAO: 8,           // H
  RESPONSAVEL: 9,      // I
  OBSERVACAO: 10,      // J
  PAGO: 11,            // K
  DATA_PRIMEIRA: 12,   // L
  FINALIZADO: 13,      // M
  VIGENTE: 14,         // N
  ULTIMA_COMP: 15,     // O
  MESES_DECORRIDOS: 16 // P
};


/* =========================================================
   MENU
========================================================= */

function onOpen() {

  prepararConfiguracoes_();


  SpreadsheetApp
    .getUi()

    .createMenu(
      '💰 Financeiro'
    )

    .addItem(
      '➕ Novo Parcelamento',
      'abrirNovoParcelamento'
    )

    .addItem(
      '📄 Importar Fatura PDF',
      'abrirImportarFatura'
    )

    .addSeparator()

    .addItem(
      '⚙️ Configurações',
      'abrirConfiguracoes'
    )

    .addSeparator()

    .addItem(
      '↕️ Ordenar Parcelamentos',
      'ordenarParcelamentos'
    )

    .addToUi();

}

/* =========================================================
   JANELA
========================================================= */

function abrirNovoParcelamento() {
  const html = HtmlService
    .createHtmlOutputFromFile('NovoParcelamento')
    .setWidth(500)
    .setHeight(650);

  SpreadsheetApp.getUi()
    .showModalDialog(html, '➕ Novo Parcelamento');
}



/* =========================================================
   OPÇÕES DOS DROPDOWNS
========================================================= */

function obterOpcoesNovoParcelamento() {
  const aba = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.ABA_PARCELAMENTOS);

  if (!aba) {
    throw new Error('A aba 💳 Parcelamentos não foi encontrada.');
  }

  return {
    categorias: obterOpcoesValidacao_(aba, 'A4'),
    cartoes: obterOpcoesValidacao_(aba, 'H4'),
    responsaveis: obterOpcoesValidacao_(aba, 'I4')
  };
}


function obterOpcoesValidacao_(aba, celula) {
  const regra = aba.getRange(celula).getDataValidation();

  if (!regra) return [];

  const criterio = regra.getCriteriaType();
  const valores = regra.getCriteriaValues();

  if (
    criterio === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST
  ) {
    return valores[0];
  }

  if (
    criterio === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE
  ) {
    return valores[0]
      .getDisplayValues()
      .flat()
      .filter(valor => valor !== '');
  }

  return [];
}


/* =========================================================
   SALVAR NOVO PARCELAMENTO
========================================================= */

function salvarNovoParcelamento(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG.ABA_PARCELAMENTOS);

  if (!aba) {
    throw new Error('A aba 💳 Parcelamentos não foi encontrada.');
  }

  validarDados_(dados);

  const linhaModelo = encontrarLinhaModelo_(aba);

  if (!linhaModelo) {
    throw new Error(
      'Não encontrei uma linha com as fórmulas automáticas E/F/M/N/O/P.'
    );
  }

  const novaLinha = criarNovaLinha_(aba);

  prepararNovaLinha_(aba, linhaModelo, novaLinha);


  const valorTotal = converterNumero_(dados.valorTotal);
  const qtdParcelas = Number(dados.qtdParcelas);

  const partesData = dados.dataPrimeiraParcela.split('-');

  /*
   Meio-dia evita o problema de 01/08 aparecer como 31/07.
  */
  const dataPrimeiraParcela = new Date(
    Number(partesData[0]),
    Number(partesData[1]) - 1,
    Number(partesData[2]),
    12,
    0,
    0
  );


  /* =======================================================
     DADOS INFORMADOS NA JANELA
  ======================================================= */

  aba.getRange(novaLinha, CONFIG.CATEGORIA)
    .setValue(dados.categoria);

  aba.getRange(novaLinha, CONFIG.PRODUTO)
    .setValue(dados.produto);

  aba.getRange(novaLinha, CONFIG.VALOR_TOTAL)
    .setValue(valorTotal);

  aba.getRange(novaLinha, CONFIG.QTD_PARCELAS)
    .setValue(qtdParcelas);

  aba.getRange(novaLinha, CONFIG.CARTAO)
    .setValue(dados.cartao);

  aba.getRange(novaLinha, CONFIG.RESPONSAVEL)
    .setValue(dados.responsavel);

  aba.getRange(novaLinha, CONFIG.OBSERVACAO)
    .setValue(dados.observacao || '');

  aba.getRange(novaLinha, CONFIG.PAGO)
    .clearContent();

  aba.getRange(novaLinha, CONFIG.DATA_PRIMEIRA)
    .setValue(dataPrimeiraParcela);


  /* =======================================================
     VALOR DA PARCELA
     Não depende mais de fórmula.
  ======================================================= */

  aba.getRange(novaLinha, CONFIG.VALOR_PARCELA)
    .setValue(valorTotal / qtdParcelas);


  /* =======================================================
     FORMATAÇÃO
  ======================================================= */

  aba.getRange(novaLinha, CONFIG.VALOR_TOTAL)
    .setNumberFormat('R$ #,##0.00');

  aba.getRange(novaLinha, CONFIG.VALOR_PARCELA)
    .setNumberFormat('R$ #,##0.00');

  aba.getRange(novaLinha, CONFIG.DATA_PRIMEIRA)
    .setNumberFormat('dd/mm/yyyy');

  aba.getRange(novaLinha, CONFIG.ULTIMA_COMP)
    .setNumberFormat('dd/mm/yyyy');


  SpreadsheetApp.flush();

  Utilities.sleep(400);

  ordenarParcelamentos();

  return `✅ ${dados.produto} cadastrado com sucesso!`;
}


/* =========================================================
   LOCALIZAR LINHA MODELO
========================================================= */

function encontrarLinhaModelo_(aba) {
  const ultimaLinha = aba.getLastRow();

  for (
    let linha = CONFIG.PRIMEIRA_LINHA;
    linha <= ultimaLinha;
    linha++
  ) {

    const produto = aba
      .getRange(linha, CONFIG.PRODUTO)
      .getDisplayValue()
      .trim();

    if (!produto) continue;


    const formulaE = aba
      .getRange(linha, CONFIG.PARCELA_ATUAL)
      .getFormula();

    const formulaF = aba
      .getRange(linha, CONFIG.RESTANTES)
      .getFormula();

    const formulaM = aba
      .getRange(linha, CONFIG.FINALIZADO)
      .getFormula();

    const formulaN = aba
      .getRange(linha, CONFIG.VIGENTE)
      .getFormula();

    const formulaO = aba
      .getRange(linha, CONFIG.ULTIMA_COMP)
      .getFormula();

    const formulaP = aba
      .getRange(linha, CONFIG.MESES_DECORRIDOS)
      .getFormula();


    if (
      formulaE &&
      formulaF &&
      formulaM &&
      formulaN &&
      formulaO &&
      formulaP
    ) {
      return linha;
    }
  }

  return null;
}


/* =========================================================
   CRIAR UMA NOVA LINHA
========================================================= */

function criarNovaLinha_(aba) {
  const ultimaLinhaFisica = aba.getMaxRows();

  aba.insertRowAfter(ultimaLinhaFisica);

  return ultimaLinhaFisica + 1;
}


/* =========================================================
   PREPARAR NOVA LINHA
========================================================= */

function prepararNovaLinha_(aba, linhaModelo, novaLinha) {

  const origemCompleta = aba.getRange(
    linhaModelo,
    1,
    1,
    CONFIG.ULTIMA_COLUNA
  );

  const destinoCompleto = aba.getRange(
    novaLinha,
    1,
    1,
    CONFIG.ULTIMA_COLUNA
  );


  /*
   FORMATAÇÃO
  */
  origemCompleta.copyTo(
    destinoCompleto,
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );


  /*
   DROPDOWNS / CHECKBOX / VALIDAÇÕES
  */
  origemCompleta.copyTo(
    destinoCompleto,
    SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
    false
  );


  /*
   Copia SOMENTE as fórmulas das colunas automáticas.
   O Sheets adapta os números das linhas sozinho.
  */

  copiarFormula_(aba, linhaModelo, novaLinha, CONFIG.PARCELA_ATUAL);   // E
  copiarFormula_(aba, linhaModelo, novaLinha, CONFIG.RESTANTES);      // F
  copiarFormula_(aba, linhaModelo, novaLinha, CONFIG.FINALIZADO);     // M
  copiarFormula_(aba, linhaModelo, novaLinha, CONFIG.VIGENTE);        // N
  copiarFormula_(aba, linhaModelo, novaLinha, CONFIG.ULTIMA_COMP);    // O
  copiarFormula_(aba, linhaModelo, novaLinha, CONFIG.MESES_DECORRIDOS); // P


  /*
   Limpa todos os campos manuais antes de preencher.
  */
  const colunasManuais = [
    CONFIG.CATEGORIA,
    CONFIG.PRODUTO,
    CONFIG.VALOR_TOTAL,
    CONFIG.QTD_PARCELAS,
    CONFIG.VALOR_PARCELA,
    CONFIG.CARTAO,
    CONFIG.RESPONSAVEL,
    CONFIG.OBSERVACAO,
    CONFIG.PAGO,
    CONFIG.DATA_PRIMEIRA
  ];

  colunasManuais.forEach(coluna => {
    aba.getRange(novaLinha, coluna).clearContent();
  });
}


/* =========================================================
   COPIAR UMA FÓRMULA
========================================================= */

function copiarFormula_(aba, origemLinha, destinoLinha, coluna) {

  const origem = aba.getRange(
    origemLinha,
    coluna
  );

  const destino = aba.getRange(
    destinoLinha,
    coluna
  );

  origem.copyTo(
    destino,
    SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
    false
  );
}


/* =========================================================
   VALIDAR DADOS
========================================================= */

function validarDados_(dados) {
  if (!dados.categoria) {
    throw new Error('Informe a categoria.');
  }

  if (!dados.produto || dados.produto.trim() === '') {
    throw new Error('Informe o produto ou serviço.');
  }

  if (!dados.valorTotal) {
    throw new Error('Informe o valor total.');
  }

  const valor = converterNumero_(dados.valorTotal);

  if (valor <= 0) {
    throw new Error('Informe um valor total válido.');
  }

  if (
    !dados.qtdParcelas ||
    Number(dados.qtdParcelas) <= 0
  ) {
    throw new Error(
      'Informe uma quantidade de parcelas válida.'
    );
  }

  if (!dados.cartao) {
    throw new Error('Informe o cartão.');
  }

  if (!dados.responsavel) {
    throw new Error('Informe o responsável.');
  }

  if (!dados.dataPrimeiraParcela) {
    throw new Error(
      'Informe a data da primeira parcela.'
    );
  }
}


/* =========================================================
   CONVERTER VALOR
========================================================= */

function converterNumero_(valor) {
  if (typeof valor === 'number') {
    return valor;
  }

  let texto = String(valor).trim();

  if (
    texto.includes(',') &&
    texto.includes('.')
  ) {

    texto = texto
      .replace(/\./g, '')
      .replace(',', '.');

  } else if (texto.includes(',')) {

    texto = texto.replace(',', '.');
  }

  const numero = Number(texto);

  if (isNaN(numero)) {
    throw new Error('Valor total inválido.');
  }

  return numero;
}


/* =========================================================
   ORDENAR
========================================================= */

function ordenarParcelamentos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG.ABA_PARCELAMENTOS);

  if (!aba) return;

  SpreadsheetApp.flush();

  const primeiraLinha = CONFIG.PRIMEIRA_LINHA;
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < primeiraLinha) return;


  const produtos = aba
    .getRange(
      primeiraLinha,
      CONFIG.PRODUTO,
      ultimaLinha - primeiraLinha + 1,
      1
    )
    .getDisplayValues();


  let ultimaLinhaProduto = primeiraLinha - 1;


  for (let i = 0; i < produtos.length; i++) {

    if (produtos[i][0].trim() !== '') {
      ultimaLinhaProduto = primeiraLinha + i;
    }
  }


  if (ultimaLinhaProduto < primeiraLinha) {
    return;
  }


  const quantidadeLinhas =
    ultimaLinhaProduto - primeiraLinha + 1;


  /*
   A:P é movimentado junto.
   Categoria nunca se separa do Produto.
  */
  aba.getRange(
    primeiraLinha,
    1,
    quantidadeLinhas,
    CONFIG.ULTIMA_COLUNA
  )
  .sort({
    column: CONFIG.RESTANTES,
    ascending: false
  });
}


/* =========================================================
   QUANDO MUDAR O MÊS DO DASHBOARD
========================================================= */

function onEdit(e) {
  if (!e || !e.range) return;

  const aba = e.range.getSheet();

  if (
    aba.getName() === CONFIG.ABA_DASHBOARD &&
    e.range.getA1Notation() === CONFIG.MES_REFERENCIA
  ) {

    SpreadsheetApp.flush();

    Utilities.sleep(700);

    ordenarParcelamentos();
  }
}