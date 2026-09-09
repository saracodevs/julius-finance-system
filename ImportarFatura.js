/* =========================================================
   CONFIGURAÇÕES DO IMPORTADOR
========================================================= */

const FATURA_CONFIG = {
  ABA_DESPESAS: '🛒 Desp.Variáveis',
  ABA_DASHBOARD: '📊 Dashboard',

  LINHA_CABECALHO: 6,
  PRIMEIRA_LINHA: 7,

  COL_DATA: 1,
  COL_CATEGORIA: 2,
  COL_DESCRICAO: 3,
  COL_VALOR: 4,
  COL_FORMA: 5,
  COL_CARTAO: 6,
  COL_RESPONSAVEL: 7,

  COL_DATA_ORIGINAL: 8,
  COL_DESC_ORIGINAL: 9,
  COL_ID: 10
};


/* =========================================================
   ABRIR IMPORTADOR
========================================================= */

function abrirImportarFatura() {

  prepararEstruturaImportacao_();


  if (
    typeof configuracaoInicialCompleta_ === 'function' &&
    !configuracaoInicialCompleta_()
  ) {

    SpreadsheetApp
      .getUi()
      .alert(
        'Configuração necessária',
        'Antes de importar uma fatura, cadastre pelo menos um responsável.',
        SpreadsheetApp
          .getUi()
          .ButtonSet.OK
      );


    abrirConfiguracoes();

    return;
  }


  const html =
    HtmlService
      .createHtmlOutputFromFile(
        'ImportarFaturaa'
      )
      .setWidth(1450)
      .setHeight(850);


  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      '📄 Importar fatura'
    );
}


/* =========================================================
   PREPARAR ESTRUTURA DA PLANILHA
========================================================= */

function prepararEstruturaImportacao_() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const aba =
    ss.getSheetByName(
      FATURA_CONFIG.ABA_DESPESAS
    );


  if (!aba) {

    throw new Error(
      'Não encontrei a aba 🛒 Desp.Variáveis.'
    );
  }


  if (
    aba.getMaxColumns() < 10
  ) {

    aba.insertColumnsAfter(
      aba.getMaxColumns(),
      10 - aba.getMaxColumns()
    );
  }


  aba
    .getRange(
      FATURA_CONFIG.LINHA_CABECALHO,
      8,
      1,
      3
    )
    .setValues([
      [
        'Data Original',
        'Descrição Original',
        'ID Importação'
      ]
    ]);


  aba.hideColumns(
    8,
    3
  );
}


/* =========================================================
   OPÇÕES PARA O IMPORTADOR
========================================================= */

function obterOpcoesImportador() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const dashboard =
    ss.getSheetByName(
      FATURA_CONFIG.ABA_DASHBOARD
    );


  /*
   Agora todas as opções vêm
   da Central de Configurações.
  */

  const config =
    obterConfiguracoesSistema();


  const categorias =
    config.categorias || [];


  const formas =
    config.formas || [];


  const cartoes =
    config.cartoes || [];


  const responsaveis =
    (
      config.responsaveis ||
      []
    )
    .map(
      item =>
        item.nome
    );


  let competencia =
    '';


  if (dashboard) {

    const referencia =
      dashboard
        .getRange('B7')
        .getValue();


    if (
      referencia instanceof Date
    ) {

      const primeiroDia =
        new Date(
          referencia.getFullYear(),
          referencia.getMonth(),
          1,
          12
        );


      competencia =
        Utilities.formatDate(
          primeiroDia,
          ss.getSpreadsheetTimeZone(),
          'yyyy-MM-dd'
        );
    }
  }


  return {

    categorias,

    formas,

    cartoes,

    responsaveis,

    competencia

  };
}

/* =========================================================
   LER DROPDOWNS EXISTENTES
========================================================= */

function obterValidacaoImportador_(
  aba,
  celula
) {

  if (!aba) {
    return [];
  }


  const regra =
    aba
      .getRange(celula)
      .getDataValidation();


  if (!regra) {
    return [];
  }


  const tipo =
    regra.getCriteriaType();


  const valores =
    regra.getCriteriaValues();


  if (
    tipo ===
    SpreadsheetApp
      .DataValidationCriteria
      .VALUE_IN_LIST
  ) {

    return valores[0]
      .filter(
        valor =>
          valor !== ''
      );
  }


  if (
    tipo ===
    SpreadsheetApp
      .DataValidationCriteria
      .VALUE_IN_RANGE
  ) {

    return valores[0]
      .getDisplayValues()
      .flat()
      .filter(
        valor =>
          valor !== ''
      );
  }


  return [];
}


/* =========================================================
   RESPONSÁVEIS
========================================================= */

function obterOpcoesResponsaveis_(
  aba
) {

  if (
    typeof obterResponsaveisCadastrados ===
      'function'
  ) {

    return obterResponsaveisCadastrados();
  }


  return [];
}


/* =========================================================
   PROCESSAR PDF
========================================================= */

function processarFaturaPdf(
  dados
) {

  if (
    !dados ||
    !dados.base64
  ) {

    throw new Error(
      'Nenhum PDF foi recebido.'
    );
  }


  const base64 =
    dados.base64.includes(',')
      ? dados.base64.split(',')[1]
      : dados.base64;


  const bytes =
    Utilities.base64Decode(
      base64
    );


  const blob =
    Utilities.newBlob(
      bytes,
      'application/pdf',
      dados.nome || 'fatura.pdf'
    );


  let arquivoTemporario =
    null;


  try {

    arquivoTemporario =
      Drive.Files.create(
        {
          name:
            'TEMP_FATURA_' +
            new Date().getTime(),

          mimeType:
            'application/vnd.google-apps.document'
        },
        blob,
        {
          ocrLanguage: 'pt',
          fields: 'id,name'
        }
      );


    Utilities.sleep(1800);


    const texto =
      DocumentApp
        .openById(
          arquivoTemporario.id
        )
        .getBody()
        .getText();


    if (
      !texto ||
      !texto.trim()
    ) {

      throw new Error(
        'Não consegui extrair texto deste PDF.'
      );
    }


    const banco =
      identificarBancoFatura_(
        texto
      );


    let resultado;


    if (
      banco === 'NUBANK'
    ) {

      resultado =
        extrairFaturaNubank_(
          texto,
          dados.nome
        );

    } else if (
      banco === 'INTER'
    ) {

      resultado =
        extrairFaturaInter_(
          texto,
          dados.nome
        );

    } else {

      throw new Error(
        'Não reconheci o banco desta fatura.'
      );
    }


    return resultado;


  } finally {

    if (
      arquivoTemporario
    ) {

      try {

        DriveApp
          .getFileById(
            arquivoTemporario.id
          )
          .setTrashed(true);

      } catch (erro) {}
    }
  }
}


/* =========================================================
   IDENTIFICAR BANCO
========================================================= */

function identificarBancoFatura_(
  texto
) {

  const t =
    normalizarFatura_(
      texto
    );


  if (
    t.includes('NU PAGAMENTOS') ||
    t.includes('NUCEL') ||
    t.includes('NUBANK')
  ) {

    return 'NUBANK';
  }


  if (
    t.includes('BANCO INTER') ||
    t.includes('DESPESAS DA FATURA') ||
    t.includes('SUPER APP')
  ) {

    return 'INTER';
  }


  return 'DESCONHECIDO';
}


/* =========================================================
   NUBANK
========================================================= */

function extrairFaturaNubank_(
  texto,
  nomeArquivo
) {

  const linhas =
    texto
      .split(/\r?\n/)
      .map(
        linha =>
          linha
            .replace(
              /\s+/g,
              ' '
            )
            .trim()
      )
      .filter(Boolean);


  const meses =
    mapaMeses_();


  let anoFatura =
    new Date()
      .getFullYear();


  let mesVencimento =
    new Date()
      .getMonth();


  for (
    const linha of linhas
  ) {

    const normal =
      normalizarFatura_(
        linha
      );


    const match =
      normal.match(
        /FATURA\s+\d{1,2}\s+(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\s+(\d{4})/
      );


    if (match) {

      mesVencimento =
        meses[
          match[1]
        ];


      anoFatura =
        Number(
          match[2]
        );


      break;
    }
  }


  const encontrados =
    [];


  let indiceOriginal =
    0;


  for (
    const linhaOriginal of linhas
  ) {

    const linha =
      linhaOriginal
        .replace(
          /−/g,
          '-'
        )
        .replace(
          /–/g,
          '-'
        );


    const inicio =
      normalizarFatura_(
        linha
      )
      .match(
        /^(\d{1,2})\s+(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\s+/
      );


    if (!inicio) {
      continue;
    }


    const valorMatch =
      linha.match(
        /([+-]?\s*R\$\s*[\d.]+,\d{2})\s*$/
      );


    if (!valorMatch) {
      continue;
    }


    const dia =
      Number(
        inicio[1]
      );


    const mes =
      meses[
        inicio[2]
      ];


    let ano =
      anoFatura;


    if (
      mesVencimento === 0 &&
      mes === 11
    ) {

      ano--;
    }


    const valor =
      converterMoedaFatura_(
        valorMatch[1]
      );


    let descricao =
      linha
        .replace(
          /^\d{1,2}\s+(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\s+/i,
          ''
        )
        .replace(
          valorMatch[0],
          ''
        )
        .trim();


    let finalCartao =
      '';


    const finalMatch =
      descricao.match(
        /(?:•{2,}|\.{3,}|\*{3,})\s*(\d{4})\s+/i
      );


    if (finalMatch) {

      finalCartao =
        finalMatch[1];


      descricao =
        descricao
          .replace(
            finalMatch[0],
            ''
          )
          .trim();
    }


    encontrados.push(
      criarTransacaoFatura_({
        banco: 'Nubank',

        indice:
          indiceOriginal++,

        dataOriginal:
          criarIsoFatura_(
            ano,
            mes,
            dia
          ),

        descricao,

        valor,

        finalCartao
      })
    );
  }


  return finalizarResultadoFatura_(
    'Nubank',
    nomeArquivo,
    encontrados
  );
}


/* =========================================================
   INTER
========================================================= */

function extrairFaturaInter_(
  texto,
  nomeArquivo
) {

  const linhas =
    String(texto || '')
      .replace(
        /−/g,
        '-'
      )
      .replace(
        /–/g,
        '-'
      )
      .split(/\r?\n/)
      .map(
        linha =>
          linha.trim()
      )
      .filter(Boolean);


  let inicio =
    -1;


  let fim =
    linhas.length;


  for (
    let i = 0;
    i < linhas.length;
    i++
  ) {

    const normal =
      normalizarFatura_(
        linhas[i]
      );


    if (
      inicio < 0 &&
      normal ===
        'DESPESAS DA FATURA'
    ) {

      inicio =
        i;

      continue;
    }


    if (
      inicio >= 0 &&
      normal ===
        'PROXIMA FATURA'
    ) {

      fim =
        i;

      break;
    }
  }


  if (
    inicio < 0
  ) {

    throw new Error(
      'Não encontrei a área de despesas da fatura do Inter.'
    );
  }


  const secao =
    linhas.slice(
      inicio + 1,
      fim
    );


  const transacoes =
    extrairEstruturaInter_(
      secao
    );


  if (
    !transacoes.length
  ) {

    throw new Error(
      'Não consegui localizar as compras do Inter.'
    );
  }


  const valores =
    extrairValoresInter_(
      secao
    );


  if (
    valores.length <
    transacoes.length
  ) {

    throw new Error(
      'Encontrei ' +
      transacoes.length +
      ' compras, mas somente ' +
      valores.length +
      ' valores no OCR.'
    );
  }


  const valoresResolvidos =
    reconciliarValoresInter_(
      transacoes,
      valores
    );


  if (
    valoresResolvidos.length !==
    transacoes.length
  ) {

    throw new Error(
      'Não consegui relacionar corretamente compras e valores do Inter.'
    );
  }


  const encontrados =
    [];


  transacoes.forEach(
    (
      transacao,
      indice
    ) => {

      encontrados.push(
        criarTransacaoFatura_({
          banco: 'Inter',

          indice,

          dataOriginal:
            transacao.dataOriginal,

          descricao:
            transacao.descricao,

          valor:
            valoresResolvidos[
              indice
            ],

          finalCartao:
            transacao.finalCartao
        })
      );
    }
  );


  return finalizarResultadoFatura_(
    'Inter',
    nomeArquivo,
    encontrados
  );
}


/* =========================================================
   INTER - EXTRAIR DATA / DESCRIÇÃO
========================================================= */

function extrairEstruturaInter_(
  linhas
) {

  const resultados =
    [];


  let cartaoAtual =
    '';


  let dataPendente =
    null;


  for (
    let i = 0;
    i < linhas.length;
    i++
  ) {

    const linha =
      linhas[i];


    const normal =
      normalizarFatura_(
        linha
      );


    const cartaoMatch =
      normal.match(
        /CARTAO\s+\d{4}\*+(\d{4})/
      );


    if (
      cartaoMatch
    ) {

      cartaoAtual =
        cartaoMatch[1];

      continue;
    }


    const dataMatch =
      normal.match(
        /^(\d{1,2})\s+DE\s+(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[A-Z.]*\s+(\d{4})$/
      );


    if (
      dataMatch
    ) {

      dataPendente = {
        dia:
          Number(
            dataMatch[1]
          ),

        mes:
          mapaMeses_()[
            dataMatch[2]
          ],

        ano:
          Number(
            dataMatch[3]
          )
      };


      continue;
    }


    if (
      dataPendente &&
      linhaPodeSerDescricaoInter_(
        linha
      )
    ) {

      resultados.push({
        dataOriginal:
          criarIsoFatura_(
            dataPendente.ano,
            dataPendente.mes,
            dataPendente.dia
          ),

        descricao:
          linha,

        finalCartao:
          cartaoAtual
      });


      dataPendente =
        null;
    }
  }


  return resultados;
}


/* =========================================================
   INTER - VALIDAR DESCRIÇÃO
========================================================= */

function linhaPodeSerDescricaoInter_(
  linha
) {

  const normal =
    normalizarFatura_(
      linha
    );


  if (!normal) {
    return false;
  }


  if (
    /^[+-]?\s*R\$\s*[\d.]+,\d{2}$/i
      .test(
        linha
      )
  ) {

    return false;
  }


  if (
    /^\d{1,2}\s+DE\s+/i
      .test(
        normal
      )
  ) {

    return false;
  }


  const proibidos = [
    'DATA',
    'MOVIMENTACAO',
    'BENEFICIARIO',
    'VALOR',
    'VALOR TOTAL',
    'VENCIMENTO',
    'DESPESAS DA FATURA',
    'TOTAL CARTAO',
    'CARTAO ',
    'VICTORIA SILVA MONTEIRO',
    'WINTER',
    'INTER',
    'PROXIMA FATURA'
  ];


  for (
    const texto of proibidos
  ) {

    if (
      normal === texto ||
      normal.startsWith(
        texto
      )
    ) {

      return false;
    }
  }


  if (
    /^\d{4}\*+\d{4}$/
      .test(
        normal
      )
  ) {

    return false;
  }


  if (
    /^\d{2}\/\d{2}\/\d{4}$/
      .test(
        normal
      )
  ) {

    return false;
  }


  return true;
}


/* =========================================================
   INTER - EXTRAIR VALORES
========================================================= */

function extrairValoresInter_(
  linhas
) {

  const valores =
    [];


  for (
    let i = 0;
    i < linhas.length;
    i++
  ) {

    const linha =
      linhas[i];


    if (
      !/^[+-]?\s*R\$\s*[\d.]+,\d{2}$/i
        .test(
          linha
        )
    ) {

      continue;
    }


    let cabecalhoTotal =
      false;


    for (
      let volta = 1;
      volta <= 3;
      volta++
    ) {

      if (
        i - volta < 0
      ) {

        break;
      }


      const anterior =
        normalizarFatura_(
          linhas[
            i - volta
          ]
        );


      if (
        anterior ===
        'VALOR TOTAL'
      ) {

        cabecalhoTotal =
          true;

        break;
      }


      if (
        anterior ===
          'VALOR' ||
        anterior ===
          'BENEFICIARIO'
      ) {

        break;
      }
    }


    if (
      cabecalhoTotal
    ) {

      continue;
    }


    valores.push(
      converterMoedaFatura_(
        linha
      )
    );
  }


  return valores;
}


/* =========================================================
   INTER - RECONCILIAR VALORES
========================================================= */

function reconciliarValoresInter_(
  transacoes,
  candidatos
) {

  const quantidadeTransacoes =
    transacoes.length;


  if (
    candidatos.length ===
    quantidadeTransacoes
  ) {

    return candidatos;
  }


  const excesso =
    candidatos.length -
    quantidadeTransacoes;


  if (
    excesso <= 0
  ) {

    return [];
  }


  if (
    candidatos.length > 25 ||
    excesso > 5
  ) {

    throw new Error(
      'O OCR do Inter trouxe valores demais para reconciliar automaticamente.'
    );
  }


  const combinacoes =
    gerarCombinacoes_(
      candidatos.length,
      excesso
    );


  let melhor =
    null;


  let melhorPontuacao =
    Infinity;


  for (
    const ignorados of combinacoes
  ) {

    const ignoradosSet =
      new Set(
        ignorados
      );


    const restantes =
      candidatos.filter(
        (
          valor,
          indice
        ) =>
          !ignoradosSet.has(
            indice
          )
      );


    if (
      restantes.length !==
      quantidadeTransacoes
    ) {

      continue;
    }


    const somas =
      {};


    transacoes.forEach(
      (
        transacao,
        indice
      ) => {

        const descricao =
          normalizarFatura_(
            transacao.descricao
          );


        if (
          descricao.includes(
            'PAGAMENTO'
          )
        ) {

          return;
        }


        const cartao =
          transacao.finalCartao ||
          'SEM_CARTAO';


        somas[cartao] =
          (
            somas[cartao] ||
            0
          ) +
          Math.abs(
            restantes[
              indice
            ]
          );
      }
    );


    const totaisCalculados =
      Object.values(
        somas
      );


    const valoresIgnorados =
      ignorados.map(
        indice =>
          Math.abs(
            candidatos[
              indice
            ]
          )
      );


    const pontuacao =
      calcularPontuacaoTotais_(
        totaisCalculados,
        valoresIgnorados
      );


    if (
      pontuacao <
      melhorPontuacao
    ) {

      melhorPontuacao =
        pontuacao;


      melhor =
        restantes;
    }
  }


  if (
    melhor &&
    melhorPontuacao < 1
  ) {

    return melhor;
  }


  throw new Error(
    'Encontrei as compras do Inter, mas não consegui confirmar automaticamente quais valores são os totais dos cartões.'
  );
}


/* =========================================================
   COMPARAR TOTAIS
========================================================= */

function calcularPontuacaoTotais_(
  calculados,
  ignorados
) {

  if (
    !calculados.length
  ) {

    return Infinity;
  }


  const usados =
    new Set();


  let pontuacao =
    0;


  for (
    const calculado of calculados
  ) {

    let melhorIndice =
      -1;


    let melhorDiferenca =
      Infinity;


    for (
      let i = 0;
      i < ignorados.length;
      i++
    ) {

      if (
        usados.has(i)
      ) {

        continue;
      }


      const diferenca =
        Math.abs(
          calculado -
          ignorados[i]
        );


      if (
        diferenca <
        melhorDiferenca
      ) {

        melhorDiferenca =
          diferenca;


        melhorIndice =
          i;
      }
    }


    if (
      melhorIndice >= 0
    ) {

      usados.add(
        melhorIndice
      );


      pontuacao +=
        melhorDiferenca;

    } else {

      pontuacao +=
        999999;
    }
  }


  return pontuacao;
}


/* =========================================================
   GERAR COMBINAÇÕES
========================================================= */

function gerarCombinacoes_(
  tamanho,
  quantidade
) {

  const resultado =
    [];


  function montar(
    inicio,
    atual
  ) {

    if (
      atual.length ===
      quantidade
    ) {

      resultado.push(
        [...atual]
      );


      return;
    }


    for (
      let i = inicio;
      i < tamanho;
      i++
    ) {

      atual.push(i);


      montar(
        i + 1,
        atual
      );


      atual.pop();
    }
  }


  montar(
    0,
    []
  );


  return resultado;
}


/* =========================================================
   CRIAR TRANSAÇÃO
========================================================= */

function criarTransacaoFatura_(
  dados
) {

  const descricaoNormal =
    normalizarFatura_(
      dados.descricao
    );


  const parcelado =
    /PARCELA\s+\d+\s*(?:\/|DE)\s*\d+/i
      .test(
        descricaoNormal
      );


  const pagamento =
    descricaoNormal.includes(
      'PAGAMENTO'
    ) ||
    descricaoNormal.includes(
      'PAGTO'
    ) ||
    descricaoNormal.includes(
      'SALDO RESTANTE DA FATURA'
    );


  const credito =
    dados.valor < 0 ||
    descricaoNormal.includes(
      'ESTORNO'
    ) ||
    descricaoNormal.includes(
      'AJUSTE A CREDITO'
    );


  let tipo =
    'COMPRA';


  if (
    parcelado
  ) {

    tipo =
      'PARCELADO';
  }


  if (
    credito
  ) {

    tipo =
      'CREDITO';
  }


  if (
    pagamento
  ) {

    tipo =
      'PAGAMENTO';
  }


  return {
    banco:
      dados.banco,

    indice:
      dados.indice,

    dataOriginal:
      dados.dataOriginal,

    dataControle:
      dados.dataOriginal,

    descricaoOriginal:
      dados.descricao,

    descricao:
      dados.descricao,

    valor:
      dados.valor,

    finalCartao:
      dados.finalCartao || '',

    tipo,

    importar:
      tipo === 'COMPRA'
  };
}


/* =========================================================
   FINALIZAR RESULTADO
========================================================= */

function finalizarResultadoFatura_(
  banco,
  nomeArquivo,
  transacoes
) {

  if (
    !transacoes.length
  ) {

    throw new Error(
      'Não encontrei transações reconhecíveis nesta fatura.'
    );
  }


  const ocorrencias =
    {};


  transacoes.forEach(
    transacao => {

      const assinatura =
        [
          banco,
          transacao.dataOriginal,

          normalizarFatura_(
            transacao
              .descricaoOriginal
          ),

          Number(
            transacao.valor
          )
          .toFixed(2),

          transacao.finalCartao

        ]
        .join('|');


      ocorrencias[
        assinatura
      ] =
        (
          ocorrencias[
            assinatura
          ] ||
          0
        ) + 1;


      transacao.id =
        gerarIdFatura_(
          assinatura +
          '|' +
          ocorrencias[
            assinatura
          ]
        );
    }
  );


  return {
    banco,

    arquivo:
      nomeArquivo || '',

    quantidade:
      transacoes.length,

    transacoes
  };
}


/* =========================================================
   SALVAR IMPORTAÇÃO
========================================================= */

function salvarImportacaoFatura(
  dados
) {

  prepararEstruturaImportacao_();


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const aba =
    ss.getSheetByName(
      FATURA_CONFIG
        .ABA_DESPESAS
    );


  if (!aba) {

    throw new Error(
      'Não encontrei a aba 🛒 Desp.Variáveis.'
    );
  }


  const itens =
    (
      dados.itens ||
      []
    )
    .filter(
      item =>
        item.importar
    );


  if (
    !itens.length
  ) {

    throw new Error(
      'Nenhuma transação foi selecionada.'
    );
  }


  const ultimaLinha =
    aba.getLastRow();


  let idsExistentes =
    [];


  if (
    ultimaLinha >=
    FATURA_CONFIG.PRIMEIRA_LINHA
  ) {

    const quantidadeLinhas =
      ultimaLinha -
      FATURA_CONFIG.PRIMEIRA_LINHA +
      1;


    const linhasExistentes =
      aba
        .getRange(
          FATURA_CONFIG.PRIMEIRA_LINHA,
          1,
          quantidadeLinhas,
          10
        )
        .getDisplayValues();


    idsExistentes =
      linhasExistentes
        .filter(
          linha => {

            const data =
              linha[0];


            const descricao =
              linha[2];


            const valor =
              linha[3];


            const id =
              linha[9];


            return (
              id &&
              data &&
              descricao &&
              valor
            );
          }
        )
        .map(
          linha =>
            linha[9]
        );
  }


  const conjunto =
    new Set(
      idsExistentes
    );


  const novos =
    itens.filter(
      item =>
        !conjunto.has(
          item.id
        )
    );


  if (
    !novos.length
  ) {

    throw new Error(
      'Todas as transações selecionadas já foram importadas.'
    );
  }


  const linhaInicial =
    aba.getLastRow() + 1;


  const linhaFinal =
    linhaInicial +
    novos.length -
    1;


  if (
    linhaFinal >
    aba.getMaxRows()
  ) {

    aba.insertRowsAfter(
      aba.getMaxRows(),
      linhaFinal -
      aba.getMaxRows()
    );
  }


  const modelo =
    aba.getRange(
      FATURA_CONFIG.PRIMEIRA_LINHA,
      1,
      1,
      7
    );


  for (
    let i = 0;
    i < novos.length;
    i++
  ) {

    const destino =
      aba.getRange(
        linhaInicial + i,
        1,
        1,
        7
      );


    modelo.copyTo(
      destino,
      SpreadsheetApp
        .CopyPasteType
        .PASTE_FORMAT,
      false
    );


    modelo.copyTo(
      destino,
      SpreadsheetApp
        .CopyPasteType
        .PASTE_DATA_VALIDATION,
      false
    );
  }


  const valores =
    novos.map(
      item => [

        criarDataImportacao_(
          item.dataControle
        ),

        item.categoria,

        item.descricao,

        Number(
          item.valor
        ),

        item.formaPagamento ||
          'Crédito',

        dados.cartao,

        item.responsavel,

        criarDataImportacao_(
          item.dataOriginal
        ),

        item.descricaoOriginal,

        item.id
      ]
    );


  aba
    .getRange(
      linhaInicial,
      1,
      valores.length,
      10
    )
    .setValues(
      valores
    );


  aba
    .getRange(
      linhaInicial,
      1,
      valores.length,
      1
    )
    .setNumberFormat(
      'dd/mm/yyyy'
    );


  aba
    .getRange(
      linhaInicial,
      8,
      valores.length,
      1
    )
    .setNumberFormat(
      'dd/mm/yyyy'
    );


  aba
    .getRange(
      linhaInicial,
      4,
      valores.length,
      1
    )
    .setNumberFormat(
      'R$ #,##0.00'
    );


  aba.hideColumns(
    8,
    3
  );


  SpreadsheetApp.flush();


  if (
    typeof filtrarDespesasVariaveisPorMes ===
      'function'
  ) {

    filtrarDespesasVariaveisPorMes();
  }


  return {
    importadas:
      novos.length,

    duplicadas:
      itens.length -
      novos.length
  };
}


/* =========================================================
   UTILITÁRIOS
========================================================= */

function normalizarFatura_(
  texto
) {

  return String(
    texto ||
    ''
  )
  .normalize('NFD')
  .replace(
    /[\u0300-\u036f]/g,
    ''
  )
  .toUpperCase()
  .replace(
    /\s+/g,
    ' '
  )
  .trim();
}


function converterMoedaFatura_(
  texto
) {

  let valor =
    String(texto)
      .replace(
        /\s+/g,
        ''
      )
      .replace(
        'R$',
        ''
      );


  let negativo =
    false;


  if (
    valor.startsWith('-')
  ) {

    negativo =
      true;


    valor =
      valor.substring(1);
  }


  if (
    valor.startsWith('+')
  ) {

    valor =
      valor.substring(1);
  }


  const numero =
    Number(
      valor
        .replace(
          /\./g,
          ''
        )
        .replace(
          ',',
          '.'
        )
    );


  if (
    isNaN(numero)
  ) {

    return 0;
  }


  return negativo
    ? -numero
    : numero;
}


function mapaMeses_() {

  return {
    JAN: 0,
    FEV: 1,
    MAR: 2,
    ABR: 3,
    MAI: 4,
    JUN: 5,
    JUL: 6,
    AGO: 7,
    SET: 8,
    OUT: 9,
    NOV: 10,
    DEZ: 11
  };
}


function criarIsoFatura_(
  ano,
  mes,
  dia
) {

  const data =
    new Date(
      ano,
      mes,
      dia,
      12
    );


  return Utilities
    .formatDate(
      data,
      SpreadsheetApp
        .getActive()
        .getSpreadsheetTimeZone(),
      'yyyy-MM-dd'
    );
}


function criarDataImportacao_(
  iso
) {

  const partes =
    String(
      iso
    )
    .split('-');


  return new Date(
    Number(
      partes[0]
    ),
    Number(
      partes[1]
    ) - 1,
    Number(
      partes[2]
    ),
    12
  );
}


function gerarIdFatura_(
  texto
) {

  const digest =
    Utilities.computeDigest(
      Utilities
        .DigestAlgorithm
        .MD5,

      texto,

      Utilities
        .Charset
        .UTF_8
    );


  return Utilities
    .base64EncodeWebSafe(
      digest
    )
    .replace(
      /=/g,
      ''
    );
}