/**
 * ============================================================
 * 💸 REEMBOLSOS — JULIUS FINANCE SYSTEM V3
 * ============================================================
 *
 * MOTOR DE CONTAS A RECEBER
 *
 * A Data
 * B Responsável
 * C Origem
 * D Descrição
 * E Valor Devido
 * F Valor Pago          <- agora calculado pelas alocações
 * G Saldo
 * H Status
 * I Observação
 * J Competência         <- técnica
 * K Chave Sistema       <- técnica
 * L Fonte               <- técnica
 *
 * IMPORTANTE:
 * - F não deve mais ser preenchido manualmente.
 * - Pagamentos ficam em "💵 Pagamentos Reembolsos".
 * - Rateios ficam em "_REEMB_ALOCACOES".
 * - G/H continuam automáticos.
 * - A sincronização continua compatível com Automacoes.gs.
 */


const REEMBOLSOS_CONFIG = {

  ABA:
    '💸 Reembolsos',

  DASHBOARD:
    '📊 Dashboard',

  CELULA_COMPETENCIA:
    'B7',

  DESPESAS_VARIAVEIS:
    '🛒 Desp.Variáveis',

  PARCELAMENTOS:
    '💳 Parcelamentos',

  CONFIG:
    '_CONFIG',

  ABA_ALOCACOES:
    '_REEMB_ALOCACOES',

  LINHA_CABECALHO:
    4,

  PRIMEIRA_LINHA:
    5,

  TIMEZONE:
    'America/Sao_Paulo'
};


/**
 * ============================================================
 * 🚀 PREPARAR
 * ============================================================
 */

function prepararReembolsos() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  corrigirTimezoneReembolsos_(
    ss
  );


  let aba =
    ss.getSheetByName(
      REEMBOLSOS_CONFIG.ABA
    );


  if (!aba) {

    aba =
      ss.insertSheet(
        REEMBOLSOS_CONFIG.ABA
      );

  }


  prepararCabecalhoReembolsos_(
    aba
  );


  aplicarValidacoesReembolsos_();


  aplicarFormatosBasicosReembolsos_(
    aba
  );


  aba.hideColumns(
    10,
    3
  );


  SpreadsheetApp.flush();

}


/**
 * ============================================================
 * 🔄 SINCRONIZAÇÃO PRINCIPAL
 * ============================================================
 */

function sincronizarReembolsos() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  corrigirTimezoneReembolsos_(
    ss
  );


  let aba =
    ss.getSheetByName(
      REEMBOLSOS_CONFIG.ABA
    );


  if (!aba) {

    prepararReembolsos();

    aba =
      ss.getSheetByName(
        REEMBOLSOS_CONFIG.ABA
      );

  }


  /**
   * Competência atual.
   */

  const competenciaAtual =
    obterCompetenciaAtualReembolsos_(
      ss
    );


  if (!competenciaAtual) {

    throw new Error(
      'Não foi possível identificar a competência em 📊 Dashboard!B7.'
    );

  }


  /**
   * Pagamentos já distribuídos.
   *
   * Map:
   * chave da dívida -> total alocado
   */

  const alocacoes =
    obterAlocacoesReembolsosPorChave_(
      ss
    );


  /**
   * Linhas manuais existentes.
   */

  const manuais =
    obterRegistrosManuaisReembolsos_(
      aba,
      alocacoes
    );


  /**
   * Parcelamentos históricos.
   *
   * Preservamos meses diferentes da competência
   * atualmente selecionada.
   */

  const historicoParcelamentos =
    obterParcelamentosHistoricosReembolsos_(
      aba,
      competenciaAtual,
      alocacoes
    );


  /**
   * Despesas variáveis de terceiros.
   */

  const despesas =
    obterReembolsosDespesasVariaveis_(
      ss,
      alocacoes
    );


  /**
   * Parcelamentos da competência selecionada.
   */

  const parcelamentos =
    obterReembolsosParcelamentos_(
      ss,
      competenciaAtual,
      alocacoes
    );


  /**
   * Combina tudo.
   */

  let registros =
    despesas
      .concat(
        historicoParcelamentos
      )
      .concat(
        parcelamentos
      )
      .concat(
        manuais
      );


  /**
   * Segurança contra duplicatas de chave.
   */

  registros =
    removerDuplicatasReembolsos_(
      registros
    );


  /**
   * Visual:
   * mais recentes primeiro.
   */

  registros.sort(
    compararReembolsosMaisRecentes_
  );


  /**
   * Limpa somente conteúdo da base.
   *
   * Linhas 1:4 e visual permanecem.
   */

  const ultimaLinhaAnterior =
    Math.max(
      aba.getLastRow(),
      REEMBOLSOS_CONFIG.PRIMEIRA_LINHA
    );


  if (
    ultimaLinhaAnterior >=
    REEMBOLSOS_CONFIG.PRIMEIRA_LINHA
  ) {

    aba
      .getRange(
        REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
        1,
        ultimaLinhaAnterior -
          REEMBOLSOS_CONFIG.PRIMEIRA_LINHA +
          1,
        12
      )
      .clearContent();

  }


  /**
   * Escreve os registros.
   */

  if (
    registros.length > 0
  ) {

    const linhas =
      registros.map(
        registro =>
          registroParaLinhaReembolsos_(
            registro
          )
      );


    aba
      .getRange(
        REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
        1,
        linhas.length,
        12
      )
      .setValues(
        linhas
      );

  }


  /**
   * G/H automáticos.
   */

  aplicarFormulasReembolsos_(
    aba
  );


  aplicarValidacoesReembolsos_();


  aplicarFormatosBasicosReembolsos_(
    aba
  );


  /**
   * F passa a ser somente leitura conceitualmente.
   *
   * Não bloqueamos a célula porque isso exigiria
   * proteção/permissões, mas o valor será reconstruído
   * pelas alocações em toda sincronização.
   */

  aba
    .getRange('F4')
    .setNote(
      'Valor calculado automaticamente a partir dos pagamentos registrados em 💵 Pagamentos Reembolsos. Não preencher manualmente.'
    );


  aba.hideColumns(
    10,
    3
  );


  SpreadsheetApp.flush();


  ss.toast(
    registros.length +
      ' dívidas sincronizadas 💸',
    'Julius Finance',
    4
  );

}


/**
 * ============================================================
 * 💳 DESPESAS VARIÁVEIS
 * ============================================================
 */

function obterReembolsosDespesasVariaveis_(
  ss,
  alocacoes
) {

  const aba =
    ss.getSheetByName(
      REEMBOLSOS_CONFIG.DESPESAS_VARIAVEIS
    );


  if (!aba) {
    return [];
  }


  const ultimaLinha =
    aba.getLastRow();


  if (
    ultimaLinha < 7
  ) {
    return [];
  }


  /**
   * A:J
   *
   * A Data
   * B Categoria
   * C Descrição
   * D Valor
   * E Forma Pagamento
   * F Cartão
   * G Responsável
   * H Data Original
   * I Descrição Original
   * J ID Importação
   */

  const dados =
    aba
      .getRange(
        7,
        1,
        ultimaLinha - 6,
        10
      )
      .getValues();


  const terceiros =
    obterSetResponsaveisTerceiros_(
      ss
    );


  const vistos =
    new Set();


  const registros =
    [];


  dados.forEach(
    linha => {

      const data =
        linha[0];


      const categoria =
        String(
          linha[1] || ''
        ).trim();


      const descricao =
        String(
          linha[2] || ''
        ).trim();


      const valor =
        Number(
          linha[3]
        ) || 0;


      const formaPagamento =
        String(
          linha[4] || ''
        ).trim();


      const cartao =
        String(
          linha[5] || ''
        ).trim();


      const responsavel =
        String(
          linha[6] || ''
        ).trim();


      const dataOriginal =
        linha[7];


      const descricaoOriginal =
        String(
          linha[8] || ''
        ).trim();


      const idImportacao =
        String(
          linha[9] || ''
        ).trim();


      if (
        !terceiros.has(
          normalizarTextoReembolsos_(
            responsavel
          )
        )
      ) {
        return;
      }


      if (
        !(data instanceof Date) ||
        isNaN(
          data.getTime()
        )
      ) {
        return;
      }


      if (
        valor <= 0
      ) {
        return;
      }


      const dataLimpa =
        normalizarDataReembolsos_(
          data
        );


      const competencia =
        primeiroDiaMesReembolsos_(
          dataLimpa
        );


      const origem =
        cartao ||
        formaPagamento ||
        'Outro';


      let chave;


      if (
        idImportacao
      ) {

        chave =
          'VAR|ID|' +
          limparChaveReembolsos_(
            idImportacao
          );

      } else {

        const assinatura =
          montarAssinaturaDespesaVariavel_(
            dataLimpa,
            responsavel,
            origem,
            descricao,
            valor,
            categoria,
            dataOriginal,
            descricaoOriginal
          );


        chave =
          'VAR|AUTO|' +
          gerarHashReembolsos_(
            assinatura
          );

      }


      /**
       * Remove eventual repetição da mesma obrigação.
       */

      if (
        vistos.has(
          chave
        )
      ) {
        return;
      }


      vistos.add(
        chave
      );


      const valorPago =
        limitarValorPagoReembolsos_(
          alocacoes.get(
            chave
          ) || 0,
          valor
        );


      registros.push({

        data:
          dataLimpa,

        responsavel:
          responsavel,

        origem:
          origem,

        descricao:
          descricao,

        valorDevido:
          valor,

        valorPago:
          valorPago,

        observacao:
          categoria,

        competencia:
          competencia,

        chave:
          chave,

        fonte:
          'Despesa Variável'
      });

    }
  );


  return registros;

}


/**
 * ============================================================
 * 💳 PARCELAMENTOS — COMPETÊNCIA ATUAL
 * ============================================================
 */

function obterReembolsosParcelamentos_(
  ss,
  competenciaAtual,
  alocacoes
) {

  const aba =
    ss.getSheetByName(
      REEMBOLSOS_CONFIG.PARCELAMENTOS
    );


  if (!aba) {
    return [];
  }


  const ultimaLinha =
    aba.getLastRow();


  if (
    ultimaLinha < 4
  ) {
    return [];
  }


  const dados =
    aba
      .getRange(
        4,
        1,
        ultimaLinha - 3,
        16
      )
      .getValues();


  const terceiros =
    obterSetResponsaveisTerceiros_(
      ss
    );


  const registros =
    [];


  dados.forEach(
    linha => {

      const categoria =
        String(
          linha[0] || ''
        ).trim();


      const produto =
        String(
          linha[1] || ''
        ).trim();


      const valorTotal =
        Number(
          linha[2]
        ) || 0;


      const qtdParcelas =
        Number(
          linha[3]
        ) || 0;


      const parcelaAtual =
        Number(
          linha[4]
        ) || 0;


      const valorParcela =
        Number(
          linha[6]
        ) || 0;


      const cartao =
        String(
          linha[7] || ''
        ).trim();


      const responsavel =
        String(
          linha[8] || ''
        ).trim();


      const observacao =
        String(
          linha[9] || ''
        ).trim();


      /**
       * K = Pago
       *
       * Continua ignorado nos cálculos.
       */

      const primeiraParcela =
        linha[11];


      const vigente =
        String(
          linha[13] || ''
        ).trim();


      if (
        !terceiros.has(
          normalizarTextoReembolsos_(
            responsavel
          )
        )
      ) {
        return;
      }


      if (
        vigente !== 'Sim'
      ) {
        return;
      }


      if (
        valorParcela <= 0 ||
        !produto
      ) {
        return;
      }


      const primeiraParcelaLimpa =
        primeiraParcela instanceof Date
          ? normalizarDataReembolsos_(
              primeiraParcela
            )
          : null;


      const dataCompetencia =
        primeiroDiaMesReembolsos_(
          competenciaAtual
        );


      const descricao =
        produto +
        ' • Parcela ' +
        parcelaAtual +
        '/' +
        qtdParcelas;


      const chave =
        montarChaveParcelamentoReembolsos_(
          responsavel,
          produto,
          valorTotal,
          qtdParcelas,
          valorParcela,
          cartao,
          primeiraParcelaLimpa,
          dataCompetencia
        );


      const valorPago =
        limitarValorPagoReembolsos_(
          alocacoes.get(
            chave
          ) || 0,
          valorParcela
        );


      registros.push({

        data:
          dataCompetencia,

        responsavel:
          responsavel,

        origem:
          cartao ||
          'Parcelamento',

        descricao:
          descricao,

        valorDevido:
          valorParcela,

        valorPago:
          valorPago,

        observacao:
          [
            categoria,
            observacao
          ]
            .filter(Boolean)
            .join(' • '),

        competencia:
          dataCompetencia,

        chave:
          chave,

        fonte:
          'Parcelamento'
      });

    }
  );


  return registros;

}


/**
 * ============================================================
 * 📚 PARCELAMENTOS HISTÓRICOS
 * ============================================================
 */

function obterParcelamentosHistoricosReembolsos_(
  aba,
  competenciaAtual,
  alocacoes
) {

  const ultimaLinha =
    aba.getLastRow();


  if (
    ultimaLinha <
    REEMBOLSOS_CONFIG.PRIMEIRA_LINHA
  ) {
    return [];
  }


  const dados =
    aba
      .getRange(
        REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
        1,
        ultimaLinha -
          REEMBOLSOS_CONFIG.PRIMEIRA_LINHA +
          1,
        12
      )
      .getValues();


  const competenciaAtualChave =
    chaveMesReembolsos_(
      competenciaAtual
    );


  const vistos =
    new Set();


  const registros =
    [];


  dados.forEach(
    linha => {

      const chave =
        String(
          linha[10] || ''
        ).trim();


      const fonte =
        String(
          linha[11] || ''
        ).trim();


      if (
        !chave ||
        fonte !== 'Parcelamento'
      ) {
        return;
      }


      let competencia =
        extrairCompetenciaDaChaveParcelamento_(
          chave
        );


      if (
        !competencia &&
        linha[9] instanceof Date
      ) {

        competencia =
          primeiroDiaMesReembolsos_(
            linha[9]
          );

      }


      if (!competencia) {
        return;
      }


      /**
       * A competência atual será reconstruída.
       */

      if (
        chaveMesReembolsos_(
          competencia
        ) ===
        competenciaAtualChave
      ) {
        return;
      }


      if (
        vistos.has(
          chave
        )
      ) {
        return;
      }


      vistos.add(
        chave
      );


      const valorDevido =
        Number(
          linha[4]
        ) || 0;


      registros.push({

        data:
          competencia,

        responsavel:
          linha[1],

        origem:
          linha[2],

        descricao:
          linha[3],

        valorDevido:
          valorDevido,

        valorPago:
          limitarValorPagoReembolsos_(
            alocacoes.get(
              chave
            ) || 0,
            valorDevido
          ),

        observacao:
          linha[8],

        competencia:
          competencia,

        chave:
          chave,

        fonte:
          'Parcelamento'
      });

    }
  );


  return registros;

}


/**
 * ============================================================
 * ✍️ REGISTROS MANUAIS
 * ============================================================
 *
 * Preserva eventuais dívidas criadas manualmente.
 *
 * Na primeira sincronização uma chave MAN será criada.
 */

function obterRegistrosManuaisReembolsos_(
  aba,
  alocacoes
) {

  const ultimaLinha =
    aba.getLastRow();


  if (
    ultimaLinha <
    REEMBOLSOS_CONFIG.PRIMEIRA_LINHA
  ) {
    return [];
  }


  const dados =
    aba
      .getRange(
        REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
        1,
        ultimaLinha -
          REEMBOLSOS_CONFIG.PRIMEIRA_LINHA +
          1,
        12
      )
      .getValues();


  const registros =
    [];


  dados.forEach(
    linha => {

      const chaveAtual =
        String(
          linha[10] || ''
        ).trim();


      const fonteAtual =
        String(
          linha[11] || ''
        ).trim();


      /**
       * Automáticos não são manuais.
       */

      if (
        chaveAtual &&
        fonteAtual
      ) {
        return;
      }


      const data =
        linha[0];


      const responsavel =
        String(
          linha[1] || ''
        ).trim();


      const origem =
        String(
          linha[2] || ''
        ).trim();


      const descricao =
        String(
          linha[3] || ''
        ).trim();


      const valor =
        Number(
          linha[4]
        ) || 0;


      const observacao =
        String(
          linha[8] || ''
        ).trim();


      if (
        !responsavel &&
        !descricao &&
        valor <= 0
      ) {
        return;
      }


      const dataLimpa =
        data instanceof Date
          ? normalizarDataReembolsos_(
              data
            )
          : normalizarDataReembolsos_(
              new Date()
            );


      const competencia =
        primeiroDiaMesReembolsos_(
          dataLimpa
        );


      const assinatura =
        [
          formatarDataChaveReembolsos_(
            dataLimpa
          ),
          limparChaveReembolsos_(
            responsavel
          ),
          limparChaveReembolsos_(
            origem
          ),
          limparChaveReembolsos_(
            descricao
          ),
          numeroChaveReembolsos_(
            valor
          )
        ].join('|');


      const chave =
        chaveAtual ||
        (
          'MAN|' +
          gerarHashReembolsos_(
            assinatura
          )
        );


      registros.push({

        data:
          dataLimpa,

        responsavel:
          responsavel,

        origem:
          origem,

        descricao:
          descricao,

        valorDevido:
          valor,

        valorPago:
          limitarValorPagoReembolsos_(
            alocacoes.get(
              chave
            ) || 0,
            valor
          ),

        observacao:
          observacao,

        competencia:
          competencia,

        chave:
          chave,

        fonte:
          'Manual'
      });

    }
  );


  return registros;

}


/**
 * ============================================================
 * 💵 LÊ ALOCAÇÕES
 * ============================================================
 *
 * _REEMB_ALOCACOES
 *
 * A ID Pagamento
 * B Chave Reembolso
 * C Valor Alocado
 * D Data Pagamento
 * E Responsável
 * F Competência Dívida
 */

function obterAlocacoesReembolsosPorChave_(
  ss
) {

  const mapa =
    new Map();


  const aba =
    ss.getSheetByName(
      REEMBOLSOS_CONFIG.ABA_ALOCACOES
    );


  if (!aba) {
    return mapa;
  }


  const ultimaLinha =
    aba.getLastRow();


  if (
    ultimaLinha < 2
  ) {
    return mapa;
  }


  const dados =
    aba
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        6
      )
      .getValues();


  dados.forEach(
    linha => {

      const chave =
        String(
          linha[1] || ''
        ).trim();


      const valor =
        Number(
          linha[2]
        ) || 0;


      if (
        !chave ||
        valor <= 0
      ) {
        return;
      }


      mapa.set(
        chave,
        (
          mapa.get(
            chave
          ) || 0
        ) +
        valor
      );

    }
  );


  return mapa;

}


/**
 * ============================================================
 * 🧠 PENDÊNCIAS PARA FUTURO RATEIO
 * ============================================================
 *
 * Esta função será usada no próximo passo,
 * quando criarmos o registro de pagamento.
 *
 * Retorna as dívidas em ordem:
 * mais antiga -> mais nova.
 */

function obterPendenciasReembolsosParaAlocacao_(
  responsavel
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const aba =
    ss.getSheetByName(
      REEMBOLSOS_CONFIG.ABA
    );


  if (!aba) {
    return [];
  }


  const ultimaLinha =
    aba.getLastRow();


  if (
    ultimaLinha <
    REEMBOLSOS_CONFIG.PRIMEIRA_LINHA
  ) {
    return [];
  }


  const dados =
    aba
      .getRange(
        REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
        1,
        ultimaLinha -
          REEMBOLSOS_CONFIG.PRIMEIRA_LINHA +
          1,
        12
      )
      .getValues();


  const nomeNormalizado =
    normalizarTextoReembolsos_(
      responsavel
    );


  const pendencias =
    [];


  dados.forEach(
    linha => {

      const resp =
        String(
          linha[1] || ''
        ).trim();


      if (
        normalizarTextoReembolsos_(
          resp
        ) !==
        nomeNormalizado
      ) {
        return;
      }


      const chave =
        String(
          linha[10] || ''
        ).trim();


      if (!chave) {
        return;
      }


      const devido =
        Number(
          linha[4]
        ) || 0;


      const pago =
        Number(
          linha[5]
        ) || 0;


      const saldo =
        Math.max(
          devido - pago,
          0
        );


      if (
        saldo <= 0
      ) {
        return;
      }


      const data =
        linha[0] instanceof Date
          ? normalizarDataReembolsos_(
              linha[0]
            )
          : new Date(
              2099,
              0,
              1,
              12
            );


      const competencia =
        linha[9] instanceof Date
          ? primeiroDiaMesReembolsos_(
              linha[9]
            )
          : primeiroDiaMesReembolsos_(
              data
            );


      pendencias.push({

        chave:
          chave,

        data:
          data,

        competencia:
          competencia,

        responsavel:
          resp,

        descricao:
          linha[3],

        valorDevido:
          devido,

        valorPago:
          pago,

        saldo:
          saldo

      });

    }
  );


  /**
   * Mais antiga primeiro.
   */

  pendencias.sort(
    (a, b) =>
      a.data.getTime() -
      b.data.getTime()
  );


  return pendencias;

}


/**
 * ============================================================
 * 🧮 G / H
 * ============================================================
 */

function aplicarFormulasReembolsos_(
  aba
) {

  const ultimaLinha =
    aba.getLastRow();


  if (
    ultimaLinha <
    REEMBOLSOS_CONFIG.PRIMEIRA_LINHA
  ) {
    return;
  }


  const quantidade =
    ultimaLinha -
    REEMBOLSOS_CONFIG.PRIMEIRA_LINHA +
    1;


  /**
   * G = saldo
   */

  const formulasSaldo = [];


  /**
   * H = status
   */

  const formulasStatus = [];


  for (
    let i = 0;
    i < quantidade;
    i++
  ) {

    const linha =
      REEMBOLSOS_CONFIG.PRIMEIRA_LINHA +
      i;


    formulasSaldo.push([
      '=IF(E' +
        linha +
        '="";"";MAX(E' +
        linha +
        '-F' +
        linha +
        ';0))'
    ]);


    formulasStatus.push([
      '=IF(E' +
        linha +
        '="";"";' +
        'IF(G' +
        linha +
        '=0;' +
        '"✅ Pago";' +
        'IF(F' +
        linha +
        '>0;' +
        '"🟡 Parcial";' +
        '"🔴 Pendente"' +
        ')' +
        ')' +
        ')'
    ]);

  }


  aba
    .getRange(
      REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
      7,
      quantidade,
      1
    )
    .setFormulas(
      formulasSaldo
    );


  aba
    .getRange(
      REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
      8,
      quantidade,
      1
    )
    .setFormulas(
      formulasStatus
    );

}


/**
 * ============================================================
 * 👤 VALIDAÇÃO
 * ============================================================
 */

function aplicarValidacoesReembolsos_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const aba =
    ss.getSheetByName(
      REEMBOLSOS_CONFIG.ABA
    );


  if (!aba) {
    return;
  }


  const responsaveis =
    Array.from(
      obterSetResponsaveisTerceiros_(
        ss
      )
    );


  if (
    responsaveis.length === 0
  ) {
    return;
  }


  /**
   * Recuperamos a capitalização original da _CONFIG.
   */

  const config =
    ss.getSheetByName(
      REEMBOLSOS_CONFIG.CONFIG
    );


  const listaOriginal = [];


  if (
    config &&
    config.getLastRow() >= 2
  ) {

    const dados =
      config
        .getRange(
          2,
          1,
          config.getLastRow() - 1,
          2
        )
        .getDisplayValues();


    dados.forEach(
      linha => {

        if (
          String(
            linha[1] || ''
          ).trim() ===
          'Terceiro / Reembolso'
        ) {

          const nome =
            String(
              linha[0] || ''
            ).trim();


          if (nome) {

            listaOriginal.push(
              nome
            );

          }

        }

      }
    );

  }


  if (
    listaOriginal.length === 0
  ) {
    return;
  }


  const regra =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        listaOriginal,
        true
      )
      .setAllowInvalid(false)
      .build();


  aba
    .getRange(
      REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
      2,
      Math.max(
        aba.getMaxRows() -
          REEMBOLSOS_CONFIG.PRIMEIRA_LINHA +
          1,
        1
      ),
      1
    )
    .setDataValidation(
      regra
    );

}


/**
 * ============================================================
 * 📋 CABEÇALHO
 * ============================================================
 */

function prepararCabecalhoReembolsos_(
  aba
) {

  aba
    .getRange(
      REEMBOLSOS_CONFIG.LINHA_CABECALHO,
      1,
      1,
      12
    )
    .setValues([[
      'Data',
      'Responsável',
      'Origem',
      'Descrição',
      'Valor Devido',
      'Valor Pago',
      'Saldo',
      'Status',
      'Observação',
      'Competência',
      'Chave Sistema',
      'Fonte'
    ]]);

}


/**
 * ============================================================
 * 💵 FORMATOS
 * ============================================================
 */

function aplicarFormatosBasicosReembolsos_(
  aba
) {

  const ultimaLinha =
    Math.max(
      aba.getLastRow(),
      REEMBOLSOS_CONFIG.PRIMEIRA_LINHA
    );


  const quantidade =
    Math.max(
      ultimaLinha -
        REEMBOLSOS_CONFIG.PRIMEIRA_LINHA +
        1,
      1
    );


  aba
    .getRange(
      REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
      1,
      quantidade,
      1
    )
    .setNumberFormat(
      'dd/mm/yyyy'
    );


  aba
    .getRange(
      REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
      5,
      quantidade,
      3
    )
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  aba
    .getRange(
      REEMBOLSOS_CONFIG.PRIMEIRA_LINHA,
      10,
      quantidade,
      1
    )
    .setNumberFormat(
      'mm/yyyy'
    );

}


/**
 * ============================================================
 * 🧾 REGISTRO → LINHA
 * ============================================================
 */

function registroParaLinhaReembolsos_(
  registro
) {

  return [

    registro.data,

    registro.responsavel,

    registro.origem,

    registro.descricao,

    Number(
      registro.valorDevido
    ) || 0,

    Number(
      registro.valorPago
    ) || 0,

    '',

    '',

    registro.observacao || '',

    registro.competencia,

    registro.chave,

    registro.fonte

  ];

}


/**
 * ============================================================
 * 👥 TERCEIROS
 * ============================================================
 */

function obterSetResponsaveisTerceiros_(
  ss
) {

  const set =
    new Set();


  /**
   * Primeiro tenta Configuracoes.gs.
   */

  try {

    if (
      typeof obterResponsaveisTerceiros_ ===
      'function'
    ) {

      const lista =
        obterResponsaveisTerceiros_();


      if (
        Array.isArray(
          lista
        )
      ) {

        lista.forEach(
          item => {

            const nome =
              normalizarTextoReembolsos_(
                item
              );


            if (nome) {

              set.add(
                nome
              );

            }

          }
        );

      }

    }

  } catch (erro) {

    console.log(
      erro
    );

  }


  if (
    set.size > 0
  ) {
    return set;
  }


  /**
   * Fallback.
   */

  const config =
    ss.getSheetByName(
      REEMBOLSOS_CONFIG.CONFIG
    );


  if (
    !config ||
    config.getLastRow() < 2
  ) {

    return set;
  }


  const dados =
    config
      .getRange(
        2,
        1,
        config.getLastRow() - 1,
        2
      )
      .getDisplayValues();


  dados.forEach(
    linha => {

      if (
        String(
          linha[1] || ''
        ).trim() ===
        'Terceiro / Reembolso'
      ) {

        const nome =
          normalizarTextoReembolsos_(
            linha[0]
          );


        if (nome) {

          set.add(
            nome
          );

        }

      }

    }
  );


  return set;

}


/**
 * ============================================================
 * 📅 COMPETÊNCIA
 * ============================================================
 */

function obterCompetenciaAtualReembolsos_(
  ss
) {

  const dashboard =
    ss.getSheetByName(
      REEMBOLSOS_CONFIG.DASHBOARD
    );


  if (!dashboard) {
    return null;
  }


  const data =
    dashboard
      .getRange(
        REEMBOLSOS_CONFIG.CELULA_COMPETENCIA
      )
      .getValue();


  if (
    !(data instanceof Date) ||
    isNaN(
      data.getTime()
    )
  ) {
    return null;
  }


  return primeiroDiaMesReembolsos_(
    data
  );

}


/**
 * ============================================================
 * 🔑 CHAVE PARCELAMENTO
 * ============================================================
 */

function montarChaveParcelamentoReembolsos_(
  responsavel,
  produto,
  valorTotal,
  qtdParcelas,
  valorParcela,
  cartao,
  primeiraParcela,
  competencia
) {

  return (
    'PARC|' +
    [
      limparChaveReembolsos_(
        responsavel
      ),

      limparChaveReembolsos_(
        produto
      ),

      numeroChaveReembolsos_(
        valorTotal
      ),

      Number(
        qtdParcelas
      ) || 0,

      numeroChaveReembolsos_(
        valorParcela
      ),

      limparChaveReembolsos_(
        cartao
      ),

      primeiraParcela
        ? formatarDataChaveReembolsos_(
            primeiraParcela
          )
        : '',

      formatarDataChaveReembolsos_(
        competencia
      )

    ].join('|')
  );

}


/**
 * ============================================================
 * 🔑 COMPETÊNCIA DA CHAVE PARC
 * ============================================================
 */

function extrairCompetenciaDaChaveParcelamento_(
  chave
) {

  if (
    !String(
      chave || ''
    ).startsWith(
      'PARC|'
    )
  ) {
    return null;
  }


  const partes =
    String(
      chave
    ).split('|');


  const dataTexto =
    partes[
      partes.length - 1
    ];


  const match =
    String(
      dataTexto
    ).match(
      /^(\d{4})-(\d{2})-(\d{2})$/
    );


  if (!match) {
    return null;
  }


  return new Date(
    Number(
      match[1]
    ),
    Number(
      match[2]
    ) - 1,
    1,
    12,
    0,
    0,
    0
  );

}


/**
 * ============================================================
 * 🔑 ASSINATURA VARIÁVEL
 * ============================================================
 */

function montarAssinaturaDespesaVariavel_(
  data,
  responsavel,
  origem,
  descricao,
  valor,
  categoria,
  dataOriginal,
  descricaoOriginal
) {

  const original =
    dataOriginal instanceof Date
      ? normalizarDataReembolsos_(
          dataOriginal
        )
      : null;


  return [

    formatarDataChaveReembolsos_(
      data
    ),

    limparChaveReembolsos_(
      responsavel
    ),

    limparChaveReembolsos_(
      origem
    ),

    limparChaveReembolsos_(
      descricao
    ),

    numeroChaveReembolsos_(
      valor
    ),

    limparChaveReembolsos_(
      categoria
    ),

    original
      ? formatarDataChaveReembolsos_(
          original
        )
      : '',

    limparChaveReembolsos_(
      descricaoOriginal
    )

  ].join('|');

}


/**
 * ============================================================
 * 🧹 DUPLICATAS
 * ============================================================
 */

function removerDuplicatasReembolsos_(
  registros
) {

  const mapa =
    new Map();


  registros.forEach(
    registro => {

      const chave =
        String(
          registro.chave || ''
        ).trim();


      if (!chave) {
        return;
      }


      if (
        !mapa.has(
          chave
        )
      ) {

        mapa.set(
          chave,
          registro
        );

      }

    }
  );


  return Array.from(
    mapa.values()
  );

}


/**
 * ============================================================
 * 📊 ORDEM VISUAL
 * ============================================================
 */

function compararReembolsosMaisRecentes_(
  a,
  b
) {

  const dataA =
    a.data instanceof Date
      ? a.data.getTime()
      : 0;


  const dataB =
    b.data instanceof Date
      ? b.data.getTime()
      : 0;


  if (
    dataA !== dataB
  ) {

    return dataB -
      dataA;

  }


  return String(
    a.descricao || ''
  ).localeCompare(
    String(
      b.descricao || ''
    ),
    'pt-BR'
  );

}


/**
 * ============================================================
 * 📅 DATA SEGURA
 * ============================================================
 */

function normalizarDataReembolsos_(
  data
) {

  if (
    !(data instanceof Date) ||
    isNaN(
      data.getTime()
    )
  ) {
    return null;
  }


  return new Date(
    data.getFullYear(),
    data.getMonth(),
    data.getDate(),
    12,
    0,
    0,
    0
  );

}


function primeiroDiaMesReembolsos_(
  data
) {

  if (
    !(data instanceof Date) ||
    isNaN(
      data.getTime()
    )
  ) {
    return null;
  }


  return new Date(
    data.getFullYear(),
    data.getMonth(),
    1,
    12,
    0,
    0,
    0
  );

}


/**
 * ============================================================
 * 🔑 TEXTO
 * ============================================================
 */

function normalizarTextoReembolsos_(
  valor
) {

  return String(
    valor || ''
  )
    .trim()
    .toLowerCase();

}


function limparChaveReembolsos_(
  valor
) {

  return String(
    valor || ''
  )
    .trim()
    .toLowerCase()
    .replace(
      /\s+/g,
      ' '
    )
    .replace(
      /\|/g,
      '/'
    );

}


/**
 * ============================================================
 * 🔢 NÚMERO PARA CHAVE
 * ============================================================
 */

function numeroChaveReembolsos_(
  valor
) {

  return (
    Math.round(
      (
        Number(
          valor
        ) || 0
      ) *
      100
    ) /
    100
  ).toFixed(
    2
  );

}


/**
 * ============================================================
 * 📅 DATA PARA CHAVE
 * ============================================================
 */

function formatarDataChaveReembolsos_(
  data
) {

  if (
    !(data instanceof Date) ||
    isNaN(
      data.getTime()
    )
  ) {
    return '';
  }


  return (
    data.getFullYear() +
    '-' +
    String(
      data.getMonth() + 1
    ).padStart(
      2,
      '0'
    ) +
    '-' +
    String(
      data.getDate()
    ).padStart(
      2,
      '0'
    )
  );

}


/**
 * ============================================================
 * #️⃣ HASH
 * ============================================================
 */

function gerarHashReembolsos_(
  texto
) {

  let hash =
    0;


  const str =
    String(
      texto || ''
    );


  for (
    let i = 0;
    i < str.length;
    i++
  ) {

    hash =
      (
        (
          hash << 5
        ) -
        hash
      ) +
      str.charCodeAt(
        i
      );


    hash |= 0;

  }


  return (
    hash >>> 0
  ).toString(
    36
  );

}


/**
 * ============================================================
 * 💰 LIMITE DO VALOR PAGO
 * ============================================================
 */

function limitarValorPagoReembolsos_(
  pago,
  devido
) {

  return Math.min(
    Math.max(
      Number(
        pago
      ) || 0,
      0
    ),
    Math.max(
      Number(
        devido
      ) || 0,
      0
    )
  );

}


/**
 * ============================================================
 * 📅 CHAVE DO MÊS
 * ============================================================
 */

function chaveMesReembolsos_(
  data
) {

  if (
    !(data instanceof Date)
  ) {
    return '';
  }


  return (
    data.getFullYear() +
    '-' +
    String(
      data.getMonth() + 1
    ).padStart(
      2,
      '0'
    )
  );

}


/**
 * ============================================================
 * 🌎 TIMEZONE
 * ============================================================
 */

function corrigirTimezoneReembolsos_(
  ss
) {

  try {

    if (
      ss.getSpreadsheetTimeZone() !==
      REEMBOLSOS_CONFIG.TIMEZONE
    ) {

      ss.setSpreadsheetTimeZone(
        REEMBOLSOS_CONFIG.TIMEZONE
      );

    }

  } catch (erro) {

    console.log(
      'Não foi possível ajustar timezone: ' +
      erro
    );

  }

}