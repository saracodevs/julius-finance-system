  /**
   * ============================================================
   * 💵 PAGAMENTOS DE REEMBOLSOS
   * JULIUS FINANCE SYSTEM
   * ============================================================
   *
   * Responsabilidades deste arquivo:
   *
   * 1. Preparar a aba de histórico de pagamentos.
   * 2. Preparar a aba técnica de alocações.
   * 3. Abrir o modal "Registrar Pagamento".
   * 4. Listar responsáveis classificados como reembolso.
   * 5. Registrar pagamentos recebidos.
   * 6. Distribuir o valor automaticamente entre as dívidas
   *    mais antigas do responsável.
   * 7. Atualizar Valor Pago, Saldo e Status dos reembolsos.
   */

  const PAG_REEMB_CONFIG = {
    ABA_PAGAMENTOS: '💵 Pagamentos Reembolsos',
    ABA_ALOCACOES: '_REEMB_ALOCACOES',
    ABA_REEMBOLSOS: '💸 Reembolsos',
    ABA_CONFIG: '_CONFIG',
    TIMEZONE: 'America/Sao_Paulo'
  };


  /**
   * ============================================================
   * 🚀 PREPARAÇÃO DO SISTEMA
   * ============================================================
   */

  function prepararSistemaPagamentosReembolsos() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
      if (
        ss.getSpreadsheetTimeZone() !==
        PAG_REEMB_CONFIG.TIMEZONE
      ) {
        ss.setSpreadsheetTimeZone(
          PAG_REEMB_CONFIG.TIMEZONE
        );
      }
    } catch (erro) {
      console.log(
        'Timezone não alterado: ' + erro
      );
    }

    prepararAbaHistoricoPagamentos_(ss);
    prepararAbaAlocacoesReembolsos_(ss);

    SpreadsheetApp.flush();

    ss.toast(
      'Sistema de pagamentos preparado 💵',
      'Julius Finance',
      4
    );
  }


  /**
   * ============================================================
   * 🪟 ABRIR MODAL
   * ============================================================
   */

  function abrirRegistrarPagamentoReembolso() {
    const html =
      HtmlService
        .createHtmlOutputFromFile(
          'RegistrarPagamentoReembolso'
        )
        .setWidth(570)
        .setHeight(670);

    SpreadsheetApp
      .getUi()
      .showModalDialog(
        html,
        '💵 Registrar Pagamento'
      );
  }


  /**
   * ============================================================
   * 👤 LISTAR RESPONSÁVEIS
   * ============================================================
   *
   * Busca na aba _CONFIG somente quem estiver classificado como:
   *
   * Terceiro / Reembolso
   */

  function listarResponsaveisReembolso() {
    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const config =
      ss.getSheetByName(
        PAG_REEMB_CONFIG.ABA_CONFIG
      );

    if (!config) {
      throw new Error(
        'A aba _CONFIG não foi encontrada.'
      );
    }

    const ultimaLinha =
      config.getLastRow();

    if (ultimaLinha < 2) {
      return [];
    }

    const dados =
      config
        .getRange(
          2,
          1,
          ultimaLinha - 1,
          2
        )
        .getValues();

    const responsaveis =
      dados
        .filter(function (linha) {
          const nome =
            String(
              linha[0] || ''
            ).trim();

          const tipo =
            String(
              linha[1] || ''
            ).trim();

          return (
            nome &&
            tipo ===
              'Terceiro / Reembolso'
          );
        })
        .map(function (linha) {
          return String(
            linha[0]
          ).trim();
        });

    return [
      ...new Set(responsaveis)
    ].sort();
  }

  /**
   * ============================================================
   * 📅 LISTAR COMPETÊNCIAS PENDENTES DO RESPONSÁVEL
   * ============================================================
   *
   * Retorna somente competências que ainda possuem saldo.
   *
   * Exemplo:
   *
   * [
   *   { valor: '2026-09', rotulo: 'Setembro/2026' },
   *   { valor: '2026-08', rotulo: 'Agosto/2026' },
   *   { valor: '2026-07', rotulo: 'Julho/2026' }
   * ]
   */

  function listarCompetenciasPendentesReembolso(
    responsavel
  ) {

    const nome =
      String(
        responsavel || ''
      ).trim();

    if (!nome) {
      return [];
    }

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const aba =
      ss.getSheetByName(
        PAG_REEMB_CONFIG.ABA_REEMBOLSOS
      );

    if (!aba) {
      return [];
    }

    const primeiraLinha = 5;
    const ultimaLinha =
      aba.getLastRow();

    if (
      ultimaLinha <
      primeiraLinha
    ) {
      return [];
    }

    const quantidade =
      ultimaLinha -
      primeiraLinha +
      1;

    const valores =
      aba
        .getRange(
          primeiraLinha,
          1,
          quantidade,
          12
        )
        .getValues();

    const competencias = {};

    valores.forEach(
      function (linha) {

        const responsavelLinha =
          String(
            linha[1] || ''
          ).trim();

        if (
          responsavelLinha !== nome
        ) {
          return;
        }

        const valorDevido =
          Number(
            linha[4]
          ) || 0;

        const valorPago =
          Number(
            linha[5]
          ) || 0;

        const saldoPlanilha =
          Number(
            linha[6]
          );

        let saldo =
          isFinite(
            saldoPlanilha
          )
            ? saldoPlanilha
            : valorDevido -
              valorPago;

        saldo =
          Math.max(
            0,
            Number(
              saldo.toFixed(2)
            )
          );

        if (
          saldo <= 0.009
        ) {
          return;
        }

        const competencia =
          linha[9];

        if (
          !(
            competencia instanceof Date
          ) ||
          isNaN(
            competencia.getTime()
          )
        ) {
          return;
        }

        const ano =
          competencia.getFullYear();

        const mes =
          competencia.getMonth() + 1;

        const chave =
          ano +
          '-' +
          String(mes)
            .padStart(
              2,
              '0'
            );

        competencias[chave] = {
          valor: chave,
          ano: ano,
          mes: mes
        };

      }
    );

    return Object
      .values(
        competencias
      )
      .sort(
        function (a, b) {

          if (
            a.ano !== b.ano
          ) {
            return (
              b.ano -
              a.ano
            );
          }

          return (
            b.mes -
            a.mes
          );

        }
      )
      .map(
        function (item) {

          return {
            valor:
              item.valor,

            rotulo:
              formatarCompetenciaPagamentoReembolso_(
                item.ano,
                item.mes
              )
          };

        }
      );

  }


  /**
   * ============================================================
   * 📅 COMPETÊNCIA TEXTO -> DATA
   * ============================================================
   *
   * Recebe:
   *
   * 2026-08
   *
   * Retorna:
   *
   * Date referente a 01/08/2026.
   */

  function converterCompetenciaPagamentoReembolso_(
    texto
  ) {

    const partes =
      String(
        texto || ''
      )
        .trim()
        .split('-');

    if (
      partes.length !== 2
    ) {
      throw new Error(
        'Competência inválida.'
      );
    }

    const ano =
      Number(
        partes[0]
      );

    const mes =
      Number(
        partes[1]
      );

    if (
      !ano ||
      mes < 1 ||
      mes > 12
    ) {
      throw new Error(
        'Competência inválida.'
      );
    }

    return new Date(
      ano,
      mes - 1,
      1,
      12,
      0,
      0
    );

  }


  /**
   * ============================================================
   * 📅 COMPARAR COMPETÊNCIAS
   * ============================================================
   */

  function mesmaCompetenciaPagamentoReembolso_(
    data,
    competencia
  ) {

    if (
      !(data instanceof Date) ||
      isNaN(
        data.getTime()
      )
    ) {
      return false;
    }

    if (
      !(
        competencia instanceof Date
      ) ||
      isNaN(
        competencia.getTime()
      )
    ) {
      return false;
    }

    return (
      data.getFullYear() ===
        competencia.getFullYear() &&
      data.getMonth() ===
        competencia.getMonth()
    );

  }


  /**
   * ============================================================
   * 📅 FORMATAR COMPETÊNCIA
   * ============================================================
   */

  function formatarCompetenciaPagamentoReembolso_(
    ano,
    mes
  ) {

    const meses = [
      'Janeiro',
      'Fevereiro',
      'Março',
      'Abril',
      'Maio',
      'Junho',
      'Julho',
      'Agosto',
      'Setembro',
      'Outubro',
      'Novembro',
      'Dezembro'
    ];

    return (
      meses[mes - 1] +
      '/' +
      ano
    );

  }

  /**
   * ============================================================
   * 💰 REGISTRAR PAGAMENTO
   * ============================================================
   */

  function registrarPagamentoReembolso(dados) {

    const lock =
      LockService.getDocumentLock();

    lock.waitLock(
      30000
    );

    try {

      validarDadosPagamentoReembolso_(
        dados
      );

      const ss =
        SpreadsheetApp
          .getActiveSpreadsheet();

      prepararAbaHistoricoPagamentos_(
        ss
      );

      prepararAbaAlocacoesReembolsos_(
        ss
      );

      SpreadsheetApp.flush();


      const dataPagamento =
        converterDataPagamentoReembolso_(
          dados.data
        );


      const responsavel =
        String(
          dados.responsavel
        ).trim();


      const valorRecebido =
        Number(
          dados.valor
        );


      const forma =
        String(
          dados.forma
        ).trim();


      const observacao =
        String(
          dados.observacao || ''
        ).trim();


      /**
       * ========================================================
       * COMPETÊNCIA ESCOLHIDA
       * ========================================================
       */

      const competenciaDestino =
        converterCompetenciaPagamentoReembolso_(
          dados.competencia
        );


      /**
       * ========================================================
       * BUSCA SOMENTE DÍVIDAS DA COMPETÊNCIA ESCOLHIDA
       * ========================================================
       */

      const dividas =
        obterDividasPendentesReembolso_(
          ss,
          responsavel,
          competenciaDestino
        );


      if (
        !dividas.length
      ) {

        throw new Error(
          'Não existem reembolsos pendentes para ' +
          responsavel +
          ' em ' +
          formatarCompetenciaPagamentoReembolso_(
            competenciaDestino.getFullYear(),
            competenciaDestino.getMonth() + 1
          ) +
          '.'
        );

      }


      /**
       * ========================================================
       * SALDO DA COMPETÊNCIA
       * ========================================================
       */

      const totalPendente =
        dividas.reduce(
          function (
            total,
            divida
          ) {

            return (
              total +
              divida.saldo
            );

          },
          0
        );


      if (
        valorRecebido >
        totalPendente + 0.009
      ) {

        throw new Error(
          'O valor informado é maior que o saldo pendente da competência ' +
          formatarCompetenciaPagamentoReembolso_(
            competenciaDestino.getFullYear(),
            competenciaDestino.getMonth() + 1
          ) +
          '. Saldo disponível: ' +
          formatarMoedaReembolso_(
            totalPendente
          )
        );

      }


      const idPagamento =
        gerarIdPagamentoReembolso_();


      /**
       * ========================================================
       * HISTÓRICO
       * ========================================================
       */

      registrarHistoricoPagamentoReembolso_(
        ss,
        {
          idPagamento:
            idPagamento,

          dataPagamento:
            dataPagamento,

          responsavel:
            responsavel,

          valorRecebido:
            valorRecebido,

          forma:
            forma,

          observacao:
            observacao,

          competenciaDestino:
            competenciaDestino
        }
      );


      /**
       * ========================================================
       * DISTRIBUIÇÃO
       * ========================================================
       *
       * Agora "dividas" contém somente a competência escolhida.
       */

      const resultadoAlocacao =
        distribuirPagamentoReembolso_(
          ss,
          {
            idPagamento:
              idPagamento,

            dataPagamento:
              dataPagamento,

            responsavel:
              responsavel,

            valorRecebido:
              valorRecebido,

            dividas:
              dividas
          }
        );


      SpreadsheetApp.flush();


      /**
       * ========================================================
       * SINCRONIZA BASE
       * ========================================================
       */

      tentarSincronizarSistemaReembolsos_();

      SpreadsheetApp.flush();


      /**
       * ========================================================
       * REAPLICA ALOCAÇÕES
       * ========================================================
       */

      atualizarReembolsosPelasAlocacoes_(
        ss,
        responsavel
      );

      SpreadsheetApp.flush();


      /**
       * ========================================================
       * REAPLICA VISÃO ATUAL
       * ========================================================
       */

      try {

        if (
          typeof aplicarVisaoCompetenciaAtualReembolsos ===
          'function'
        ) {

          aplicarVisaoCompetenciaAtualReembolsos();

        }

      } catch (
        erroFiltro
      ) {

        console.log(
          'Filtro de competência não reaplicado: ' +
          erroFiltro
        );

      }


      SpreadsheetApp.flush();


      return {

        sucesso:
          true,

        idPagamento:
          idPagamento,

        valor:
          valorRecebido,

        competencia:
          formatarCompetenciaPagamentoReembolso_(
            competenciaDestino.getFullYear(),
            competenciaDestino.getMonth() + 1
          ),

        alocacoes:
          resultadoAlocacao.quantidade,

        mensagem:
          'Pagamento de ' +
          formatarMoedaReembolso_(
            valorRecebido
          ) +
          ' destinado a ' +
          formatarCompetenciaPagamentoReembolso_(
            competenciaDestino.getFullYear(),
            competenciaDestino.getMonth() + 1
          ) +
          ' registrado com sucesso.'

      };


    } finally {

      lock.releaseLock();

    }

  }


  /**
   * ============================================================
   * ✅ VALIDAÇÃO
   * ============================================================
   */

  function validarDadosPagamentoReembolso_(
    dados
  ) {

    if (!dados) {

      throw new Error(
        'Nenhum dado de pagamento foi informado.'
      );

    }


    if (!dados.data) {

      throw new Error(
        'Informe a data do recebimento.'
      );

    }


    if (
      !String(
        dados.responsavel || ''
      ).trim()
    ) {

      throw new Error(
        'Selecione o responsável.'
      );

    }


    if (
      !String(
        dados.competencia || ''
      ).trim()
    ) {

      throw new Error(
        'Selecione a competência que deseja pagar.'
      );

    }


    const valor =
      Number(
        dados.valor
      );


    if (
      !isFinite(valor) ||
      valor <= 0
    ) {

      throw new Error(
        'Informe um valor recebido maior que zero.'
      );

    }


    if (
      !String(
        dados.forma || ''
      ).trim()
    ) {

      throw new Error(
        'Selecione a forma de pagamento.'
      );

    }

  }

  /**
   * ============================================================
   * 📅 CONVERTER DATA
   * ============================================================
   *
   * Recebe:
   *
   * 2026-09-13
   *
   * e cria a data usando meio-dia para evitar alteração de dia
   * causada por timezone.
   */

  function converterDataPagamentoReembolso_(texto) {
    const partes =
      String(
        texto
      ).split('-');

    if (
      partes.length !== 3
    ) {
      throw new Error(
        'Data do pagamento inválida.'
      );
    }

    const ano =
      Number(
        partes[0]
      );

    const mes =
      Number(
        partes[1]
      );

    const dia =
      Number(
        partes[2]
      );

    const data =
      new Date(
        ano,
        mes - 1,
        dia,
        12,
        0,
        0
      );

    if (
      isNaN(
        data.getTime()
      )
    ) {
      throw new Error(
        'Data do pagamento inválida.'
      );
    }

    return data;
  }


  /**
   * ============================================================
   * 🔎 BUSCAR DÍVIDAS PENDENTES
   * ============================================================
   *
   * Colunas de 💸 Reembolsos:
   *
   * A Data
   * B Responsável
   * C Origem
   * D Descrição
   * E Valor Devido
   * F Valor Pago
   * G Saldo
   * H Status
   * I Observação
   * J Competência
   * K Chave Sistema
   * L Fonte
   */

  function obterDividasPendentesReembolso_(
    ss,
    responsavel,
    competenciaDestino
  ) {

    const aba =
      ss.getSheetByName(
        PAG_REEMB_CONFIG.ABA_REEMBOLSOS
      );


    if (!aba) {

      throw new Error(
        'A aba 💸 Reembolsos não foi encontrada.'
      );

    }


    SpreadsheetApp.flush();


    const primeiraLinha =
      5;


    const ultimaLinha =
      aba.getLastRow();


    if (
      ultimaLinha <
      primeiraLinha
    ) {

      return [];

    }


    const quantidadeLinhas =
      ultimaLinha -
      primeiraLinha +
      1;


    const valores =
      aba
        .getRange(
          primeiraLinha,
          1,
          quantidadeLinhas,
          12
        )
        .getValues();


    const dividas = [];


    valores.forEach(
      function (
        linha,
        indice
      ) {

        const responsavelLinha =
          String(
            linha[1] || ''
          ).trim();


        if (
          responsavelLinha !==
          responsavel
        ) {

          return;

        }


        const competenciaLinha =
          linha[9];


        /**
         * ======================================================
         * FILTRO DA COMPETÊNCIA
         * ======================================================
         */

        if (
          competenciaDestino &&
          !mesmaCompetenciaPagamentoReembolso_(
            competenciaLinha,
            competenciaDestino
          )
        ) {

          return;

        }


        const valorDevido =
          Number(
            linha[4]
          ) || 0;


        const valorPago =
          Number(
            linha[5]
          ) || 0;


        const saldoPlanilha =
          Number(
            linha[6]
          );


        const chave =
          String(
            linha[10] || ''
          ).trim();


        if (!chave) {

          return;

        }


        let saldo;


        if (
          isFinite(
            saldoPlanilha
          )
        ) {

          saldo =
            saldoPlanilha;

        } else {

          saldo =
            valorDevido -
            valorPago;

        }


        saldo =
          Math.max(
            0,
            Number(
              saldo.toFixed(2)
            )
          );


        if (
          saldo <= 0.009
        ) {

          return;

        }


        let dataDivida =
          linha[0];


        if (
          !(
            dataDivida instanceof Date
          ) ||
          isNaN(
            dataDivida.getTime()
          )
        ) {

          dataDivida =
            competenciaLinha;

        }


        if (
          !(
            dataDivida instanceof Date
          ) ||
          isNaN(
            dataDivida.getTime()
          )
        ) {

          dataDivida =
            new Date(
              9999,
              11,
              31
            );

        }


        dividas.push({

          linha:
            primeiraLinha +
            indice,

          chave:
            chave,

          data:
            dataDivida,

          competencia:
            competenciaLinha,

          valorDevido:
            valorDevido,

          valorPago:
            valorPago,

          saldo:
            saldo

        });

      }
    );


    /**
     * ==========================================================
     * MAIS ANTIGA PRIMEIRO,
     * MAS SOMENTE DENTRO DA COMPETÊNCIA ESCOLHIDA
     * ==========================================================
     */

    dividas.sort(
      function (
        a,
        b
      ) {

        const diferencaData =
          a.data.getTime() -
          b.data.getTime();


        if (
          diferencaData !== 0
        ) {

          return diferencaData;

        }


        return (
          a.linha -
          b.linha
        );

      }
    );


    return dividas;

  }

  /**
   * ============================================================
   * 🧾 REGISTRAR HISTÓRICO
   * ============================================================
   */

  function registrarHistoricoPagamentoReembolso_(
    ss,
    pagamento
  ) {

    const aba =
      ss.getSheetByName(
        PAG_REEMB_CONFIG.ABA_PAGAMENTOS
      );


    if (!aba) {

      throw new Error(
        'A aba de pagamentos não foi encontrada.'
      );

    }


    aba.appendRow([
      pagamento.idPagamento,
      pagamento.dataPagamento,
      pagamento.responsavel,
      pagamento.valorRecebido,
      pagamento.forma,
      pagamento.observacao,
      new Date(),
      pagamento.competenciaDestino || ''
    ]);

  }


  /**
   * ============================================================
   * 🔀 DISTRIBUIR PAGAMENTO
   * ============================================================
   *
   * Regra:
   *
   * O pagamento é usado primeiro na dívida mais antiga.
   *
   * Exemplo:
   *
   * Dívida 1 = 100
   * Dívida 2 = 80
   *
   * Pagamento recebido = 130
   *
   * Resultado:
   *
   * Dívida 1 recebe 100
   * Dívida 2 recebe 30
   */

  function distribuirPagamentoReembolso_(
    ss,
    dados
  ) {
    const aba =
      ss.getSheetByName(
        PAG_REEMB_CONFIG.ABA_ALOCACOES
      );

    if (!aba) {
      throw new Error(
        'A aba técnica de alocações não foi encontrada.'
      );
    }

    let restante =
      Number(
        dados.valorRecebido
      );

    const linhas = [];

    dados.dividas.forEach(
      function (divida) {
        if (
          restante <= 0.009
        ) {
          return;
        }

        const valorAlocar =
          Math.min(
            restante,
            divida.saldo
          );

        const valorFinal =
          Number(
            valorAlocar.toFixed(2)
          );

        if (
          valorFinal <= 0
        ) {
          return;
        }

        linhas.push([
          dados.idPagamento,
          divida.chave,
          valorFinal,
          dados.dataPagamento,
          dados.responsavel,
          divida.competencia || ''
        ]);

        restante =
          Number(
            (
              restante -
              valorFinal
            ).toFixed(2)
          );
      }
    );

    if (
      restante > 0.009
    ) {
      throw new Error(
        'Não foi possível distribuir todo o pagamento.'
      );
    }

    if (!linhas.length) {
      throw new Error(
        'Nenhuma dívida recebeu o pagamento.'
      );
    }

    const primeiraLinha =
      aba.getLastRow() + 1;

    aba
      .getRange(
        primeiraLinha,
        1,
        linhas.length,
        6
      )
      .setValues(
        linhas
      );

    return {
      quantidade:
        linhas.length
    };
  }


  /**
   * ============================================================
   * 🔄 ATUALIZAR REEMBOLSOS
   * ============================================================
   *
   * Soma todas as alocações existentes para cada chave
   * e atualiza:
   *
   * F Valor Pago
   * G Saldo
   * H Status
   */

  function atualizarReembolsosPelasAlocacoes_(
    ss,
    responsavel
  ) {
    const abaReembolsos =
      ss.getSheetByName(
        PAG_REEMB_CONFIG.ABA_REEMBOLSOS
      );

    const abaAlocacoes =
      ss.getSheetByName(
        PAG_REEMB_CONFIG.ABA_ALOCACOES
      );

    if (
      !abaReembolsos ||
      !abaAlocacoes
    ) {
      return;
    }

    const pagamentosPorChave =
      {};

    const ultimaAlocacao =
      abaAlocacoes.getLastRow();

    if (
      ultimaAlocacao >= 2
    ) {
      const alocacoes =
        abaAlocacoes
          .getRange(
            2,
            1,
            ultimaAlocacao - 1,
            6
          )
          .getValues();

      alocacoes.forEach(
        function (linha) {
          const chave =
            String(
              linha[1] || ''
            ).trim();

          const valor =
            Number(
              linha[2]
            ) || 0;

          if (!chave) {
            return;
          }

          pagamentosPorChave[chave] =
            (
              pagamentosPorChave[chave] ||
              0
            ) +
            valor;
        }
      );
    }

    const ultimaLinha =
      abaReembolsos.getLastRow();

    if (
      ultimaLinha < 2
    ) {
      return;
    }

    const intervalo =
      abaReembolsos
        .getRange(
          2,
          1,
          ultimaLinha - 1,
          12
        );

    const valores =
      intervalo.getValues();

    const saidaPagamento = [];
    const saidaSaldo = [];
    const saidaStatus = [];

    valores.forEach(
      function (linha) {
        const responsavelLinha =
          String(
            linha[1] || ''
          ).trim();

        const valorDevido =
          Number(
            linha[4]
          ) || 0;

        const chave =
          String(
            linha[10] || ''
          ).trim();

        if (
          responsavel &&
          responsavelLinha !==
            responsavel
        ) {
          saidaPagamento.push([
            linha[5]
          ]);

          saidaSaldo.push([
            linha[6]
          ]);

          saidaStatus.push([
            linha[7]
          ]);

          return;
        }

        if (!chave) {
          saidaPagamento.push([
            linha[5]
          ]);

          saidaSaldo.push([
            linha[6]
          ]);

          saidaStatus.push([
            linha[7]
          ]);

          return;
        }

        const pago =
          Number(
            (
              pagamentosPorChave[chave] ||
              0
            ).toFixed(2)
          );

        const saldo =
          Number(
            Math.max(
              0,
              valorDevido -
              pago
            ).toFixed(2)
          );

        let status =
          'Pendente';

        if (
          saldo <= 0.009
        ) {
          status =
            'Pago';
        } else if (
          pago > 0
        ) {
          status =
            'Parcial';
        }

        saidaPagamento.push([
          pago
        ]);

        saidaSaldo.push([
          saldo
        ]);

        saidaStatus.push([
          status
        ]);
      }
    );

    abaReembolsos
      .getRange(
        2,
        6,
        saidaPagamento.length,
        1
      )
      .setValues(
        saidaPagamento
      );

    abaReembolsos
      .getRange(
        2,
        7,
        saidaSaldo.length,
        1
      )
      .setValues(
        saidaSaldo
      );

    abaReembolsos
      .getRange(
        2,
        8,
        saidaStatus.length,
        1
      )
      .setValues(
        saidaStatus
      );
  }


  /**
   * ============================================================
   * 🔄 SINCRONIZAÇÃO OPCIONAL
   * ============================================================
   *
   * Se a função principal do sistema de reembolsos existir,
   * usamos ela também.
   */

  function tentarSincronizarSistemaReembolsos_() {
    try {
      if (
        typeof sincronizarReembolsos ===
        'function'
      ) {
        sincronizarReembolsos();
      }
    } catch (erro) {
      console.log(
        'Sincronização complementar não executada: ' +
        erro
      );
    }
  }


  /**
   * ============================================================
   * 🆔 GERAR ID
   * ============================================================
   */

  function gerarIdPagamentoReembolso_() {
    const agora =
      new Date();

    const data =
      Utilities.formatDate(
        agora,
        PAG_REEMB_CONFIG.TIMEZONE,
        'yyyyMMdd-HHmmss'
      );

    const aleatorio =
      Utilities
        .getUuid()
        .substring(
          0,
          6
        )
        .toUpperCase();

    return (
      'PAG-' +
      data +
      '-' +
      aleatorio
    );
  }


  /**
   * ============================================================
   * 💲 FORMATAR MOEDA
   * ============================================================
   */

  function formatarMoedaReembolso_(valor) {
    return (
      'R$ ' +
      Number(
        valor || 0
      )
        .toFixed(2)
        .replace(
          '.',
          ','
        )
    );
  }


  /**
   * ============================================================
   * 💵 HISTÓRICO DE PAGAMENTOS
   * ============================================================
   */

  function prepararAbaHistoricoPagamentos_(ss) {

  let aba =
    ss.getSheetByName(
      PAG_REEMB_CONFIG.ABA_PAGAMENTOS
    );


  if (!aba) {

    aba =
      ss.insertSheet(
        PAG_REEMB_CONFIG.ABA_PAGAMENTOS
      );

  }


  /**
   * ==========================================================
   * 🔄 MIGRAÇÃO SEGURA DA ESTRUTURA ANTIGA
   * ==========================================================
   *
   * Estrutura antiga:
   *
   * Linha 1 = cabeçalho
   * Linha 2+ = pagamentos
   *
   * Estrutura nova:
   *
   * Linha 1   = título
   * Linhas 2-3 = cards
   * Linha 4   = cabeçalho
   * Linha 5+  = pagamentos
   *
   * A migração só acontece se A1 ainda contiver
   * "ID Pagamento".
   */

  const valorA1 =
    String(
      aba
        .getRange('A1')
        .getDisplayValue() || ''
    ).trim();


  if (
    valorA1 === 'ID Pagamento'
  ) {

    const ultimaLinhaAntiga =
      aba.getLastRow();


    let historicoAntigo = [];


    if (
      ultimaLinhaAntiga >= 2
    ) {

      historicoAntigo =
        aba
          .getRange(
            2,
            1,
            ultimaLinhaAntiga - 1,
            8
          )
          .getValues();

    }


    /**
     * Remove somente o conteúdo antigo de A:H.
     * Os dados já estão guardados em memória.
     */

    aba
      .getRange(
        1,
        1,
        Math.max(
          ultimaLinhaAntiga,
          4
        ),
        8
      )
      .clearContent();


    /**
     * Regrava os pagamentos a partir da linha 5.
     */

    if (
      historicoAntigo.length > 0
    ) {

      aba
        .getRange(
          5,
          1,
          historicoAntigo.length,
          8
        )
        .setValues(
          historicoAntigo
        );

    }

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

    azul: '#286173',
    azulEscuro: '#194754',

    verde: '#4F7A3B',
    verdeClaro: '#DDE8D4',

    caramelo: '#C78342',
    amareloClaro: '#F3E7BD',

    marrom: '#654229',
    texto: '#33291F',

    branco: '#FFF9F0',
    borda: '#C3A989'

  };


  /**
   * ==========================================================
   * 🔝 TÍTULO
   * ==========================================================
   */

  aba
    .getRange('A1:H1')
    .getMergedRanges()
    .forEach(
      function (range) {

        range.breakApart();

      }
    );


  aba
    .getRange('A1:H1')
    .merge()
    .setValue(
      '💵 HISTÓRICO DE PAGAMENTOS'
    )
    .setBackground(
      C.azul
    )
    .setFontColor(
      C.branco
    )
    .setFontWeight(
      'bold'
    )
    .setFontSize(
      20
    )
    .setHorizontalAlignment(
      'center'
    )
    .setVerticalAlignment(
      'middle'
    );


  /**
   * ==========================================================
   * 🧹 ÁREA DOS CARDS
   * ==========================================================
   */

  aba
    .getRange('A2:H3')
    .getMergedRanges()
    .forEach(
      function (range) {

        range.breakApart();

      }
    );


  aba
    .getRange('A2:H3')
    .clearContent()
    .setBackground(
      C.fundo
    );


  /**
   * ==========================================================
   * 💰 TOTAL RECEBIDO
   * ==========================================================
   */

  aba
    .getRange('A2:C3')
    .merge()
    .setBackground(
      C.verdeClaro
    )
    .setFormula(
      '="💰 TOTAL RECEBIDO"&CHAR(10)&' +
      'TEXT(IFERROR(SUM(D5:D);0);"R$ #,##0.00")'
    )
    .setFontColor(
      C.verde
    )
    .setFontWeight(
      'bold'
    )
    .setFontSize(
      11
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
   * 🧾 PAGAMENTOS REGISTRADOS
   * ==========================================================
   */

  aba
    .getRange('D2:E3')
    .merge()
    .setBackground(
      C.card
    )
    .setFormula(
      '="🧾 PAGAMENTOS"&CHAR(10)&' +
      'COUNTA(A5:A)&" registros"'
    )
    .setFontColor(
      C.azul
    )
    .setFontWeight(
      'bold'
    )
    .setFontSize(
      11
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
   * 📅 ÚLTIMA COMPETÊNCIA PAGA
   * ==========================================================
   */

  aba
    .getRange('F2:H3')
    .merge()
    .setBackground(
      C.amareloClaro
    )
    .setFormula(
      '="📅 ÚLTIMA COMPETÊNCIA"&CHAR(10)&' +
      'IFERROR(TEXT(MAX(H5:H);"mmmm/yyyy");"—")'
    )
    .setFontColor(
      C.marrom
    )
    .setFontWeight(
      'bold'
    )
    .setFontSize(
      11
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
   * 📋 CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A4:H4')
    .setValues([[
      '🆔 ID Pagamento',
      '📅 Data Recebimento',
      '👤 Responsável',
      '💰 Valor Recebido',
      '💳 Forma de Pagamento',
      '📝 Observação',
      '🕒 Registrado em',
      '🎯 Competência Destino'
    ]]);


  aba
    .getRange('A4:H4')
    .setBackground(
      C.azul
    )
    .setFontColor(
      C.branco
    )
    .setFontWeight(
      'bold'
    )
    .setFontSize(
      9
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
   * 🧱 BASE
   * ==========================================================
   */

  aba.setHiddenGridlines(
    true
  );


  aba.setFrozenRows(
    4
  );


  aba
    .getRange(
      1,
      1,
      aba.getMaxRows(),
      8
    )
    .setFontFamily(
      'Arial'
    )
    .setVerticalAlignment(
      'middle'
    );


  /**
   * ==========================================================
   * 📐 DIMENSÕES
   * ==========================================================
   */

  aba.setRowHeight(
    1,
    48
  );

  aba.setRowHeight(
    2,
    34
  );

  aba.setRowHeight(
    3,
    34
  );

  aba.setRowHeight(
    4,
    40
  );


  aba.setColumnWidth(
    1,
    205
  );

  aba.setColumnWidth(
    2,
    130
  );

  aba.setColumnWidth(
    3,
    145
  );

  aba.setColumnWidth(
    4,
    140
  );

  aba.setColumnWidth(
    5,
    165
  );

  aba.setColumnWidth(
    6,
    260
  );

  aba.setColumnWidth(
    7,
    170
  );

  aba.setColumnWidth(
    8,
    165
  );


  /**
   * ==========================================================
   * 💵 CORPO
   * ==========================================================
   */

  const quantidadeCorpo =
    Math.max(
      aba.getMaxRows() - 4,
      1
    );


  aba
    .getRange(
      5,
      1,
      quantidadeCorpo,
      8
    )
    .setBackground(
      C.card
    )
    .setFontColor(
      C.texto
    )
    .setFontSize(
      9
    );


  /**
   * ==========================================================
   * 💚 VALOR RECEBIDO
   * ==========================================================
   */

  aba
    .getRange(
      5,
      4,
      quantidadeCorpo,
      1
    )
    .setFontColor(
      C.verde
    )
    .setFontWeight(
      'bold'
    );


  /**
   * ==========================================================
   * 🎯 COMPETÊNCIA DESTINO
   * ==========================================================
   */

  aba
    .getRange(
      5,
      8,
      quantidadeCorpo,
      1
    )
    .setBackground(
      C.amareloClaro
    )
    .setFontColor(
      C.marrom
    )
    .setFontWeight(
      'bold'
    );


  /**
   * ==========================================================
   * 🔢 FORMATOS
   * ==========================================================
   */

  aba
    .getRange(
      'B5:B' +
      aba.getMaxRows()
    )
    .setNumberFormat(
      'dd/mm/yyyy'
    );


  aba
    .getRange(
      'D5:D' +
      aba.getMaxRows()
    )
    .setNumberFormat(
      '"R$ " #,##0.00;-"R$ " #,##0.00'
    );


  aba
    .getRange(
      'G5:G' +
      aba.getMaxRows()
    )
    .setNumberFormat(
      'dd/mm/yyyy hh:mm'
    );


  aba
    .getRange(
      'H5:H' +
      aba.getMaxRows()
    )
    .setNumberFormat(
      'mm/yyyy'
    );


  /**
   * ==========================================================
   * 🔲 BORDAS DOS CARDS
   * ==========================================================
   */

  [
    'A2:C3',
    'D2:E3',
    'F2:H3'
  ].forEach(
    function (intervalo) {

      aba
        .getRange(intervalo)
        .setBorder(
          true,
          true,
          true,
          true,
          false,
          false,
          C.borda,
          SpreadsheetApp
            .BorderStyle
            .SOLID
        );

    }
  );


  /**
   * ==========================================================
   * 🔲 BORDA DO CABEÇALHO
   * ==========================================================
   */

  aba
    .getRange('A4:H4')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      C.borda,
      SpreadsheetApp
        .BorderStyle
        .SOLID
    );

}


  /**
   * ============================================================
   * ⚙️ ALOCAÇÕES TÉCNICAS
   * ============================================================
   */

  function prepararAbaAlocacoesReembolsos_(ss) {
    let aba =
      ss.getSheetByName(
        PAG_REEMB_CONFIG.ABA_ALOCACOES
      );

    if (!aba) {
      aba =
        ss.insertSheet(
          PAG_REEMB_CONFIG.ABA_ALOCACOES
        );
    }

    aba
      .getRange('A1:F1')
      .setValues([[
        'ID Pagamento',
        'Chave Reembolso',
        'Valor Alocado',
        'Data Pagamento',
        'Responsável',
        'Competência Dívida'
      ]]);

    aba
      .getRange('A1:F1')
      .setFontWeight(
        'bold'
      );

    aba
      .getRange(
        'C2:C' +
        aba.getMaxRows()
      )
      .setNumberFormat(
        '"R$ " #,##0.00'
      );

    aba
      .getRange(
        'D2:D' +
        aba.getMaxRows()
      )
      .setNumberFormat(
        'dd/mm/yyyy'
      );

    aba
      .getRange(
        'F2:F' +
        aba.getMaxRows()
      )
      .setNumberFormat(
        'mm/yyyy'
      );

    if (
      !aba.isSheetHidden()
    ) {
      aba.hideSheet();
    }
  }