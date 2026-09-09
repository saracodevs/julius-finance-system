/**
 * ============================================================
 * ⚙️ CONFIGURAÇÕES DO SISTEMA
 * ============================================================
 *
 * Responsável por:
 *
 * - Responsáveis
 * - Tipo dos responsáveis
 * - Categorias
 * - Cartões
 * - Formas de pagamento
 * - Validações / dropdowns
 * - Sincronização da aba 💳 Cartões
 * - Sincronização da aba 💸 Reembolsos
 *
 * IMPORTANTE:
 * - NÃO possui onEdit().
 * - O único onEdit do projeto deve continuar em Automacoes.gs.
 */


const SISTEMA_CONFIG = {

  ABA_CONFIG: '_CONFIG',

  // ----------------------------------------------------------
  // RESPONSÁVEIS
  // ----------------------------------------------------------
  COL_RESPONSAVEL: 1,       // A
  COL_TIPO_RESPONSAVEL: 2,  // B

  // ----------------------------------------------------------
  // CATEGORIAS
  // ----------------------------------------------------------
  COL_CATEGORIA: 4,         // D

  // ----------------------------------------------------------
  // CARTÕES
  // ----------------------------------------------------------
  COL_CARTAO: 6,            // F

  // ----------------------------------------------------------
  // FORMAS DE PAGAMENTO
  // ----------------------------------------------------------
  COL_FORMA_PAGAMENTO: 8,   // H

  PRIMEIRA_LINHA_DADOS: 2
};


/**
 * ============================================================
 * ABRIR CONFIGURAÇÕES
 * ============================================================
 */
function abrirConfiguracoes() {

  prepararConfiguracoes_();

  const html =
    HtmlService
      .createHtmlOutputFromFile('ConfiguracoesPainel')
      .setWidth(900)
      .setHeight(650);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      '⚙️ Configurações'
    );
}


/**
 * ============================================================
 * PREPARAR CONFIGURAÇÕES
 * ============================================================
 *
 * Garante que a aba _CONFIG exista e tenha os cabeçalhos.
 */
function prepararConfiguracoes_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let aba =
    ss.getSheetByName(SISTEMA_CONFIG.ABA_CONFIG);


  if (!aba) {

    aba =
      ss.insertSheet(SISTEMA_CONFIG.ABA_CONFIG);

    aba.hideSheet();
  }


  /**
   * Cabeçalhos
   */
  aba.getRange('A1').setValue('RESPONSAVEL');
  aba.getRange('B1').setValue('TIPO_RESPONSAVEL');

  aba.getRange('D1').setValue('CATEGORIA');

  aba.getRange('F1').setValue('CARTAO');

  aba.getRange('H1').setValue('FORMA_PAGAMENTO');


  /**
   * Faz migração inicial caso a configuração
   * ainda esteja vazia.
   */
  migrarConfiguracoesExistentes_(aba);


  return aba;
}


/**
 * ============================================================
 * MIGRAÇÃO INICIAL
 * ============================================================
 *
 * Só adiciona dados padrão quando a área correspondente
 * estiver completamente vazia.
 *
 * Não apaga configurações existentes.
 */
function migrarConfiguracoesExistentes_(aba) {

  /**
   * ----------------------------------------------------------
   * RESPONSÁVEIS
   * ----------------------------------------------------------
   */

  if (!colunaConfigPossuiDados_(aba, 1)) {

    aba
      .getRange('A2:B3')
      .setValues([
        ['Sara & Viih', 'Principal'],
        ['Renata', 'Terceiro / Reembolso']
      ]);
  }


  /**
   * ----------------------------------------------------------
   * CATEGORIAS
   * ----------------------------------------------------------
   */

  if (!colunaConfigPossuiDados_(aba, 4)) {

    const categorias = [
      'Alimentação',
      'Farmácia',
      'Lazer',
      'Transporte',
      'Compras',
      'Saúde',
      'Vestuário',
      'Educação',
      'Cuidados Pessoais'
    ];

    gravarListaConfig_(
      aba,
      SISTEMA_CONFIG.COL_CATEGORIA,
      categorias
    );
  }


  /**
   * ----------------------------------------------------------
   * CARTÕES
   * ----------------------------------------------------------
   */

  if (!colunaConfigPossuiDados_(aba, 6)) {

    const cartoes = [
      'Nubank',
      'Inter',
      'Rico',
      'Havan',
      'Mercado Pago',
      'Santander',
      'PicPay',
      'Caju'
    ];

    gravarListaConfig_(
      aba,
      SISTEMA_CONFIG.COL_CARTAO,
      cartoes
    );
  }


  /**
   * ----------------------------------------------------------
   * FORMAS DE PAGAMENTO
   * ----------------------------------------------------------
   */

  if (!colunaConfigPossuiDados_(aba, 8)) {

    const formas = [
      'Dinheiro',
      'Débito',
      'Crédito',
      'PIX',
      'Outros'
    ];

    gravarListaConfig_(
      aba,
      SISTEMA_CONFIG.COL_FORMA_PAGAMENTO,
      formas
    );
  }
}


/**
 * ============================================================
 * OBTER TODAS AS CONFIGURAÇÕES
 * ============================================================
 *
 * Retorno utilizado pelo HTML ConfiguracoesPainel.
 */
function obterConfiguracoesSistema() {

  const aba =
    prepararConfiguracoes_();


  return {

    responsaveis:
      obterResponsaveisConfig_(aba),

    categorias:
      obterListaConfig_(
        aba,
        SISTEMA_CONFIG.COL_CATEGORIA
      ),

    cartoes:
      obterListaConfig_(
        aba,
        SISTEMA_CONFIG.COL_CARTAO
      ),

    formas:
      obterListaConfig_(
        aba,
        SISTEMA_CONFIG.COL_FORMA_PAGAMENTO
      ),

    // compatibilidade caso algum HTML utilize outro nome
    formasPagamento:
      obterListaConfig_(
        aba,
        SISTEMA_CONFIG.COL_FORMA_PAGAMENTO
      )
  };
}


/**
 * ============================================================
 * RESPONSÁVEIS + TIPO
 * ============================================================
 */
function obterResponsaveisConfig_(aba) {

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {
    return [];
  }


  const dados =
    aba
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        2
      )
      .getValues();


  return dados
    .filter(linha => {

      const nome =
        String(linha[0] || '').trim();

      return nome !== '';
    })
    .map(linha => {

      return {
        nome:
          String(linha[0] || '').trim(),

        tipo:
          String(linha[1] || 'Principal').trim()
      };
    });
}


/**
 * ============================================================
 * SOMENTE NOMES DOS RESPONSÁVEIS
 * ============================================================
 */
function obterResponsaveisCadastrados() {

  const aba =
    prepararConfiguracoes_();

  return obterResponsaveisConfig_(aba)
    .map(item => item.nome);
}


/**
 * ============================================================
 * RESPONSÁVEIS PRINCIPAIS
 * ============================================================
 */
function obterResponsaveisPrincipais_() {

  const aba =
    prepararConfiguracoes_();

  return obterResponsaveisConfig_(aba)
    .filter(item =>
      item.tipo === 'Principal'
    )
    .map(item =>
      item.nome
    );
}


/**
 * ============================================================
 * TERCEIROS / REEMBOLSO
 * ============================================================
 *
 * Usada também pelo Reembolsos.gs.
 */
function obterResponsaveisTerceiros_() {

  const aba =
    prepararConfiguracoes_();

  return obterResponsaveisConfig_(aba)
    .filter(item =>
      item.tipo === 'Terceiro / Reembolso'
    )
    .map(item =>
      item.nome
    );
}


/**
 * ============================================================
 * VERIFICAR SE RESPONSÁVEL É TERCEIRO
 * ============================================================
 */
function responsavelEhTerceiro_(nome) {

  if (!nome) {
    return false;
  }

  const terceiros =
    obterResponsaveisTerceiros_();

  const nomeTratado =
    String(nome).trim();


  return terceiros.some(item =>
    String(item).trim() === nomeTratado
  );
}


/**
 * ============================================================
 * SALVAR CONFIGURAÇÕES
 * ============================================================
 *
 * Essa é a função chamada pelo ConfiguracoesPainel.html.
 *
 * Ao salvar:
 *
 * 1. grava _CONFIG
 * 2. atualiza dropdowns
 * 3. sincroniza cartões
 * 4. atualiza dropdown da aba Reembolsos
 */
function salvarConfiguracoesSistema(dados) {

  if (!dados) {
    throw new Error(
      'Nenhum dado de configuração foi recebido.'
    );
  }


  const aba =
    prepararConfiguracoes_();


  /**
   * ----------------------------------------------------------
   * RESPONSÁVEIS
   * ----------------------------------------------------------
   */

  const responsaveis =
    Array.isArray(dados.responsaveis)
      ? dados.responsaveis
      : [];


  limparListaConfiguracao_(
    aba,
    SISTEMA_CONFIG.COL_RESPONSAVEL,
    2
  );


  if (responsaveis.length > 0) {

    garantirLinhasConfig_(
      aba,
      responsaveis.length + 1
    );


    const valoresResponsaveis =
      responsaveis
        .map(item => {

          /**
           * Aceita tanto:
           *
           * { nome:"Renata", tipo:"Terceiro / Reembolso" }
           *
           * quanto string simples.
           */
          if (
            typeof item === 'object' &&
            item !== null
          ) {

            return [
              String(item.nome || '').trim(),
              String(item.tipo || 'Principal').trim()
            ];
          }


          return [
            String(item || '').trim(),
            'Principal'
          ];
        })
        .filter(item =>
          item[0] !== ''
        );


    if (valoresResponsaveis.length > 0) {

      aba
        .getRange(
          2,
          SISTEMA_CONFIG.COL_RESPONSAVEL,
          valoresResponsaveis.length,
          2
        )
        .setValues(
          valoresResponsaveis
        );
    }
  }


  /**
   * ----------------------------------------------------------
   * CATEGORIAS
   * ----------------------------------------------------------
   */

  const categorias =
    Array.isArray(dados.categorias)
      ? dados.categorias
      : [];


  limparListaConfiguracao_(
    aba,
    SISTEMA_CONFIG.COL_CATEGORIA
  );


  gravarListaConfig_(
    aba,
    SISTEMA_CONFIG.COL_CATEGORIA,
    categorias
  );


  /**
   * ----------------------------------------------------------
   * CARTÕES
   * ----------------------------------------------------------
   */

  const cartoes =
    Array.isArray(dados.cartoes)
      ? dados.cartoes
      : [];


  limparListaConfiguracao_(
    aba,
    SISTEMA_CONFIG.COL_CARTAO
  );


  gravarListaConfig_(
    aba,
    SISTEMA_CONFIG.COL_CARTAO,
    cartoes
  );


  /**
   * ----------------------------------------------------------
   * FORMAS DE PAGAMENTO
   * ----------------------------------------------------------
   *
   * Aceita tanto "formas" quanto "formasPagamento".
   */

  const formas =
    Array.isArray(dados.formas)
      ? dados.formas
      : (
          Array.isArray(dados.formasPagamento)
            ? dados.formasPagamento
            : []
        );


  limparListaConfiguracao_(
    aba,
    SISTEMA_CONFIG.COL_FORMA_PAGAMENTO
  );


  gravarListaConfig_(
    aba,
    SISTEMA_CONFIG.COL_FORMA_PAGAMENTO,
    formas
  );


  /**
   * ----------------------------------------------------------
   * ATUALIZAR SISTEMA
   * ----------------------------------------------------------
   */

  aplicarValidacoesSistema_();

  sincronizarCartoes_();


  /**
   * ----------------------------------------------------------
   * REEMBOLSOS
   * ----------------------------------------------------------
   *
   * Se Reembolsos.gs existir, atualiza automaticamente
   * o dropdown de terceiros.
   */

  try {

    if (
      typeof aplicarValidacoesReembolsos_ === 'function'
    ) {

      aplicarValidacoesReembolsos_();
    }

  } catch (erro) {

    console.log(
      'Não foi possível atualizar Reembolsos: ' +
      erro
    );
  }


  SpreadsheetApp.flush();


  return {
    sucesso: true,
    mensagem:
      'Configurações salvas com sucesso.'
  };
}


/**
 * ============================================================
 * CONFIGURAÇÃO INICIAL COMPLETA?
 * ============================================================
 */
function configuracaoInicialCompleta_() {

  const config =
    obterConfiguracoesSistema();


  return (
    config.responsaveis.length > 0 &&
    config.categorias.length > 0 &&
    config.cartoes.length > 0 &&
    config.formas.length > 0
  );
}


/**
 * ============================================================
 * APLICAR VALIDAÇÕES DO SISTEMA
 * ============================================================
 */
function aplicarValidacoesSistema_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const config =
    obterConfiguracoesSistema();


  const responsaveis =
    config.responsaveis
      .map(item =>
        item.nome
      );


  const categorias =
    config.categorias;


  const cartoes =
    config.cartoes;


  const formas =
    config.formas;


  /**
   * ==========================================================
   * 🛒 DESPESAS VARIÁVEIS
   * ==========================================================
   *
   * B Categoria
   * E Forma Pagamento
   * F Cartão
   * G Responsável
   */

  const variaveis =
    ss.getSheetByName(
      '🛒 Desp.Variáveis'
    );


  if (variaveis) {

    aplicarDropdown_(
      variaveis,
      7,
      2,
      categorias
    );

    aplicarDropdown_(
      variaveis,
      7,
      5,
      formas
    );

    aplicarDropdown_(
      variaveis,
      7,
      6,
      cartoes
    );

    aplicarDropdown_(
      variaveis,
      7,
      7,
      responsaveis
    );
  }


  /**
   * ==========================================================
   * 💳 PARCELAMENTOS
   * ==========================================================
   *
   * A Categoria
   * H Cartão
   * I Responsável
   */

  const parcelamentos =
    ss.getSheetByName(
      '💳 Parcelamentos'
    );


  if (parcelamentos) {

    aplicarDropdown_(
      parcelamentos,
      4,
      1,
      categorias
    );

    aplicarDropdown_(
      parcelamentos,
      4,
      8,
      cartoes
    );

    aplicarDropdown_(
      parcelamentos,
      4,
      9,
      responsaveis
    );
  }


  /**
   * ==========================================================
   * 📋 DESPESAS FIXAS
   * ==========================================================
   *
   * A Categoria
   * B Responsável
   * F Forma Pagamento
   */

  const fixas =
    ss.getSheetByName(
      '📋 Desp.Fixas'
    );


  if (fixas) {

    aplicarDropdown_(
      fixas,
      6,
      1,
      categorias
    );

    aplicarDropdown_(
      fixas,
      6,
      2,
      responsaveis
    );

    aplicarDropdown_(
      fixas,
      6,
      6,
      formas
    );
  }


  /**
   * ==========================================================
   * 💸 REEMBOLSOS
   * ==========================================================
   *
   * B = somente Terceiro / Reembolso
   */

  try {

    if (
      typeof aplicarValidacoesReembolsos_ === 'function'
    ) {

      aplicarValidacoesReembolsos_();
    }

  } catch (erro) {

    console.log(
      'Validação de Reembolsos não aplicada: ' +
      erro
    );
  }
}


/**
 * ============================================================
 * SINCRONIZAR CARTÕES
 * ============================================================
 *
 * Cria cartões novos na aba 💳 Cartões.
 *
 * NÃO apaga cartões existentes.
 *
 * Assim preservamos:
 * - limites
 * - fechamento
 * - vencimento
 * - histórico
 */
function sincronizarCartoes_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const abaCartoes =
    ss.getSheetByName(
      '💳 Cartões'
    );


  if (!abaCartoes) {
    return;
  }


  const config =
    obterConfiguracoesSistema();


  const cartoesConfig =
    config.cartoes
      .map(item =>
        String(item || '').trim()
      )
      .filter(item =>
        item !== ''
      );


  if (cartoesConfig.length === 0) {
    return;
  }


  /**
   * ----------------------------------------------------------
   * LOCALIZAR RESUMO
   * ----------------------------------------------------------
   */

  const ultimaLinha =
    Math.max(
      abaCartoes.getLastRow(),
      4
    );


  const valoresColunaA =
    abaCartoes
      .getRange(
        1,
        1,
        ultimaLinha,
        1
      )
      .getDisplayValues()
      .flat();


  let linhaResumo =
    valoresColunaA.findIndex(valor =>
      String(valor).trim() === 'LIMITE TOTAL'
    );


  if (linhaResumo >= 0) {

    linhaResumo += 1;

  } else {

    linhaResumo =
      abaCartoes.getLastRow() + 2;
  }


  /**
   * ----------------------------------------------------------
   * CARTÕES EXISTENTES
   * ----------------------------------------------------------
   */

  const ultimaLinhaCartoes =
    Math.max(
      linhaResumo - 1,
      4
    );


  const quantidadePossivel =
    Math.max(
      ultimaLinhaCartoes - 3,
      1
    );


  const existentes =
    abaCartoes
      .getRange(
        4,
        1,
        quantidadePossivel,
        1
      )
      .getDisplayValues()
      .flat()
      .map(valor =>
        String(valor || '').trim()
      )
      .filter(valor =>
        valor !== ''
      );


  /**
   * ----------------------------------------------------------
   * CRIAR CARTÕES AUSENTES
   * ----------------------------------------------------------
   */

  cartoesConfig.forEach(cartao => {

    const jaExiste =
      existentes.some(existente =>
        existente.toLowerCase() ===
        cartao.toLowerCase()
      );


    if (jaExiste) {
      return;
    }


    /**
     * Encontrar primeira linha vazia antes do resumo.
     */
    let linhaDestino = null;


    for (
      let linha = 4;
      linha < linhaResumo;
      linha++
    ) {

      const nomeAtual =
        String(
          abaCartoes
            .getRange(linha, 1)
            .getDisplayValue() || ''
        ).trim();


      if (nomeAtual === '') {

        linhaDestino = linha;

        break;
      }
    }


    /**
     * Se não existe espaço, insere uma linha
     * antes do resumo.
     */
    if (!linhaDestino) {

      abaCartoes
        .insertRowBefore(
          linhaResumo
        );

      linhaDestino =
        linhaResumo;

      linhaResumo++;
    }


    /**
     * Copiar formato da linha 4.
     */
    if (linhaDestino !== 4) {

      abaCartoes
        .getRange(
          4,
          1,
          1,
          8
        )
        .copyTo(
          abaCartoes.getRange(
            linhaDestino,
            1,
            1,
            8
          ),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
    }


    /**
     * Nome do cartão
     */
    abaCartoes
      .getRange(
        linhaDestino,
        1
      )
      .setValue(
        cartao
      );


    /**
     * D = Limite disponível
     */
    abaCartoes
      .getRange(
        linhaDestino,
        4
      )
      .setFormula(
        `=IF(OR(B${linhaDestino}="";B${linhaDestino}=0);"";B${linhaDestino}-IF(C${linhaDestino}="";0;C${linhaDestino}))`
      );


    /**
     * E = % utilizado
     */
    abaCartoes
      .getRange(
        linhaDestino,
        5
      )
      .setFormula(
        `=IF(OR(B${linhaDestino}="";B${linhaDestino}=0);"";IFERROR(C${linhaDestino}/B${linhaDestino};0))`
      );


    /**
     * H = Status
     */
    abaCartoes
      .getRange(
        linhaDestino,
        8
      )
      .setFormula(
        `=IF(OR(B${linhaDestino}="";B${linhaDestino}=0);"⚪ Configurar";IF(E${linhaDestino}>=0.9;"🔴 Crítico";IF(E${linhaDestino}>=0.75;"🟠 Alto";IF(E${linhaDestino}>=0.5;"🟡 Médio";"🟢 OK"))))`
      );


    existentes.push(
      cartao
    );
  });
}


/**
 * ============================================================
 * APLICAR DROPDOWN
 * ============================================================
 */
function aplicarDropdown_(
  aba,
  primeiraLinha,
  coluna,
  lista
) {

  if (!aba) {
    return;
  }


  const ultimaLinha =
    Math.max(
      aba.getMaxRows(),
      primeiraLinha
    );


  const quantidadeLinhas =
    ultimaLinha -
    primeiraLinha +
    1;


  const intervalo =
    aba.getRange(
      primeiraLinha,
      coluna,
      quantidadeLinhas,
      1
    );


  intervalo.clearDataValidations();


  const valores =
    (lista || [])
      .map(item =>
        String(item || '').trim()
      )
      .filter(item =>
        item !== ''
      );


  if (valores.length === 0) {
    return;
  }


  const regra =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        valores,
        true
      )
      .setAllowInvalid(false)
      .build();


  intervalo
    .setDataValidation(
      regra
    );
}


/**
 * ============================================================
 * LIMPAR ÁREA DA _CONFIG
 * ============================================================
 *
 * quantidadeColunas padrão = 1.
 *
 * Para responsáveis usamos 2 colunas: A:B.
 */
function limparListaConfiguracao_(
  aba,
  coluna,
  quantidadeColunas
) {

  quantidadeColunas =
    quantidadeColunas || 1;


  const ultimaLinha =
    Math.max(
      aba.getLastRow(),
      2
    );


  const quantidadeLinhas =
    ultimaLinha - 1;


  if (quantidadeLinhas <= 0) {
    return;
  }


  aba
    .getRange(
      2,
      coluna,
      quantidadeLinhas,
      quantidadeColunas
    )
    .clearContent();
}


/**
 * ============================================================
 * OBTER LISTA DA _CONFIG
 * ============================================================
 */
function obterListaConfig_(
  aba,
  coluna
) {

  const ultimaLinha =
    aba.getLastRow();


  if (ultimaLinha < 2) {
    return [];
  }


  return aba
    .getRange(
      2,
      coluna,
      ultimaLinha - 1,
      1
    )
    .getValues()
    .flat()
    .map(valor =>
      String(valor || '').trim()
    )
    .filter(valor =>
      valor !== ''
    );
}


/**
 * ============================================================
 * VERIFICAR SE COLUNA POSSUI DADOS
 * ============================================================
 */
function colunaConfigPossuiDados_(
  aba,
  coluna
) {

  return (
    obterListaConfig_(
      aba,
      coluna
    ).length > 0
  );
}


/**
 * ============================================================
 * GRAVAR LISTA NA _CONFIG
 * ============================================================
 */
function gravarListaConfig_(
  aba,
  coluna,
  lista
) {

  const valores =
    (lista || [])
      .map(item =>
        String(item || '').trim()
      )
      .filter(item =>
        item !== ''
      );


  if (valores.length === 0) {
    return;
  }


  garantirLinhasConfig_(
    aba,
    valores.length + 1
  );


  aba
    .getRange(
      2,
      coluna,
      valores.length,
      1
    )
    .setValues(
      valores.map(item =>
        [item]
      )
    );
}


/**
 * ============================================================
 * GARANTIR LINHAS NA _CONFIG
 * ============================================================
 */
function garantirLinhasConfig_(
  aba,
  quantidadeNecessaria
) {

  const linhasAtuais =
    aba.getMaxRows();


  if (
    linhasAtuais <
    quantidadeNecessaria
  ) {

    aba.insertRowsAfter(
      linhasAtuais,
      quantidadeNecessaria -
      linhasAtuais
    );
  }
}


/**
 * ============================================================
 * LISTA PARA VALIDAÇÃO
 * ============================================================
 *
 * Mantida por compatibilidade com outras partes
 * do projeto.
 */
function obterListaValidacao_(
  coluna
) {

  const aba =
    prepararConfiguracoes_();

  return obterListaConfig_(
    aba,
    coluna
  );
}