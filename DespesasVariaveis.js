function filtrarDespesasVariaveisPorMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const abaDespesas = ss.getSheetByName('🛒 Desp.Variáveis');
  const abaDashboard = ss.getSheetByName('📊 Dashboard');

  if (!abaDespesas || !abaDashboard) return;

  // Mês escolhido no Dashboard
  const mesReferencia = abaDashboard.getRange('B7').getValue();

  if (!(mesReferencia instanceof Date)) return;

  // Na sua planilha:
  // Linha 6 = cabeçalho
  // Linha 7 = primeiro lançamento
  // Coluna A = Data
  const primeiraLinhaDados = 7;
  const ultimaLinha = abaDespesas.getLastRow();

  if (ultimaLinha < primeiraLinhaDados) return;

  // Antes de filtrar, mostra novamente todas as linhas
  abaDespesas.showRows(
    primeiraLinhaDados,
    ultimaLinha - primeiraLinhaDados + 1
  );

  // Pega as datas da coluna A
  const datas = abaDespesas.getRange(
    primeiraLinhaDados,
    1,
    ultimaLinha - primeiraLinhaDados + 1,
    1
  ).getValues();

  const mes = mesReferencia.getMonth();
  const ano = mesReferencia.getFullYear();

  let inicioBlocoOculto = null;
  let tamanhoBloco = 0;

  function esconderBloco() {
    if (inicioBlocoOculto !== null && tamanhoBloco > 0) {
      abaDespesas.hideRows(
        inicioBlocoOculto,
        tamanhoBloco
      );
    }

    inicioBlocoOculto = null;
    tamanhoBloco = 0;
  }

  datas.forEach((linha, indice) => {
    const data = linha[0];
    const numeroLinha = primeiraLinhaDados + indice;

    const pertenceAoMes =
      data instanceof Date &&
      data.getFullYear() === ano &&
      data.getMonth() === mes;

    if (!pertenceAoMes) {

      if (inicioBlocoOculto === null) {
        inicioBlocoOculto = numeroLinha;
      }

      tamanhoBloco++;

    } else {

      esconderBloco();

    }
  });

  esconderBloco();
}