// Dashboard Patrimônio — Escrita de volta para a Google Sheet
//
// Como instalar:
// 1. Acesse script.google.com → Novo projeto → cole este código
// 2. Implante como Web App: Implantar → Nova implantação → Tipo: App da Web
//    - Executar como: Eu
//    - Quem pode acessar: Qualquer pessoa
// 3. Copie a URL gerada (termina em /exec)
// 4. Cole em SCRIPT_URL no index.html

var SHEET_ID = '1EI6lWiym8cme9akLNhZzxw9V-fe2X-ZSz1LpYq8mPIg';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    if (data.type === 'despesa') {
      appendDespesa(ss, data);
    } else if (data.type === 'rewriteDespesas') {
      rewriteDespesas(ss, data);
    } else {
      atualizarPatrimonio(ss, data);
      if (data.appendHistorico) appendHistorico(ss, data);
    }
    return ContentService
      .createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function appendDespesa(ss, data) {
  var aba = ss.getSheetByName('despesas');
  aba.appendRow([data.data, data.tipo, data.descricao, data.categoria, data.valor, data.moeda]);
}

function rewriteDespesas(ss, data) {
  var aba = ss.getSheetByName('despesas');
  var lastRow = aba.getLastRow();
  if (lastRow > 1) aba.getRange(2, 1, lastRow - 1, 6).clearContent();
  if (data.despesas && data.despesas.length > 0) {
    var rows = data.despesas.map(function(d) {
      return [d.data, d.tipo, d.descricao, d.categoria, d.valor, d.moeda];
    });
    aba.getRange(2, 1, rows.length, 6).setValues(rows);
  }
}

function atualizarPatrimonio(ss, data) {
  var aba = ss.getSheetByName('patrimonio_atual');
  var rows = aba.getDataRange().getValues();
  var campos = ['ipca','prefixado','rv','pos','fgts','k401','broker','small',
                'reserva','fx','renda','poupanca','updatedAt'];
  rows.forEach(function(row, i) {
    if (campos.indexOf(row[0]) !== -1 && data[row[0]] !== undefined) {
      aba.getRange(i + 1, 2).setValue(data[row[0]]);
    }
  });
}

function appendHistorico(ss, data) {
  var aba = ss.getSheetByName('historico_mensal');
  var rows = aba.getDataRange().getValues();
  var last = rows[rows.length - 1];
  if (last && String(last[0]).toLowerCase() === String(data.mes).toLowerCase()) {
    aba.getRange(rows.length, 1, 1, 3).setValues([[data.mes, data.brl, data.usd]]);
  } else {
    aba.appendRow([data.mes, data.brl, data.usd]);
  }
}
