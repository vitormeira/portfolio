// Dashboard Patrimônio — Email mensal de aportes
// Configura um gatilho: Executar no dia 1 de cada mês às 08:00
//
// Como instalar:
// 1. Acesse script.google.com → Novo projeto
// 2. Cole este código (substitua o SHEET_ID e EMAIL)
// 3. Executar → configurarGatilho() uma única vez para criar o trigger
// 4. Autorize as permissões (Gmail + Sheets)

var SHEET_ID = '1EI6lWiym8cme9akLNhZzxw9V-fe2X-ZSz1LpYq8mPIg';
var EMAIL    = 'SEU_EMAIL_AQUI@gmail.com';

// ─── Gatilho mensal ────────────────────────────────────────────────────────
function configurarGatilho() {
  // Remove gatilhos antigos para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarChecklistMensal') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('enviarChecklistMensal')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();
  Logger.log('Gatilho criado: dia 1 de cada mês às 08:00');
}

// ─── Função principal ──────────────────────────────────────────────────────
function enviarChecklistMensal() {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var d   = lerDados(ss);
  var mes = Utilities.formatDate(new Date(), 'America/Sao_Paulo', "MMM/yy")
              .replace(/^\w/, function(c){ return c.toUpperCase(); });

  var brl   = d.ipca + d.prefixado + d.rv + d.pos + d.fgts;
  var usd   = d.k401 + d.broker + d.small + d.reserva;
  var total = brl + usd * d.fx;
  var META  = 20000 * 12 / 0.04 * Math.pow(1.055, 7);
  var prog  = (total / META * 100).toFixed(1);

  var aporteAtual = d.renda * (d.poupanca / 100);
  var aporteNec   = calcAporteNecessario(d, total, META);
  var scPct       = d.small * d.fx / total * 100;
  var resMeses    = d.reserva / d.renda;
  var score       = calcScore(d, total, META);

  var alertas = [];
  if (scPct < 3)
    alertas.push({warn:true, msg:'Small Caps em '+scPct.toFixed(1)+'% — reforçar IWM/AVUV (alvo mínimo 3%)'});
  if (aporteAtual < aporteNec - 200)
    alertas.push({warn:true, msg:'Aporte atual abaixo do necessário em USD '+Math.round(aporteNec-aporteAtual).toLocaleString()+'/mês'});
  if (resMeses < 6)
    alertas.push({warn:true, msg:'Reserva de emergência: '+resMeses.toFixed(1)+' meses (alvo: 6)'});

  var htmlBody = montarEmail(mes, total, brl, usd, d.fx, prog, score, aporteAtual, aporteNec, alertas);

  MailApp.sendEmail({
    to:       EMAIL,
    subject:  '[Patrimônio vm] Aportes de ' + mes + ' — score ' + score + '/100',
    htmlBody: htmlBody
  });

  Logger.log('Email enviado para ' + EMAIL);
}

// ─── Leitura da planilha ───────────────────────────────────────────────────
function lerDados(ss) {
  var rows = ss.getSheetByName('patrimonio_atual').getDataRange().getValues().slice(1);
  var d = {small:0};
  rows.forEach(function(r) {
    var v = r[1];
    d[r[0]] = (v !== '' && !isNaN(v)) ? parseFloat(v) : String(v);
  });
  return d;
}

// ─── Cálculos (espelha o index.html) ──────────────────────────────────────
function calcAporteNecessario(d, total, META) {
  var n = 7; var r = 0.065;
  var rn = Math.pow(1 + r, n);
  return Math.max(0, (META - total * rn) * r / (rn - 1) / 12 / d.fx);
}

function calcScore(d, total, META) {
  var prog     = Math.min(total / META, 1);
  var scPct    = d.small * d.fx / total * 100;
  var progPts  = Math.round(prog * 40);
  var poupPts  = Math.round(Math.min(d.poupanca / 35, 1) * 30);
  var allocPts = scPct >= 3 && scPct <= 5 ? 20
               : scPct < 3 ? Math.round(scPct / 3 * 20)
               : Math.max(0, Math.round(20 - (scPct - 5) * 4));
  var resPts   = Math.round(Math.min((d.reserva / d.renda) / 6, 1) * 10);
  return progPts + poupPts + allocPts + resPts;
}

// ─── Template do email ─────────────────────────────────────────────────────
function montarEmail(mes, total, brl, usd, fx, prog, score, aporteAtual, aporteNec, alertas) {
  var scoreColor = score >= 75 ? '#16a34a' : score >= 50 ? '#d97706' : '#dc2626';
  var scoreLabel = score >= 75 ? 'Boa saúde' : score >= 50 ? 'Atenção' : 'Crítico';

  var alertasHtml = alertas.length === 0
    ? '<div style="padding:12px 14px;background:#dcfce7;border-radius:8px;color:#166534;font-size:12px">✓ Nenhum desvio detectado este mês.</div>'
    : alertas.map(function(a) {
        return '<div style="padding:10px 14px;background:#fef3c7;border-radius:8px;border:1px solid #fde68a;color:#92400e;font-size:12px;margin-bottom:6px">⚠ ' + a.msg + '</div>';
      }).join('');

  var tarefas = [
    'Aporte VOO (S&P 500) — 40% do aporte mensal',
    'Aporte VTI (Total Market) — 20% do aporte mensal',
    'Aporte SCHD (Dividends) — 17% do aporte mensal',
    'Aporte BND (Bonds) — 11% do aporte mensal',
    'Aporte IWM/AVUV (Small Caps) — 4% do aporte mensal',
    'Atualizar câmbio USD/BRL na Google Sheet',
    'Baixar extrato XP e atualizar saldos BRL',
    'Confirmar contribuição 401(k) + employer match',
  ];

  var tarefasHtml = tarefas.map(function(t) {
    return '<tr><td style="padding:9px 0;border-bottom:1px solid #e8e8e8;vertical-align:top">' +
      '<div style="width:16px;height:16px;border:1.5px solid #ccc;border-radius:50%;display:inline-block;vertical-align:middle;margin-right:10px"></div>' +
      '<span style="font-size:13px;vertical-align:middle">' + t + '</span></td></tr>';
  }).join('');

  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#fafafa;font-family:Inter,Helvetica,sans-serif">' +
    '<div style="max-width:540px;margin:40px auto;background:#fff;border:1px solid #e8e8e8;border-radius:12px;overflow:hidden">' +

    '<div style="padding:28px 32px 0">' +
      '<div style="font-size:12px;color:#999;margin-bottom:20px">patrimônio · vm</div>' +
      '<div style="font-size:22px;font-weight:400;margin-bottom:4px">Checklist de aportes</div>' +
      '<div style="font-size:13px;color:#999;margin-bottom:28px">' + mes + '</div>' +

      '<div style="display:flex;gap:16px;margin-bottom:24px">' +
        '<div style="flex:1;background:#f9fafb;border-radius:10px;padding:16px">' +
          '<div style="font-size:10px;color:#999;text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px">Patrimônio total</div>' +
          '<div style="font-size:22px;font-weight:300;letter-spacing:-1px">R$ ' + Math.round(total).toLocaleString('pt-BR') + '</div>' +
          '<div style="font-size:11px;color:#999;margin-top:4px">Progresso: ' + prog + '% da meta</div>' +
        '</div>' +
        '<div style="width:90px;background:#f9fafb;border-radius:10px;padding:16px;text-align:center">' +
          '<div style="font-size:10px;color:#999;text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px">Score</div>' +
          '<div style="font-size:28px;font-weight:300;color:' + scoreColor + '">' + score + '</div>' +
          '<div style="font-size:10px;color:' + scoreColor + '">' + scoreLabel + '</div>' +
        '</div>' +
      '</div>' +

      alertasHtml.length ? '<div style="margin-bottom:20px">' + alertasHtml + '</div>' : '' +

      '<div style="font-size:10px;color:#999;text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px">Tarefas do mês</div>' +
      '<table style="width:100%;border-collapse:collapse"><tbody>' + tarefasHtml + '</tbody></table>' +

      '<div style="margin-top:20px;margin-bottom:28px;padding:14px 16px;background:#f9fafb;border-radius:8px;font-size:12px;color:#555">' +
        'Aporte atual: <strong>USD ' + Math.round(aporteAtual).toLocaleString('en-US') + '/mês</strong>' +
        ' &nbsp;·&nbsp; Necessário para meta: <strong>USD ' + Math.round(aporteNec).toLocaleString('en-US') + '/mês</strong>' +
      '</div>' +
    '</div>' +

    '<div style="padding:16px 32px;background:#f9fafb;border-top:1px solid #e8e8e8;font-size:11px;color:#bbb">' +
      'Dashboard em github.com/vitormeira/portfolio &nbsp;·&nbsp; Gerado automaticamente' +
    '</div>' +
  '</div></body></html>';
}
