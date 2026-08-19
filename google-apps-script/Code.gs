// Backend do Virtus-Detector, em Google Apps Script.
//
// Faz três coisas:
//   1. Admin cria uma prova (admin.html) → guarda-a numa Google Sheet própria.
//   2. Aluno abre o link com ?prova=ID (index.html) → devolve as perguntas dessa prova.
//   3. Aluno finaliza a sessão → envia um único email ao professor com a folha de
//      respostas do aluno + o relatório de vigilância (estatísticas, eventos, fotos).
//
// A "base de dados" é uma Google Sheet criada automaticamente pelo próprio script na
// primeira vez que é preciso guardar algo — não depende de o script estar associado
// a nenhuma planilha específica, por isso funciona com uma implantação standalone.
//
// COMO IMPLANTAR / ATUALIZAR:
// 1. https://script.google.com → abra o projeto "Virtus-Detector Relatórios" (o que
//    já está implantado).
// 2. Apague o conteúdo do Code.gs e cole o conteúdo deste ficheiro. Guarde (Ctrl+S).
// 3. Implantar → Gerir implantações → editar (ícone de lápis) → Versão: Nova versão
//    → Implantar. Isto mantém a MESMA URL que já está no index.html e no admin.html.
// 4. Se pedir para rever permissões (pode acontecer por causa das novas
//    funcionalidades), autorize normalmente.

const DESTINATARIO = "antoniorundy6@gmail.com";
const NOME_BASE_DADOS = "Virtus-Detector — Base de Dados";
const FOLHA_PROVAS = "Provas";
const FOLHA_RELATORIOS = "Relatorios";
const FOLHA_RESPOSTAS = "Respostas";

// ══════════════════════════════════════════════
// BASE DE DADOS (auto-criada na primeira utilização)
// ══════════════════════════════════════════════
function obterBaseDados() {
  const props = PropertiesService.getScriptProperties();
  const idGuardado = props.getProperty("SPREADSHEET_ID");
  if (idGuardado) {
    try {
      return SpreadsheetApp.openById(idGuardado);
    } catch (e) { /* a folha foi apagada ou o ID é inválido — recria abaixo */ }
  }
  const nova = SpreadsheetApp.create(NOME_BASE_DADOS);
  props.setProperty("SPREADSHEET_ID", nova.getId());
  return nova;
}

function obterOuCriarFolha(ss, nome, cabecalho) {
  let sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
    sheet.appendRow(cabecalho);
  }
  return sheet;
}

// ══════════════════════════════════════════════
// GET — o aluno pede as perguntas de uma prova: ?prova=ID
// ══════════════════════════════════════════════
function doGet(e) {
  const id = e.parameter.prova;
  if (!id) {
    return saidaJSON({ erro: "Parâmetro 'prova' em falta." });
  }
  const prova = obterProva(id);
  return saidaJSON(prova || { erro: "Prova não encontrada." });
}

function obterProva(id) {
  const ss = obterBaseDados();
  const sheet = ss.getSheetByName(FOLHA_PROVAS);
  if (!sheet) return null;
  const linhas = sheet.getDataRange().getValues();
  for (let i = 1; i < linhas.length; i++) {
    if (String(linhas[i][0]) === String(id)) {
      return {
        id: linhas[i][0],
        titulo: linhas[i][1],
        enunciado: linhas[i][2],
        perguntas: JSON.parse(linhas[i][3] || "[]")
      };
    }
  }
  return null;
}

// ══════════════════════════════════════════════
// POST — dois casos: criar prova (admin) ou submeter sessão (aluno)
// ══════════════════════════════════════════════
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);

    if (dados.acao === "criarProva") {
      guardarProva(dados);
      return saidaJSON({ ok: true, id: dados.id });
    }

    enviarEmailRelatorio(dados);
    registarNaFolha(dados);
    if (dados.perguntas && dados.perguntas.length) registarRespostas(dados);

    return saidaJSON({ ok: true });
  } catch (err) {
    return saidaJSON({ ok: false, erro: String(err) });
  }
}

function saidaJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════
// PROVAS — criadas pelo admin.html
// ══════════════════════════════════════════════
function guardarProva(dados) {
  const ss = obterBaseDados();
  const sheet = obterOuCriarFolha(ss, FOLHA_PROVAS, ["ID", "Título", "Enunciado", "PerguntasJSON", "CriadoEm"]);
  const linhas = sheet.getDataRange().getValues();
  for (let i = 1; i < linhas.length; i++) {
    if (String(linhas[i][0]) === String(dados.id)) {
      sheet.getRange(i + 1, 1, 1, 4).setValues([[
        dados.id, dados.titulo || "", dados.enunciado || "", JSON.stringify(dados.perguntas || [])
      ]]);
      return;
    }
  }
  sheet.appendRow([dados.id, dados.titulo || "", dados.enunciado || "", JSON.stringify(dados.perguntas || []), new Date()]);
}

// ══════════════════════════════════════════════
// EMAIL — relatório de vigilância + folha de respostas (quando há prova associada)
// ══════════════════════════════════════════════
function linhaTabela(label, valor) {
  return '<tr>' +
    '<td style="padding:4px 16px 4px 0;color:#666;font-size:13px;">' + label + '</td>' +
    '<td style="padding:4px 0;font-weight:600;font-size:13px;">' + valor + '</td>' +
    '</tr>';
}

function corIndice(v) {
  if (v >= 80) return "#2e9e6b";
  if (v >= 60) return "#c99a2e";
  return "#d1495b";
}

function construirSecaoRespostas(dados) {
  if (!dados.perguntas || !dados.perguntas.length) return "";
  const linhas = dados.perguntas.map(function (p, i) {
    const resposta = (dados.respostas && dados.respostas[i] != null && dados.respostas[i] !== "")
      ? dados.respostas[i]
      : "<em>(sem resposta)</em>";
    return '<div style="margin-bottom:14px;padding-bottom:14px;border-bottom:1px solid #eee;">' +
      '<div style="font-size:13px;font-weight:600;color:#333;margin-bottom:4px;">' + (i + 1) + '. ' + p.pergunta + '</div>' +
      '<div style="font-size:13px;color:#12151d;background:#f5f6f8;padding:8px 12px;border-radius:6px;white-space:pre-wrap;">' + resposta + '</div>' +
      '</div>';
  }).join("");

  return '<div style="background:#eef4ff;border:1px solid #cddcf5;border-radius:8px;padding:18px 20px;margin-bottom:24px;">' +
    '<h3 style="margin:0 0 4px;font-size:16px;color:#1a3a6e;">📄 Folha de Respostas — ' + (dados.provaTitulo || "Prova") + '</h3>' +
    '<p style="margin:0 0 14px;font-size:12px;color:#5578ad;">Corrija manualmente abaixo.</p>' +
    linhas +
    '</div>';
}

function enviarEmailRelatorio(dados) {
  const est = dados.estatisticas || {};
  const porTipo = est.eventosPorTipo || {};
  const contagem = est.contagemPorInfracao || {};

  const linhasInfracoes = Object.keys(contagem).length
    ? Object.keys(contagem).map(function (k) { return linhaTabela(k, contagem[k] + "x"); }).join("")
    : '<tr><td colspan="2" style="padding:8px 0;color:#2e9e6b;font-size:13px;">Nenhuma infração registada</td></tr>';

  const linhasEventos = (dados.eventos || []).map(function (ev) {
    return '<tr>' +
      '<td style="padding:3px 10px 3px 0;color:#888;font-family:monospace;font-size:11px;">[' + ev.ts + ']</td>' +
      '<td style="padding:3px 10px 3px 0;font-size:11px;"><strong>' + ev.tipo + '</strong></td>' +
      '<td style="padding:3px 0;font-size:11px;">' + ev.descricao + '</td>' +
      '</tr>';
  }).join("");

  const fotos = dados.fotos || [];
  const fotosHtml = fotos.length
    ? '<p style="color:#666;font-size:12px;">' + fotos.length + ' fotografia(s) de evidência capturadas — anexadas a este email.</p>'
    : "";

  const html =
    '<div style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;">' +
      '<div style="background:#12151d;padding:20px 24px;border-radius:10px 10px 0 0;">' +
        '<h2 style="color:#fff;margin:0;">Virtus-Detector' + (dados.provaTitulo ? " — " + dados.provaTitulo : "") + '</h2>' +
        '<p style="color:#98a0b3;margin:6px 0 0;font-size:13px;">' + new Date().toLocaleString("pt-PT") + '</p>' +
      '</div>' +
      '<div style="border:1px solid #e3e3e3;border-top:none;padding:24px;border-radius:0 0 10px 10px;">' +

        construirSecaoRespostas(dados) +

        '<table style="width:100%;border-collapse:collapse;margin-bottom:20px;">' +
          linhaTabela("Candidato", dados.nome || "—") +
          linhaTabela("Número", dados.numero || "—") +
          linhaTabela("Duração da sessão", est.duracaoFormatada || "—") +
          linhaTabela("Total de eventos registados", est.totalEventos != null ? est.totalEventos : 0) +
          linhaTabela("Eventos do sistema", porTipo.SISTEMA || 0) +
          linhaTabela("Alertas", porTipo.ALERTA || 0) +
          linhaTabela("Ocorrências graves", porTipo.GRAVE || 0) +
          linhaTabela("Fotografias capturadas", est.fotosCapturadas != null ? est.fotosCapturadas : fotos.length) +
        '</table>' +

        '<div style="background:' + corIndice(dados.confianca) + ';color:#fff;padding:14px 18px;border-radius:8px;margin-bottom:20px;">' +
          '<div style="font-size:13px;opacity:0.9;">Índice de Integridade Final</div>' +
          '<div style="font-size:32px;font-weight:700;">' + dados.confianca + '%</div>' +
        '</div>' +

        '<h3 style="margin:0 0 8px;font-size:15px;">Infrações por tipo</h3>' +
        '<table style="width:100%;border-collapse:collapse;margin-bottom:20px;">' + linhasInfracoes + '</table>' +

        '<h3 style="margin:0 0 8px;font-size:15px;">Resumo da IA</h3>' +
        '<p style="font-size:13px;line-height:1.6;color:#333;white-space:pre-wrap;">' + (dados.resumoIA || "Indisponível.") + '</p>' +

        '<h3 style="margin:20px 0 8px;font-size:15px;">Registo completo de eventos</h3>' +
        '<table style="width:100%;border-collapse:collapse;">' + linhasEventos + '</table>' +

        fotosHtml +
      '</div>' +
    '</div>';

  const anexos = fotos.map(function (dataUrl, i) {
    const base64 = dataUrl.split(",")[1];
    return Utilities.newBlob(Utilities.base64Decode(base64), "image/jpeg", "evidencia_" + (i + 1) + ".jpg");
  });

  const assuntoBase = dados.provaTitulo ? dados.provaTitulo : "Relatório de Exame";
  MailApp.sendEmail({
    to: dados.destinatario || DESTINATARIO,
    subject: assuntoBase + " — " + (dados.nome || "Candidato") + " (" + dados.confianca + "%)",
    htmlBody: html,
    attachments: anexos
  });
}

// ══════════════════════════════════════════════
// REGISTOS (Sheets) — cópia de segurança fora do email
// ══════════════════════════════════════════════
function registarNaFolha(dados) {
  const ss = obterBaseDados();
  const sheet = obterOuCriarFolha(ss, FOLHA_RELATORIOS, ["Data", "Prova", "Nome", "Número", "Índice", "Duração", "Alertas", "Graves", "Observações"]);
  const est = dados.estatisticas || {};
  const porTipo = est.eventosPorTipo || {};
  sheet.appendRow([
    new Date(),
    dados.provaTitulo || "",
    dados.nome,
    dados.numero,
    dados.confianca,
    est.duracaoFormatada || "",
    porTipo.ALERTA || 0,
    porTipo.GRAVE || 0,
    dados.observacoes || ""
  ]);
}

function registarRespostas(dados) {
  const ss = obterBaseDados();
  const sheet = obterOuCriarFolha(ss, FOLHA_RESPOSTAS, ["Data", "Prova", "Nome", "Número", "Pergunta", "Resposta"]);
  dados.perguntas.forEach(function (p, i) {
    const resposta = (dados.respostas && dados.respostas[i] != null) ? dados.respostas[i] : "";
    sheet.appendRow([new Date(), dados.provaTitulo || "", dados.nome, dados.numero, p.pergunta, resposta]);
  });
}
