// Google Apps Script que recebe o relatório da sessão do Virtus-Detector (index.html)
// e envia sempre um email formatado, com estatísticas de todos os aspetos observados,
// para o supervisor. Também mantém um registo numa folha de cálculo (opcional).
//
// COMO IMPLANTAR:
// 1. Abra https://script.google.com e crie um novo projeto (ou reutilize o que já
//    gerava o WEBHOOK_URL atual — vá a "Extensões > Apps Script" a partir de uma
//    Google Sheet, ou "Ficheiro > Novo projeto" em script.google.com).
// 2. Apague o conteúdo do ficheiro Code.gs e cole o conteúdo deste ficheiro.
// 3. Guarde o projeto.
// 4. Clique em "Implantar" (Deploy) → "Nova implantação" → tipo "Aplicação Web".
//      - Executar como: a sua própria conta
//      - Quem tem acesso: qualquer pessoa
// 5. Autorize as permissões pedidas (a primeira vez pede para enviar email em seu nome).
// 6. Copie o URL gerado ("URL da aplicação Web") e cole-o na constante WEBHOOK_URL
//    do ficheiro index.html do projeto Virtus-Detector, substituindo o valor atual.
// 7. Sempre que editar este script, tem de fazer "Gerir implantações" → editar →
//    nova versão, para as alterações entrarem em vigor no mesmo URL.

const DESTINATARIO = "antoniorundy6@gmail.com";
const REGISTAR_EM_FOLHA = true; // true = também grava uma linha resumo numa Sheet
const NOME_FOLHA = "Relatorios";

function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    enviarEmailRelatorio(dados);
    if (REGISTAR_EM_FOLHA) registarNaFolha(dados);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

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
        '<h2 style="color:#fff;margin:0;">Virtus-Detector — Relatório de Sessão</h2>' +
        '<p style="color:#98a0b3;margin:6px 0 0;font-size:13px;">' + new Date().toLocaleString("pt-PT") + '</p>' +
      '</div>' +
      '<div style="border:1px solid #e3e3e3;border-top:none;padding:24px;border-radius:0 0 10px 10px;">' +

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

  MailApp.sendEmail({
    to: dados.destinatario || DESTINATARIO,
    subject: "Relatório de Exame — " + (dados.nome || "Candidato") + " (" + dados.confianca + "%)",
    htmlBody: html,
    attachments: anexos
  });
}

function registarNaFolha(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return; // script não está associado a nenhuma Sheet — ignora silenciosamente
  const sheet = ss.getSheetByName(NOME_FOLHA) || ss.insertSheet(NOME_FOLHA);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Data", "Nome", "Número", "Índice", "Duração", "Alertas", "Graves", "Observações"]);
  }
  const est = dados.estatisticas || {};
  const porTipo = est.eventosPorTipo || {};
  sheet.appendRow([
    new Date(),
    dados.nome,
    dados.numero,
    dados.confianca,
    est.duracaoFormatada || "",
    porTipo.ALERTA || 0,
    porTipo.GRAVE || 0,
    dados.observacoes || ""
  ]);
}
