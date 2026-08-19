// Backend do Virtus-Detector, em Google Apps Script.
// Criado para Virtus LDA — António Rundy Manuel Fernando.
//
// Faz quatro coisas:
//   1. Admin cria uma prova (admin.html) → guarda-a numa Google Sheet própria.
//   2. Aluno abre o link com ?prova=ID (index.html) → devolve as perguntas dessa prova.
//   3. Aluno finaliza a sessão → gera DOIS PDFs (relatório de vigilância com as fotos
//      de evidência embutidas, e a folha de respostas para correção manual) e envia-os
//      por email ao professor. O aluno NUNCA recebe nem vê nenhum PDF — fica só no
//      servidor e no email do professor.
//   4. Mantém uma cópia de segurança de tudo (provas, relatórios, respostas) numa
//      Google Sheet criada automaticamente.
//
// A "base de dados" é criada sozinha na primeira utilização — não depende de o script
// estar associado a nenhuma planilha específica, por isso funciona com implantação
// standalone.
//
// COMO IMPLANTAR / ATUALIZAR:
// 1. https://script.google.com → abra o projeto já implantado.
// 2. Apague o conteúdo do Code.gs e cole o conteúdo deste ficheiro. Guarde (Ctrl+S).
// 3. Implantar → Gerir implantações → editar (ícone de lápis) → Versão: Nova versão
//    → Implantar. Mantém a mesma URL — não precisa mexer no index.html/admin.html.
// 4. Se pedir para rever permissões (agora o script também cria documentos do Google
//    Docs temporários para gerar os PDFs), autorize normalmente.

const DESTINATARIO = "antoniorundy6@gmail.com";
const NOME_BASE_DADOS = "Virtus-Detector — Base de Dados";
const FOLHA_PROVAS = "Provas";
const FOLHA_RELATORIOS = "Relatorios";
const FOLHA_RESPOSTAS = "Respostas";
const FOLHA_ACESSOS = "Acessos";
const ASSINATURA = "Virtus-Detector © Virtus LDA — Criado por António Rundy Manuel Fernando";

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
// Ou verifica (sem registar nada) se já existe submissão completa: ?prova=ID&verificarAcesso=email
// ══════════════════════════════════════════════
function doGet(e) {
  const id = e.parameter.prova;
  if (!id) return saidaJSON({ erro: "Parâmetro 'prova' em falta." });

  if (e.parameter.verificarAcesso) {
    return saidaJSON(verificarAcesso(id, e.parameter.verificarAcesso));
  }

  const prova = obterProva(id);
  return saidaJSON(prova || { erro: "Prova não encontrada." });
}

// Link de utilização única por email — mas só conta depois de uma SUBMISSÃO
// COMPLETA (ver registarConclusao, chamada no fim do doPost). Esta função aqui
// apenas verifica, não regista nada, por isso o aluno pode reabrir/atualizar o
// link livremente enquanto ainda não tiver terminado a prova.
function verificarAcesso(provaId, email) {
  const emailNormalizado = String(email).trim().toLowerCase();
  if (!emailNormalizado || emailNormalizado.indexOf("@") === -1) {
    return { permitido: false, motivo: "Email inválido." };
  }

  const ss = obterBaseDados();
  const sheet = obterOuCriarFolha(ss, FOLHA_ACESSOS, ["ProvaID", "Email", "DataConclusao"]);
  const linhas = sheet.getDataRange().getValues();
  for (let i = 1; i < linhas.length; i++) {
    if (String(linhas[i][0]) === String(provaId) && String(linhas[i][1]).toLowerCase() === emailNormalizado) {
      return { permitido: false, motivo: "Já existe uma submissão completa desta prova com este email." };
    }
  }
  return { permitido: true };
}

// Chamada só depois de o email de vigilância ser enviado com sucesso — é o que
// efetivamente "gasta" o link para este email nesta prova.
function registarConclusao(dados) {
  if (!dados.provaId || !dados.emailAluno) return;
  const ss = obterBaseDados();
  const sheet = obterOuCriarFolha(ss, FOLHA_ACESSOS, ["ProvaID", "Email", "DataConclusao"]);
  sheet.appendRow([dados.provaId, String(dados.emailAluno).trim().toLowerCase(), new Date()]);
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
    registarConclusao(dados);

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
// UTILITÁRIOS
// ══════════════════════════════════════════════
function corIndice(v) {
  if (v >= 80) return "#16a34a";
  if (v >= 60) return "#ca8a04";
  return "#dc2626";
}

function base64ParaBlob(dataUrl, nome) {
  const base64 = String(dataUrl).split(",")[1] || "";
  return Utilities.newBlob(Utilities.base64Decode(base64), "image/jpeg", nome);
}

function nomeArquivoSeguro(texto) {
  return String(texto || "candidato").replace(/[^\w\-]+/g, "_");
}

// Aplica um estilo simples e consistente a uma tabela de 2 colunas (label/valor)
function estilizarTabelaInfo(tabela) {
  for (let i = 0; i < tabela.getNumRows(); i++) {
    const row = tabela.getRow(i);
    row.getCell(0).setBackgroundColor("#f5f6f8");
    row.getCell(0).editAsText().setBold(true).setFontSize(10).setForegroundColor("#444444");
    if (row.getNumCells() > 1) row.getCell(1).editAsText().setFontSize(10).setForegroundColor("#0f172a");
  }
}

function rodapeMarca(body) {
  body.appendHorizontalRule();
  const p = body.appendParagraph(ASSINATURA);
  p.editAsText().setForegroundColor("#a0aec0").setItalic(true).setFontSize(9);
}

function cabecalhoDocumento(body, subtitulo) {
  const titulo = body.appendParagraph("Virtus-Detector");
  titulo.setHeading(DocumentApp.ParagraphHeading.TITLE);
  titulo.editAsText().setForegroundColor("#0f172a");

  const sub = body.appendParagraph(subtitulo);
  sub.setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  sub.editAsText().setForegroundColor("#1a56db");

  body.appendParagraph(new Date().toLocaleString("pt-PT"))
    .editAsText().setForegroundColor("#838ba0").setItalic(true).setFontSize(10);
  body.appendHorizontalRule();
}

// Converte um Google Doc temporário em PDF e apaga o Doc, devolvendo só o PDF
function docParaPdfEDescartar(doc, nomeFicheiro) {
  doc.saveAndClose();
  const pdf = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  pdf.setName(nomeFicheiro);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdf;
}

// ══════════════════════════════════════════════
// PDF 1 — RELATÓRIO DE VIGILÂNCIA (com as fotos de evidência embutidas)
// ══════════════════════════════════════════════
function gerarPdfVigilancia(dados) {
  const est = dados.estatisticas || {};
  const porTipo = est.eventosPorTipo || {};
  const contagem = est.contagemPorInfracao || {};
  const fotos = dados.fotos || [];

  const doc = DocumentApp.create("tmp_vigilancia_" + Date.now());
  const body = doc.getBody();
  body.setMarginTop(40).setMarginBottom(40).setMarginLeft(50).setMarginRight(50);

  cabecalhoDocumento(body, "Relatório de Vigilância de Sessão");

  estilizarTabelaInfo(body.appendTable([
    ["Candidato", dados.nome || "—"],
    ["Número", dados.numero || "—"],
    ["Email", dados.emailAluno || "—"],
    ["Prova associada", dados.provaTitulo || "—"],
    ["Duração da sessão", est.duracaoFormatada || "—"],
    ["Total de eventos registados", String(est.totalEventos != null ? est.totalEventos : 0)],
    ["Eventos do sistema", String(porTipo.SISTEMA || 0)],
    ["Alertas", String(porTipo.ALERTA || 0)],
    ["Ocorrências graves", String(porTipo.GRAVE || 0)],
    ["Fotografias de evidência capturadas", String(fotos.length)]
  ]));

  const indicePar = body.appendParagraph("ÍNDICE DE INTEGRIDADE FINAL: " + dados.confianca + "%");
  indicePar.setSpacingBefore(16).setSpacingAfter(16);
  indicePar.editAsText().setBold(true).setFontSize(15).setForegroundColor(corIndice(dados.confianca));

  body.appendParagraph("Infrações por Tipo").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const linhasInfra = Object.keys(contagem).length
    ? Object.keys(contagem).map(function (k) { return [k, contagem[k] + "x"]; })
    : [["Nenhuma infração registada", ""]];
  estilizarTabelaInfo(body.appendTable(linhasInfra));

  body.appendParagraph("Resumo da IA").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(dados.resumoIA || "Indisponível.").editAsText().setFontSize(10.5);

  body.appendParagraph("Registo Completo de Eventos").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const eventos = dados.eventos || [];
  if (eventos.length) {
    estilizarTabelaInfo(body.appendTable(
      eventos.map(function (ev) { return ["[" + ev.ts + "] " + ev.tipo, ev.descricao]; })
    ));
  } else {
    body.appendParagraph("Sem eventos registados.");
  }

  if (fotos.length) {
    body.appendPageBreak();
    body.appendParagraph("Fotografias de Evidência").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    fotos.forEach(function (foto, i) {
      try {
        const dataUrl = (foto && foto.dataUrl) ? foto.dataUrl : foto;
        const blob = base64ParaBlob(dataUrl, "evidencia_" + (i + 1) + ".jpg");
        const img = body.appendImage(blob);
        img.setWidth(280);
        img.setHeight(210); // as fotos são sempre capturadas em 4:3 (320x240)
        const ts = (foto && foto.ts) ? foto.ts : null;
        const legenda = body.appendParagraph("Evidência " + (i + 1) + (ts ? " — capturada aos " + ts + " de sessão" : ""));
        legenda.editAsText().setForegroundColor("#838ba0").setItalic(true).setFontSize(9);
        legenda.setSpacingAfter(10);
      } catch (err) { /* imagem inválida/corrompida — ignora e continua */ }
    });
  }

  rodapeMarca(body);
  return docParaPdfEDescartar(doc, "relatorio_vigilancia_" + nomeArquivoSeguro(dados.nome) + ".pdf");
}

// ══════════════════════════════════════════════
// PDF 2 — FOLHA DE RESPOSTAS (para correção manual)
// ══════════════════════════════════════════════
function gerarPdfFolhaRespostas(dados) {
  if (!dados.perguntas || !dados.perguntas.length) return null;

  const doc = DocumentApp.create("tmp_folha_" + Date.now());
  const body = doc.getBody();
  body.setMarginTop(40).setMarginBottom(40).setMarginLeft(50).setMarginRight(50);

  cabecalhoDocumento(body, dados.provaTitulo || "Folha de Respostas");

  estilizarTabelaInfo(body.appendTable([
    ["Candidato", dados.nome || "—"],
    ["Número", dados.numero || "—"],
    ["Email", dados.emailAluno || "—"],
    ["Índice de Integridade", dados.confianca + "%"]
  ]));
  body.appendParagraph("").setSpacingAfter(4);

  dados.perguntas.forEach(function (p, i) {
    const pergPar = body.appendParagraph((i + 1) + ". " + p.pergunta);
    pergPar.setSpacingBefore(16);
    pergPar.editAsText().setBold(true).setFontSize(12).setForegroundColor("#0f172a");

    if (p.imagem) {
      try {
        const blobPergunta = base64ParaBlob(p.imagem, "pergunta_" + (i + 1) + ".jpg");
        const imgPergunta = body.appendImage(blobPergunta);
        const larguraOriginal = imgPergunta.getWidth();
        const alturaOriginal = imgPergunta.getHeight();
        const larguraAlvo = 320;
        imgPergunta.setWidth(larguraAlvo);
        imgPergunta.setHeight(Math.round(alturaOriginal * (larguraAlvo / larguraOriginal)));
      } catch (err) { /* imagem inválida — ignora e continua */ }
    }

    const resposta = (dados.respostas && dados.respostas[i] != null && dados.respostas[i] !== "")
      ? String(dados.respostas[i])
      : "(sem resposta)";

    const tabelaResp = body.appendTable([[resposta]]);
    const cel = tabelaResp.getRow(0).getCell(0);
    cel.setBackgroundColor("#f5f6f8");
    cel.editAsText().setFontSize(11).setForegroundColor("#0f172a");
  });

  rodapeMarca(body);
  return docParaPdfEDescartar(doc, "folha_respostas_" + nomeArquivoSeguro(dados.nome) + ".pdf");
}

// ══════════════════════════════════════════════
// EMAIL — mensagem curta + os dois PDFs em anexo
// ══════════════════════════════════════════════
function enviarEmailRelatorio(dados) {
  const est = dados.estatisticas || {};
  const porTipo = est.eventosPorTipo || {};

  const pdfVigilancia = gerarPdfVigilancia(dados);
  const pdfRespostas = gerarPdfFolhaRespostas(dados);
  const anexos = pdfRespostas ? [pdfRespostas, pdfVigilancia] : [pdfVigilancia];

  const html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;">' +
      '<div style="background:#1a56db;padding:20px 24px;border-radius:10px 10px 0 0;">' +
        '<h2 style="color:#fff;margin:0;">Virtus-Detector' + (dados.provaTitulo ? " — " + dados.provaTitulo : "") + '</h2>' +
        '<p style="color:#98a0b3;margin:6px 0 0;font-size:13px;">' + new Date().toLocaleString("pt-PT") + '</p>' +
      '</div>' +
      '<div style="border:1px solid #e3e3e3;border-top:none;padding:24px;border-radius:0 0 10px 10px;">' +
        '<p style="font-size:14px;color:#333;">Sessão de <strong>' + (dados.nome || "—") + '</strong> (Nº ' + (dados.numero || "—") + ') concluída.</p>' +
        '<div style="background:' + corIndice(dados.confianca) + ';color:#fff;padding:14px 18px;border-radius:8px;margin:16px 0;">' +
          '<div style="font-size:13px;opacity:0.9;">Índice de Integridade Final</div>' +
          '<div style="font-size:32px;font-weight:700;">' + dados.confianca + '%</div>' +
        '</div>' +
        '<table style="width:100%;border-collapse:collapse;">' +
          '<tr><td style="padding:4px 16px 4px 0;color:#666;font-size:13px;">Duração da sessão</td><td style="padding:4px 0;font-weight:600;font-size:13px;">' + (est.duracaoFormatada || "—") + '</td></tr>' +
          '<tr><td style="padding:4px 16px 4px 0;color:#666;font-size:13px;">Alertas</td><td style="padding:4px 0;font-weight:600;font-size:13px;">' + (porTipo.ALERTA || 0) + '</td></tr>' +
          '<tr><td style="padding:4px 16px 4px 0;color:#666;font-size:13px;">Ocorrências graves</td><td style="padding:4px 0;font-weight:600;font-size:13px;">' + (porTipo.GRAVE || 0) + '</td></tr>' +
        '</table>' +
        '<p style="font-size:13px;color:#666;margin-top:18px;">📎 Em anexo: ' +
          (pdfRespostas ? 'a <strong>folha de respostas</strong> (para correção) e o ' : 'o ') +
          '<strong>relatório de vigilância</strong> em PDF, com as fotografias de evidência.</p>' +
        '<p style="font-size:11px;color:#aaa;margin-top:20px;">' + ASSINATURA + '</p>' +
      '</div>' +
    '</div>';

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
  const sheet = obterOuCriarFolha(ss, FOLHA_RELATORIOS, ["Data", "Prova", "Nome", "Número", "Email", "Índice", "Duração", "Alertas", "Graves", "Observações"]);
  const est = dados.estatisticas || {};
  const porTipo = est.eventosPorTipo || {};
  sheet.appendRow([
    new Date(),
    dados.provaTitulo || "",
    dados.nome,
    dados.numero,
    dados.emailAluno || "",
    dados.confianca,
    est.duracaoFormatada || "",
    porTipo.ALERTA || 0,
    porTipo.GRAVE || 0,
    dados.observacoes || ""
  ]);
}

function registarRespostas(dados) {
  const ss = obterBaseDados();
  const sheet = obterOuCriarFolha(ss, FOLHA_RESPOSTAS, ["Data", "Prova", "Nome", "Número", "Email", "Pergunta", "Resposta"]);
  dados.perguntas.forEach(function (p, i) {
    const resposta = (dados.respostas && dados.respostas[i] != null) ? dados.respostas[i] : "";
    sheet.appendRow([new Date(), dados.provaTitulo || "", dados.nome, dados.numero, dados.emailAluno || "", p.pergunta, resposta]);
  });
}
