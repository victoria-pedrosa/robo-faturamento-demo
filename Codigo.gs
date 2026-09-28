/**
 * ROBÔ CONTHABIL - VERSÃO V44.0
 * Pasta raiz: definida em Propriedades do script (ID_PASTA_RAIZ), uma por squad
 * Estrutura: CODIGO-NOME / MMAAAA / PRESTADO + TOMADO
 */

const ID_PASTA_RAIZ = PropertiesService.getScriptProperties().getProperty('ID_PASTA_RAIZ');

// ==========================================
// MENU
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 ROBÔ CONTHABIL')
    .addItem('👉 Processar Faturamento', 'EXECUTAR_POR_SELECAO')
    .addSeparator()
    .addItem('🕵️ Detetive (Teste Unitário)', 'MODO_DETETIVE')
    .addItem('🔍 Listar Pastas da Raiz (Debug)', 'LISTAR_PASTAS_RAIZ')
    .addToUi();
}

// ==========================================
// EXECUÇÃO PRINCIPAL
// ==========================================
function EXECUTAR_POR_SELECAO() {
  const aba = SpreadsheetApp.getActiveSheet();
  const range = aba.getSelection().getActiveRange();
  if (!range) { Browser.msgBox("Selecione as linhas primeiro."); return; }

  let linhaInicial = range.getRow();
  let numLinhas    = range.getNumRows();

  aba.getRange(linhaInicial, 7, numLinhas, 4).setBackground("#f3f3f3").clearContent();
  SpreadsheetApp.flush();

  let contador = 0;

  for (let i = 0; i < numLinhas; i++) {
    let linhaAtual   = linhaInicial + i;
    let nomeEmpresa  = aba.getRange(linhaAtual, 1).getValue();
    let codigo       = aba.getRange(linhaAtual, 4).getValue();

    let celFaturamento = aba.getRange(linhaAtual, 7);
    let celTomados     = aba.getRange(linhaAtual, 8);
    let celIss         = aba.getRange(linhaAtual, 9);
    let celStatus      = aba.getRange(linhaAtual, 10);

    if (!codigo) continue;

    try {
      let res = processarEmpresa(String(codigo).trim(), nomeEmpresa, aba);

      // STATUS DA PASTA
      if (res.erroPasta) {
        celStatus.setValue("SEM PASTA")
          .setNote(res.erroMsg)
          .setFontColor("white").setBackground("red").setFontWeight("bold");
        celFaturamento.setValue("-"); celTomados.setValue("-"); celIss.setValue("-");
        continue;
      }

      if (!res.mesEncontrado) {
        celStatus.setValue("MÊS NÃO ENCONTRADO")
          .setNote("Pasta " + res.nomePasta + " existe, mas não tem a subpasta: " + res.nomePastaMesProcurado)
          .setFontColor("black").setBackground("yellow").setFontWeight("bold");
        celFaturamento.setValue("-"); celTomados.setValue("-"); celIss.setValue("-");
        continue;
      }

      celStatus.setValue("OK")
        .setNote("Pasta: " + res.nomePasta + " > " + res.nomePastaMesProcurado)
        .setFontColor("green").setBackground(null).setFontWeight("normal");

      // FATURAMENTO
      if (res.prestado.leuArquivo) {
        celFaturamento.setValue(res.prestado.valorTotal)
          .setNumberFormat('_-R$* #.##0,00_-')
          .setBackground("#d9ead3");
      } else {
        celFaturamento.setValue(0)
          .setNote("PDF 'print_nfse' não encontrado ou valor não extraído.");
      }

      // TOMADOS
      if (!res.tomado.leuArquivo) {
        celTomados.setValue("SEM CSV TOMADOS")
          .setFontColor("orange").setFontWeight("bold").setBackground("#fff2cc");
      } else {
        celTomados.setValue(res.tomado.temRetencao ? "COM RETENÇÃO" : "SEM RETENÇÃO")
          .setFontColor(res.tomado.temRetencao ? "red" : "green")
          .setFontWeight(res.tomado.temRetencao ? "bold" : "normal")
          .setBackground(null);
      }

      // GUIA ISS
      celIss.setValue(res.temGuiaISS ? "SIM" : "NÃO")
        .setFontColor(res.temGuiaISS ? "green" : "red")
        .setFontWeight(res.temGuiaISS ? "bold" : "normal");

      contador++;

    } catch (e) {
      celStatus.setValue("ERRO").setNote(e.message).setBackground("#fce8e6");
      Logger.log("ERRO linha " + linhaAtual + ": " + e.stack);
    }
  }

  Browser.msgBox("✅ Concluído! " + contador + " empresa(s) processada(s).\n\nConsulte o Log para detalhes.");
}

// ==========================================
// PROCESSAMENTO DE UMA EMPRESA
// ==========================================
function processarEmpresa(codigo, nomeEmpresa, aba) {
  let nomePastaMes = converterNomeMesParaPasta(aba.getName());

  Logger.log("\n══════════════════════════════════════");
  Logger.log("EMPRESA : " + nomeEmpresa);
  Logger.log("CÓDIGO  : " + codigo);
  Logger.log("ABA     : " + aba.getName() + "  →  PASTA MÊS: " + nomePastaMes);

  let retorno = {
    erroPasta: false, erroMsg: "",
    nomePasta: "", nomePastaMesProcurado: nomePastaMes,
    mesEncontrado: false,
    prestado:  { pastaEncontrada: false, leuArquivo: false, valorTotal: 0 },
    tomado:    { pastaEncontrada: false, leuArquivo: false, temRetencao: false },
    temGuiaISS: false
  };

  if (!nomePastaMes) {
    retorno.erroPasta = true;
    retorno.erroMsg = "Nome da aba '" + aba.getName() + "' não converteu para mês válido.";
    Logger.log("❌ " + retorno.erroMsg);
    return retorno;
  }

  // 1. ABRE A PASTA RAIZ
  let pastaRaiz;
  try {
    pastaRaiz = DriveApp.getFolderById(ID_PASTA_RAIZ);
    Logger.log("RAIZ    : " + pastaRaiz.getName());
  } catch(e) {
    retorno.erroPasta = true;
    retorno.erroMsg = "Não abriu a pasta raiz: " + e.message;
    return retorno;
  }

  // 2. BUSCA PASTA DA EMPRESA (formato: "28-EMPRESA EXEMPLO 103...")
  let pastaEmpresa = null;
  let iterEmpresas = pastaRaiz.getFolders();

  Logger.log("--- Procurando pasta da empresa ---");
  while (iterEmpresas.hasNext()) {
    let p    = iterEmpresas.next();
    let nome = p.getName().trim();

    // Extrai o código numérico antes do primeiro hífen ou espaço
    let matchCod = nome.match(/^(\d+)[\-\s]/);
    if (!matchCod) continue;

    let codPasta = matchCod[1];
    Logger.log("  [" + nome + "] → código: " + codPasta);

    if (codPasta === codigo) {
      pastaEmpresa = p;
      Logger.log("  ✅ MATCH!");
      break;
    }
  }

  if (!pastaEmpresa) {
    retorno.erroPasta = true;
    retorno.erroMsg   = "Pasta com código '" + codigo + "' não encontrada em PREFEITURA SALVADOR.";
    Logger.log("❌ " + retorno.erroMsg);
    return retorno;
  }
  retorno.nomePasta = pastaEmpresa.getName();

  // 3. BUSCA PASTA DO MÊS (formato: "032026")
  let pastaMes  = null;
  let iterMeses = pastaEmpresa.getFolders();

  Logger.log("--- Subpastas de " + retorno.nomePasta + " ---");
  while (iterMeses.hasNext()) {
    let s     = iterMeses.next();
    let sNome = s.getName().replace(/[\s\/\-]/g, "").toUpperCase();
    Logger.log("  [" + s.getName() + "] → normalizado: [" + sNome + "] | procurando: [" + nomePastaMes + "]");

    if (sNome === nomePastaMes) {
      pastaMes = s;
      Logger.log("  ✅ Pasta do mês encontrada!");
      break;
    }
  }

  if (!pastaMes) {
    Logger.log("❌ Pasta do mês [" + nomePastaMes + "] não encontrada.");
    return retorno;
  }
  retorno.mesEncontrado = true;

  // 4. PRESTADO (PDF OCR)
  let pastaPrestado = buscarSubPasta(pastaMes, "PRESTADO");
  Logger.log("PRESTADO: " + (pastaPrestado ? "✅ encontrada" : "❌ não encontrada"));
  if (pastaPrestado) {
    retorno.prestado.pastaEncontrada = true;
    let valor = extrairValorPdf(pastaPrestado, codigo);
    if (valor !== null) {
      retorno.prestado.leuArquivo = true;
      retorno.prestado.valorTotal = valor;
      Logger.log("  Valor PDF: R$ " + valor);
    }
  }

  // 5. TOMADO (CSV)
  let pastaTomado = buscarSubPasta(pastaMes, "TOMADO");
  Logger.log("TOMADO  : " + (pastaTomado ? "✅ encontrada" : "❌ não encontrada"));
  if (pastaTomado) {
    retorno.tomado.pastaEncontrada = true;
    let dados = lerCsvDaPastaTomado(pastaTomado);
    retorno.tomado.leuArquivo  = dados.leuArquivo;
    retorno.tomado.temRetencao = dados.temRetencao;
    Logger.log("  CSV lido: " + dados.leuArquivo + " | Retenção: " + dados.temRetencao);
  }

  // 6. GUIA ISS (PDF solto na pasta do mês, que NÃO seja print_nfse)
  let arquivosMes = pastaMes.getFilesByType(MimeType.PDF);
  while (arquivosMes.hasNext()) {
    let arq = arquivosMes.next();
    if (!arq.getName().toLowerCase().includes("print_nfse")) {
      retorno.temGuiaISS = true;
      Logger.log("GUIA ISS: ✅ " + arq.getName());
      break;
    }
  }

  return retorno;
}

// ==========================================
// OCR DO PDF
// ==========================================
function extrairValorPdf(pastaPrestado, codigo) {
  let arquivos   = pastaPrestado.getFilesByType(MimeType.PDF);
  let arquivoPDF = null;

  Logger.log("  PDFs na pasta PRESTADO:");
  while (arquivos.hasNext()) {
    let a = arquivos.next();
    Logger.log("    - " + a.getName());
    if (a.getName().toLowerCase().includes("ocr_temp")) continue;
    if (a.getName().toLowerCase().includes("print_nfse")) { arquivoPDF = a; break; }
  }

  if (!arquivoPDF) { Logger.log("  ❌ Nenhum print_nfse encontrado."); return null; }
  Logger.log("  ✅ PDF: " + arquivoPDF.getName());

  try {
    let blob    = arquivoPDF.getBlob();
    let recurso = { title: "TEMP_" + codigo, mimeType: MimeType.GOOGLE_DOCS };
    let criado  = Drive.Files.insert(recurso, blob, { convert: true });
    let idTemp  = criado.id;

    let textoOriginal = DocumentApp.openById(idTemp).getBody().getText();
    Drive.Files.remove(idTemp);

    Logger.log("  Trecho OCR: " + textoOriginal.substring(0, 400));

    let textoLimpo = textoOriginal.replace(/[\n\r"]/g, " ").replace(/\s+/g, " ");

    // Regex principal
    let m = textoLimpo.match(/Valor dos Servi(?:ç|c)os:.*?R\$\s*([\d\.,]+)/i);
    if (m && m[1]) {
      let v = parseFloat(m[1].replace(/\./g, '').replace(',', '.'));
      Logger.log("  ✅ Regex principal → R$ " + v);
      return v;
    }

    // Backup: maior valor numérico encontrado
    Logger.log("  ⚠️ Regex principal falhou. Usando backup...");
    let todos = textoOriginal.match(/(\d{1,3}(?:\.\d{3})*,\d{2})/g);
    let maior = 0;
    if (todos) {
      Logger.log("  Valores encontrados: " + todos.join(" | "));
      todos.forEach(t => {
        let n = parseFloat(t.replace(/\./g, '').replace(',', '.'));
        if (n > maior && n !== 2025 && n !== 2026 && n < 10000000) maior = n;
      });
    }
    Logger.log("  Backup → R$ " + maior);
    return maior > 0 ? maior : 0;

  } catch(e) {
    Logger.log("  ❌ Erro OCR: " + e.message);
    return null;
  }
}

// ==========================================
// CSV (TOMADOS)
// ==========================================
function lerCsvDaPastaTomado(pasta) {
  let output = { leuArquivo: false, temRetencao: false };
  let arquivos = pasta.getFiles();

  while (arquivos.hasNext()) {
    let arquivo = arquivos.next();
    if (!arquivo.getName().toLowerCase().endsWith(".csv") && arquivo.getMimeType() !== MimeType.CSV) continue;

    Logger.log("  CSV: " + arquivo.getName());
    output.leuArquivo = true;
    let dados = Utilities.parseCsv(arquivo.getBlob().getDataAsString('ISO-8859-1'), ';');
    if (dados.length < 2) continue;

    let header = dados[0];
    let colSit = -1, colIss = -1, colsFed = [];

    for (let h = 0; h < header.length; h++) {
      if (!header[h]) continue;
      let n = header[h].toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, "").trim();
      if (n.includes("situacao"))   colSit = h;
      if (n.includes("iss retido")) colIss = h;
      if (["pis","cofins","inss","irpj","csll"].includes(n)) colsFed.push(h);
    }

    for (let r = 1; r < dados.length; r++) {
      let linha = dados[r];
      if (!linha || !linha.length) continue;
      if (colSit > -1 && linha[colSit]) {
        let sit = linha[colSit].toUpperCase();
        if (sit.includes("CANCELADA") || sit === "C") continue;
      }
      if (colIss > -1 && linha[colIss]) {
        let v = linha[colIss].toString().toUpperCase();
        if (v === "S" || v === "SIM" || v === "1") output.temRetencao = true;
      }
      for (let idx of colsFed) {
        if (linha[idx] && parseBrFloat(linha[idx]) > 0) output.temRetencao = true;
      }
    }
  }
  return output;
}

// ==========================================
// MODO DETETIVE
// ==========================================
function MODO_DETETIVE() {
  const aba   = SpreadsheetApp.getActiveSheet();
  const linha = aba.getSelection().getActiveRange().getRow();

  let nomeEmpresa = aba.getRange(linha, 1).getValue();
  let codigo      = String(aba.getRange(linha, 4).getValue()).trim();

  if (!codigo) { Browser.msgBox("Coluna D sem código."); return; }

  Browser.msgBox("🔍 Iniciando Detetive...\n\nEmpresa : " + nomeEmpresa + "\nCódigo  : " + codigo + "\n\nAcompanhe em:\nExtensões → Apps Script → Execuções");

  try {
    let res = processarEmpresa(codigo, nomeEmpresa, aba);

    let msg = "🕵️ RESULTADO DO DETETIVE\n";
    msg    += "──────────────────────────\n";
    msg    += "Empresa : " + nomeEmpresa + "\n";
    msg    += "Código  : " + codigo + "\n\n";

    if (res.erroPasta) {
      msg += "❌ PASTA NÃO ENCONTRADA\n" + res.erroMsg;
    } else {
      msg += "✅ Pasta   : " + res.nomePasta + "\n";
      msg += "📅 Mês     : " + res.nomePastaMesProcurado + " → " + (res.mesEncontrado ? "✅ ENCONTRADO" : "❌ NÃO ENCONTRADO") + "\n\n";
      if (res.mesEncontrado) {
        msg += "📄 PRESTADO  : pasta=" + (res.prestado.pastaEncontrada?"✅":"❌") + " | pdf=" + (res.prestado.leuArquivo?"✅":"❌") + " | R$ " + res.prestado.valorTotal.toFixed(2) + "\n";
        msg += "📥 TOMADO    : pasta=" + (res.tomado.pastaEncontrada?"✅":"❌") + " | csv=" + (res.tomado.leuArquivo?"✅":"❌") + " | retenção=" + (res.tomado.temRetencao?"⚠️SIM":"✅NÃO") + "\n";
        msg += "📑 Guia ISS  : " + (res.temGuiaISS ? "✅ SIM" : "❌ NÃO") + "\n";
      }
    }

    msg += "\n📋 Log completo:\nExtensões → Apps Script → Execuções";
    Browser.msgBox(msg);

  } catch(e) {
    Browser.msgBox("❌ Erro: " + e.message);
    Logger.log("ERRO DETETIVE: " + e.stack);
  }
}

// ==========================================
// DEBUG: LISTA PASTAS DA RAIZ
// ==========================================
function LISTAR_PASTAS_RAIZ() {
  try {
    let raiz  = DriveApp.getFolderById(ID_PASTA_RAIZ);
    let iter  = raiz.getFolders();
    let lista = "📂 " + raiz.getName() + "\n\n";
    let n = 0;
    while (iter.hasNext() && n < 40) {
      lista += "• " + iter.next().getName() + "\n";
      n++;
    }
    if (n === 0) lista += "(sem subpastas)";
    Browser.msgBox(lista);
  } catch(e) {
    Browser.msgBox("❌ " + e.message);
  }
}

// ==========================================
// UTILITÁRIOS
// ==========================================
function converterNomeMesParaPasta(nomeAba) {
  try {
    let limpo = nomeAba.trim();

    // Formato "03/2026" ou "3/2026"
    let m1 = limpo.match(/^(\d{1,2})[\/\-](\d{4})$/);
    if (m1) return m1[1].padStart(2,"0") + m1[2];

    // Formato "032026" já pronto
    let m2 = limpo.match(/^(\d{2})(\d{4})$/);
    if (m2) return m2[1] + m2[2];

    // Formato "Março/2026" ou "março/2026"
    let partes = limpo.split("/");
    if (partes.length < 2) return null;

    let mesNome = partes[0].trim().toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, "");
    let ano = partes[1].trim();

    const meses = {
      "janeiro":"01","fevereiro":"02","marco":"03",
      "abril":"04","maio":"05","junho":"06",
      "julho":"07","agosto":"08","setembro":"09",
      "outubro":"10","novembro":"11","dezembro":"12"
    };

    let num = meses[mesNome];
    if (!num) { Logger.log("Mês não reconhecido: [" + mesNome + "]"); return null; }
    return num + ano;

  } catch(e) { Logger.log("Erro converterNomeMes: " + e.message); return null; }
}

function buscarSubPasta(pastaPai, termoChave) {
  let iter = pastaPai.getFolders();
  while (iter.hasNext()) {
    let p = iter.next();
    if (p.getName().toUpperCase().includes(termoChave.toUpperCase())) return p;
  }
  return null;
}

function parseBrFloat(str) {
  if (!str) return 0;
  let n = parseFloat(str.toString().replace(/\./g,"").replace(",","."));
  return isNaN(n) ? 0 : n;
}