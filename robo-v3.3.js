// =======================================================
// ROBO ASSIM → ORBITA & SUPABASE
// Versão: 3.4
// =======================================================

// =========================
// IMPORTS
// =========================
require('dotenv').config();
const { chromium } = require('playwright');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

function dentroDoHorario() {
  const agora = new Date();
  const horaBR = new Date(
    agora.toLocaleString("en-US", { timeZone: "America/Sao_Paulo" })
  );

  const dia = horaBR.getDay();
  const hora = horaBR.getHours();
  const minuto = horaBR.getMinutes();

  if (dia === 0 || dia === 6) return false;
  if (hora < 8) return false;
  if (hora === 8 && minuto < 30) return false;
  if (hora >= 18) return false;

  return true;
}

// ================
// FUNÇÃO LOG
// ================
function log(tipo, mensagem) {
  const agora = new Date();

  const data = agora.toLocaleDateString('pt-BR', {
    timeZone: 'America/Sao_Paulo'
  });

  const hora = agora.toLocaleTimeString('pt-BR', {
    timeZone: 'America/Sao_Paulo',
    hour12: false
  });

  const [dia, mes, ano] = data.split('/');
  const dataFormatada = `${ano}-${mes}-${dia}`;

  const emojis = {
    INFO: '🔵',
    SUCCESS: '✅',
    ERROR: '❌',
    WARNING: '⚠️'
  };

  const emoji = emojis[tipo] || '🔹';
  const linha = `[${dataFormatada} ${hora}] [${tipo}] ${emoji} ${mensagem}\n`;

  const pastaLogs = path.join(__dirname, 'logs');

  if (!fs.existsSync(pastaLogs)) {
    fs.mkdirSync(pastaLogs, { recursive: true });
  }

  const caminhoLog = path.join(pastaLogs, `log-${dataFormatada}.txt`);

  fs.appendFileSync(caminhoLog, linha);
  console.log(linha.trim());
}

// =====================
// REMOVER DUPLICADOS
// =====================
function removerDuplicadosPorGuia(registros) {
  const mapa = new Map();

  registros.forEach(r => {
    if (r.guia) {
      mapa.set(r.guia, r); // mantém o último (mais atualizado)
    }
  });

  return Array.from(mapa.values());
}

// =====================
// FUNÇÃO DATA PRO BANCO
// =====================
function converterData(dataStr) {
  if (!dataStr) return null;

  const partes = dataStr.split(' ');
  if (partes.length < 2) return null;

  const [data, hora] = partes;
  const partesData = data.split('/');
  if (partesData.length < 3) return null;

  const [dia, mes, ano] = partesData;

  return `${ano}-${mes}-${dia}T${hora}`;
}

// =====================
// FUNÇÃO STATUS
// =====================
function extrairStatus(r) {
  const texto = r["Codigo        Status     Sit"] || '';
  
  if (texto.includes("NAO AUTORIZADO")) return "NEGADO";
  if (texto.includes("NEGADO")) return "NEGADO";
  if (texto.includes("ERRO")) return "ERRO";
  if (texto.includes("AUTORIZADO")) return "AUTORIZADO";

  return "DESCONHECIDO";
}

function normalizarStatus(status) {
  if (!status) return null;

  const texto = status.toUpperCase();

  // ❌ NEGADOS
  if (texto.includes("NAO AUTORIZADO")) return "NEGADO";
  if (texto.includes("NEGADO")) return "NEGADO";

  // ❌ ERROS
  if (texto.includes("ERRO")) return "ERRO";

  // ✅ LIBERADO (EQUIVALE A AUTORIZADO)
  if (texto.includes("LIBERADO")) return "AUTORIZADO";

  // fallback antigo (mantém por segurança)
  if (texto.includes("AUTORIZADO")) return "AUTORIZADO";

  log("WARNING", `⚠️ Status não reconhecido: "${status}"`);

  return "DESCONHECIDO";
}

// ================
// FUNÇÃO DIVISÓRIA
// ================
function logDivisoria(titulo = '') {
  const agora = new Date();

  const data = agora.toLocaleDateString('pt-BR', {
    timeZone: 'America/Sao_Paulo'
  });

  const [dia, mes, ano] = data.split('/');
  const dataFormatada = `${ano}-${mes}-${dia}`;

  const pastaLogs = path.join(__dirname, 'logs');

  if (!fs.existsSync(pastaLogs)) {
    fs.mkdirSync(pastaLogs, { recursive: true });
  }

  const caminhoLog = path.join(pastaLogs, `log-${dataFormatada}.txt`);

  let bloco = `\n============================================================\n`;

  if (titulo) {
    bloco += `${titulo}\n`;
    bloco += `============================================================\n`;
  }

  fs.appendFileSync(caminhoLog, bloco);
  console.log(bloco.trim());
}


// ========================
// FUNÇÃO AJUSTAR Supabase
// ========================
function transformarParaSupabase(registros) {
  return registros.map(r => {
    const linhas = (r.beneficiario || "")
      .split('\n')
      .map(l => l.trim())
      .filter(Boolean);

    const matricula = linhas[0] || null;
    const nome = linhas[1] || linhas[0] || null;

    return {
      guia: r.guia?.trim() || null,
      matricula,
      paciente_nome: nome,

      data_execucao: converterData(r.dataHora),
      data_autorizacao: null,

      status: r.status?.trim() || null,
      codigo_tuss: r.codigo?.trim() || null,

      codigo_erro: null,
      descricao_erro: null,

      teve_token: !!(r.token && r.token.trim()),
      token: r.token?.trim() || null,

      updated_at: new Date().toLocaleString('sv-SE', {
        timeZone: 'America/Sao_Paulo'
      }).replace(' ', 'T'),

      biofacial: r._biofacial?.trim() || null
    };
  });
}

// ========================
// FUNÇÃO TEMPO DE ATIVAÇÃO
// ========================
function tempo(inicio) {
  return ((Date.now() - inicio) / 1000).toFixed(2) + "s";
}

function lerStatus() {
  const caminho = path.join(__dirname, 'state', 'status-site.json');
  if (!fs.existsSync(caminho)) return null;

  try {
    return JSON.parse(fs.readFileSync(caminho, 'utf-8')).status;
  } catch (erro) {
    log("ERROR", "📡 Erro ao ler status-site.json");
    return null;
  }
}

function salvarStatus(status) {
  const pasta = path.join(__dirname, 'state');
  if (!fs.existsSync(pasta)) {
    fs.mkdirSync(pasta, { recursive: true });
  }

  const caminho = path.join(pasta, 'status-site.json');
  fs.writeFileSync(caminho, JSON.stringify({ status }, null, 2));
}

async function enviarSlack(mensagem) {
  const webhook = process.env.SLACK_WEBHOOK;
  if (!webhook) {
    log("ERROR", "⚙️ SLACK_WEBHOOK não definida");
    return;
  }

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 5000);

    await fetch(webhook, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ text: mensagem }),
      signal: controller.signal
    });

    clearTimeout(timeout);
    log("INFO", "💬 Mensagem enviada ao Slack");
  } catch (erro) {
    if (erro.name === 'AbortError') {
      log("ERROR", "⏱️ Timeout ao enviar mensagem para Slack");
    } else {
      log("ERROR", "❌💬 Erro ao enviar mensagem para Slack");
      log("ERROR", `❌${erro.message}`);
    }
  }
}

async function acessarComRetry(page, url, tentativas = 3) {
  for (let i = 1; i <= tentativas; i++) {
    try {
      log("INFO", `🌐 Tentativa ${i} de acesso ao site Autorizador da ASSIM...`);

      await page.goto(url, {
        waitUntil: 'domcontentloaded',
        timeout: 15000
      });

      log("SUCCESS", "🌐 Acesso realizado com sucesso");
      return true;
    } catch (erro) {
      log("ERROR", `🔁 Erro na tentativa ${i}: ${erro.message}`);

      if (i === tentativas) {
        log("ERROR", "🌐 Falha total ao acessar o site");
        return false;
      }

      log("INFO", "🕒 Aguardando 5s para nova tentativa...");
      await page.waitForTimeout(5000);
    }
  }
}

// =========================
// EXTRAIR RELATÓRIO
// =========================
async function extrairRelatorio(page, urlConsulta) {
  try {
    await page.goto(urlConsulta, {
      waitUntil: 'domcontentloaded',
      timeout: 60000
    });

    let currentUrl = page.url();

    if (currentUrl.includes('preresultado')) {
      const urlSemPaginacao = currentUrl.replace('preresultado', 'resultadosempaginacao');
      log("INFO", `🔀 Fluxo antigo: redirecionando para resultadosempaginacao`);
      await page.goto(urlSemPaginacao, {
        waitUntil: 'domcontentloaded',
        timeout: 60000
      });
      currentUrl = page.url();
    } else if (currentUrl.includes('relatorio.csp')) {
      log("INFO", `🆕 Novo fluxo detectado: relatorio.csp com CSPToken — permanecendo na página`);
    } else {
      log("INFO", `🔍 URL não reconhecida no fluxo esperado — continuando na URL atual`);
    }

    await page.waitForSelector('pre, table', { timeout: 120000 });

    let extractionContext = page;

    for (const f of page.frames()) {
      const hasPre = await f.locator('pre').count().catch(() => 0);
      if (hasPre > 0) {
        extractionContext = f;
        break;
      }
    }

    const resultado = await extractionContext.evaluate(() => {
      // --- Localizar a linha de cabeçalho real das colunas ---
      // Ignora linhas de título (Hospital, Periodo, Executor) que ficam antes
      const allTrs = Array.from(document.querySelectorAll('tr'));
      let columnHeaderRow = null;

      for (const tr of allTrs) {
        const ths = Array.from(tr.querySelectorAll('th'));
        const texts = ths.map(th => (th.innerText || '').trim().toLowerCase());
        // A linha de cabeçalho real contém 'guia' junto com 'sv' ou 'nat'
        if (texts.some(t => t === 'guia') && texts.some(t => t === 'sv' || t === 'nat')) {
          columnHeaderRow = tr;
          break;
        }
      }

      // Fallback: primeira linha com mais de 8 th elementos
      if (!columnHeaderRow) {
        for (const tr of allTrs) {
          if (tr.querySelectorAll('th').length > 8) {
            columnHeaderRow = tr;
            break;
          }
        }
      }

      const headers = columnHeaderRow
        ? Array.from(columnHeaderRow.querySelectorAll('th')).map(th => (th.innerText || '').trim())
        : [];

      // Mapa baseado APENAS na linha de cabeçalho real → índice = posição do td nas linhas de dados
      const headerMap = {};
      headers.forEach((h, i) => {
        headerMap[h.toLowerCase().replace(/\s+/g, ' ')] = i;
      });

      function resolverCol(...candidatos) {
        for (const c of candidatos) {
          const idx = headerMap[c.toLowerCase().replace(/\s+/g, ' ')];
          if (idx !== undefined) return idx;
        }
        return -1;
      }

      const colDataHora  = resolverCol('data hora', 'data/hora', 'data', 'dt/hr', 'dt.hr');
      const colSv        = resolverCol('sv', 's.v.', 'serviço', 'servico');
      const colNat       = resolverCol('nat', 'nat.', 'natureza');
      const colBenefic   = resolverCol('beneficiário', 'beneficiario', 'paciente', 'segurado', 'benefic.');
      const colBiofacial = resolverCol('biofacial', 'bio facial', 'biometria facial');
      const colToken     = resolverCol('token', 'senha');
      const colJustif    = resolverCol('justificativa', 'justif.', 'justif');
      const colProcesso  = resolverCol('processo', 'n° processo', 'n.processo', 'nº processo');
      const colGuia      = resolverCol('guia', 'n° guia', 'n.guia', 'nº guia');
      const colSoli      = resolverCol('soli', 'solicitante', 'médico', 'medico');
      const colEspec     = resolverCol('especialidade', 'espec.', 'espec');

      // Codigo e Status não têm <th> próprio — ficam logo após Especialidade
      const colCodigo = colEspec >= 0 ? colEspec + 1 : 11;
      const colStatus = colEspec >= 0 ? colEspec + 2 : 12;

      const mapeamento = {
        colDataHora, colSv, colNat, colBenefic, colBiofacial, colToken, colJustif,
        colProcesso, colGuia, colSoli, colEspec, colCodigo, colStatus,
        totalHeadersEncontrados: headers.length
      };

      const linhas = Array.from(document.querySelectorAll('tr'));
      const dados = [];
      let ultimoRegistro = {};

      linhas.forEach(linha => {
        const td = linha.querySelectorAll('td');
        if (td.length === 0) return;

        function getText(idxMapeado, idxFallback) {
          const idx = idxMapeado >= 0 ? idxMapeado : idxFallback;
          if (idx < 0 || idx >= td.length) return undefined;
          return td[idx]?.innerText?.trim() || undefined;
        }

        const dataHora        = getText(colDataHora, 0);
        const sv              = getText(colSv, 1);
        const nat             = getText(colNat, 2);
        const beneficiarioRaw = getText(colBenefic, 3) || "";

        const partes = beneficiarioRaw.split('\n').map(p => p.trim()).filter(Boolean);
        const matricula    = partes[0] || null;
        const beneficiario = partes[1] || partes[0] || null;

        // fallbacks atualizados: Biofacial ocupa posição 4, Token desloca para 5
        const biofacial     = getText(colBiofacial, 4);
        const token         = getText(colToken, 5);
        const justificativa = getText(colJustif, 6);
        const processo      = getText(colProcesso, 7);
        const guia          = getText(colGuia, 8);
        if (!guia || !/^\d+$/.test(guia)) return;
        const soli          = getText(colSoli, 9);
        const especialidade = getText(colEspec, 10);

        const codigoStatusRaw = getText(colCodigo, 11) || "";
        const partesCodigo = codigoStatusRaw.split(/\s+/);
        const codigo = partesCodigo[0] || null;

        let status = getText(colStatus, 12);
        if (!status) {
          status = partesCodigo.slice(1).join(' ') || null;
        }

        const registro = {
          dataHora:      dataHora      ?? ultimoRegistro.dataHora,
          sv:            sv            ?? ultimoRegistro.sv,
          nat:           nat           ?? ultimoRegistro.nat,
          matricula:     matricula     ?? ultimoRegistro.matricula,
          beneficiario:  beneficiario  ?? ultimoRegistro.beneficiario,
          biofacial:     biofacial,
          token:         token,
          justificativa: justificativa ?? ultimoRegistro.justificativa,
          processo:      processo      ?? ultimoRegistro.processo,
          guia:          guia          ?? ultimoRegistro.guia,
          soli:          soli          ?? ultimoRegistro.soli,
          especialidade: especialidade ?? ultimoRegistro.especialidade,
          codigo,
          status
        };

        dados.push(registro);
        ultimoRegistro = registro;
      });

      return { dados, headers, mapeamento };
    });

    log("INFO", `📋 Colunas detectadas na tabela: ${resultado.headers.length}`);

    const registrosFiltrados = resultado.dados.filter(r =>
      !((r.matricula || '').trim() === '' && (r.beneficiario || '').trim() === '')
    );

    function normalizarDataHora(dataHora) {
      if (!dataHora) return undefined;

      if (dataHora.match(/\d{2}\/\d{2}\/\d{4}/)) {
        if (dataHora.length === 16) {
          return dataHora + ":00";
        }
        return dataHora;
      }

      const hoje = new Date();
      const ano = hoje.getFullYear();
      let novaData = `${dataHora.slice(0, 5)}/${ano} ${dataHora.slice(6)}`;

      if (novaData.length === 16) {
        novaData += ":00";
      }

      return novaData;
    }

    const registrosTratados = registrosFiltrados.map(r => {
      const matricula = (r.matricula || '').trim();
      const beneficiario = (r.beneficiario || '').trim();

      let nomeCompleto;
      if (matricula || beneficiario) {
        nomeCompleto = `${matricula}\n${beneficiario}`.trim();
      }

      const codigo = (r.codigo || '').trim();
      const status = (r.status || '').trim();
      const sit = (r.sit || '').trim();

      let codigoStatusSit;
      if (codigo || status || sit) {
        codigoStatusSit = `${codigo.padEnd(13)}${status.padEnd(13)}${sit}`;
      }

      return {
        dataHora: normalizarDataHora(r.dataHora),
        sv: r.sv || undefined,
        nat: r.nat || undefined,
        beneficiario: nomeCompleto || undefined,
        token: r.token || undefined,
        justificativa: r.justificativa || undefined,
        processo: r.processo || undefined,
        guia: r.guia || undefined,
        soli: r.soli || undefined,
        especialidade: r.especialidade || undefined,
        codigo: codigo || undefined,
        status: status || undefined,
        "Codigo        Status     Sit": codigoStatusSit || undefined,

        // campo interno: vai para o Supabase mas NÃO para o Excel do Órbita
        _biofacial: r.biofacial || undefined
      };
    });

    return registrosTratados;
  } catch (erro) {
    log("ERROR", `📊 ERRO ao acessar relatório: ${urlConsulta}`);
    log("ERROR", `❌ Detalhe: ${erro.message}`);
    return [];
  }
}

// =========================
// API POST SUPABSE
// =========================
async function enviarParaSupabase(dados) {
  const url = process.env.SUPABASE_URL + '/rest/v1/autorizacoes_assim';
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'apikey': key,
        'Authorization': `Bearer ${key}`,
        'Content-Type': 'application/json',
        'Prefer': 'resolution=merge-duplicates'
      },
      body: JSON.stringify(dados)
    });

    if (response.ok) {
      log("SUCCESS", `📡 Supabase OK (${dados.length} registros)`);
    } else {
      const text = await response.text();
      log("ERROR", `❌ Supabase erro: ${response.status}`);
      log("ERROR", text);
    }

  } catch (erro) {
    log("ERROR", "❌ Erro ao enviar para Supabase");
    log("ERROR", erro.message);
  }
}

// ============================
// ENVIO PARA SUPABASE EM LOTES 
// ============================
async function enviarEmLotes(dados, tamanho = 100) {
  for (let i = 0; i < dados.length; i += tamanho) {
    const lote = dados.slice(i, i + tamanho);
    await enviarParaSupabase(lote);
  }
}

// =========================
// LOGIN ORBITA
// =========================
async function loginOrbita(page, usuario, senha) {
  await page.goto("https://cronogramauniversoaba.com.br/app_Login/");
  await page.fill('input[placeholder="Usuário"]', usuario);
  await page.fill('input[placeholder="Senha"]', senha);

  await Promise.all([
    page.waitForNavigation({ waitUntil: 'domcontentloaded' }),
    page.click('text=Entrar')
  ]);

  await page.waitForLoadState('networkidle');
}

// =========================
// UPLOAD EXCEL ORBITA
// =========================
async function enviarExcelOrbita(page, arquivoExcel, dataHoje) {

  // 🔥 acesso direto à página
  await page.goto('https://cronogramauniversoaba.com.br/blank_upload_registros_assim/', {
    waitUntil: 'networkidle'
  });

  // 🔥 upload
  await page.locator('input[type="file"]').setInputFiles(arquivoExcel);

  // 🔥 preencher datas (usando name - robusto)
  const [dia, mes, ano] = dataHoje.split('/');
  const dataISO = `${ano}-${mes}-${dia}`;
  
  await page.locator('input[name="data_inicial"]').fill(dataISO);
  await page.locator('input[name="data_final"]').fill(dataISO);

  // 🔥 botão carregar
  await page.getByRole('button', { name: 'Carregar, visualizar e linkar' }).click();

  // espera processamento real (melhor que só networkidle)
  await page.waitForLoadState('networkidle');
  await page.waitForTimeout(2000);

  // 🔥 botão confirmar (flexível)
  const botaoConfirmar = page.getByRole('button').filter({
    hasText: /confirmar|finalizar|processar/i
  }).first();

  if (await botaoConfirmar.isVisible()) {
    await botaoConfirmar.click({ force: true });
    log("SUCCESS", "🏁 Upload Órbita confirmado");
  } else {
    log("WARNING", "⚠️ Botão de confirmação não encontrado no Órbita");
  }

  await page.waitForTimeout(2000);
}

// =========================
// EXECUÇÃO PRINCIPAL
// =========================

(async () => {
  
  logDivisoria('🚀 INICIANDO NOVA EXECUÇÃO DO ROBÔ');
  log('INFO', '🕒 Iniciando verificação de rotina...');

  // if (!dentroDoHorario()) {
  //  log("INFO", "🕒 Fora do horário de execução (08:00 - 18:00). Encerrando.");
  //  process.exit(0);
  // }

  const atraso = 10000 + Math.random() * 20000;
  log("INFO", `⏳ Aguardando ${(atraso / 1000).toFixed(1)}s antes de iniciar...`);
  await new Promise(r => setTimeout(r, atraso));

  const inicioTotal = Date.now();

  const browser = await chromium.launch({
    headless: true
  });

  const context = await browser.newContext();
  const bloqueados = new Set(['image', 'font', 'stylesheet', 'media']);
  
  await context.route('**/*', route => {
    const tipo = route.request().resourceType();
    return bloqueados.has(tipo) ? route.abort() : route.continue();
  });

  const page = await context.newPage();

  const sucesso = await acessarComRetry(
    page,
    'https://sirius.assim.com.br/assimcsp/autorizador/login.csp'
  );

  const statusAnterior = await obterStatusRemoto();

  if (!sucesso) {
    if (statusAnterior !== "offline") {
      log("ERROR", "🚨 Site ficou OFFLINE");
      
      // ✅ CORRIGIDO: Adicionar aspas (template literal com backticks)
      await enviarSlack(
        `🔴 *INDISPONIBILIDADE DETECTADA*\nO site do Autorizador da Assim está FORA do ar.\n⏰ ${new Date().toLocaleString('pt-BR')}`
      );
    }

    await salvarStatusRemoto("offline");
    await browser.close();
    process.exit(1);
  }

  if (statusAnterior === "offline") {
    log("SUCCESS", "✅ Site voltou ao normal");
    
    // ✅ CORRIGIDO: Adicionar aspas (template literal com backticks)
    await enviarSlack(
      `🟢 *DISPONIBILIDADE RESTAURADA*\nO site do Autorizador da Assim voltou ao ar.\n⏰ ${new Date().toLocaleString('pt-BR')}`
    );
  }

  await salvarStatusRemoto("online");

  await page.selectOption('select', '52345');
  await page.fill('input[type="password"]', process.env.SENHA);

  await Promise.all([
    page.waitForNavigation(),
    page.click('text=Entrar')
  ]);

  await page.waitForSelector('select[name="DiaFim"]');

  const hoje = new Date();
  const dia = hoje.getDate().toString().padStart(2, '0');
  const mes = (hoje.getMonth() + 1).toString().padStart(2, '0');
  const ano = hoje.getFullYear();

  const dataHoje = `${dia}/${mes}/${ano}`;
  const dataArquivo = `${dia}-${mes}-${ano}`;

  const urlNormal =
    `https://sirius.assim.com.br/assimcsp/autorizador/preresultado.csp?idHospital=52345&DataIni=${dataHoje}&DataFim=${dataHoje}&executor=52345&natservico=T&servico=T&especialidade=T&amb=&prefeitura=0&tuss=`;
  const urlPrefeitura =
    `https://sirius.assim.com.br/assimcsp/autorizador/preresultado.csp?idHospital=52345&DataIni=${dataHoje}&DataFim=${dataHoje}&executor=52345&natservico=T&servico=T&especialidade=T&amb=&prefeitura=1&tuss=`;

  const registrosNormal     = await extrairRelatorio(page, urlNormal);
  const registrosPrefeitura = await extrairRelatorio(page, urlPrefeitura);

  log("INFO", `📋 Normal: ${registrosNormal.length} | Prefeitura: ${registrosPrefeitura.length}`);

  const registrosTodos = [...registrosNormal, ...registrosPrefeitura];

  const dadosBancoBruto = transformarParaSupabase(registrosTodos);
  const dadosBanco = removerDuplicadosPorGuia(dadosBancoBruto);

  log("INFO", `📦 Enviando ${dadosBanco.length} registros em lotes`);
  await enviarEmLotes(dadosBanco);

  if (dadosBanco.length === 0) {
    log("INFO", "📭 Nenhum dado encontrado. Pulando envio para Órbita.");
    await browser.close();
    process.exit(0);
  }

  const pastaRelatorios = path.join(__dirname, 'relatorios');
  if (!fs.existsSync(pastaRelatorios)) fs.mkdirSync(pastaRelatorios, { recursive: true });

  const nomeArquivo = `relatorio_assim_${dataArquivo}.xlsx`;
  const dados = registrosTodos.map(({ _biofacial, ...r }) => r);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dados), "Relatorio");
  const caminhoArquivo = path.join(pastaRelatorios, nomeArquivo);
  XLSX.writeFile(wb, caminhoArquivo);
  log("INFO", `📊 Excel gerado: ${nomeArquivo} (${registrosTodos.length} registros)`);

  await enviarRelatorioDrive(caminhoArquivo, nomeArquivo);

  const userOrbita = process.env.ORBITA_USER;
  const passOrbita = process.env.ORBITA_PASS;

  if (!userOrbita || !passOrbita) {
    log("ERROR", "🔐 Credenciais do Órbita não encontradas");
    await browser.close();
    process.exit(1);
  }

  await loginOrbita(page, userOrbita, passOrbita);

  log("INFO", `📤 Enviando relatório de hoje (${dataHoje}) para Órbita...`);
  await enviarExcelOrbita(page, caminhoArquivo, dataHoje);

  const pastaLogs = path.join(__dirname, 'logs');
  const arquivos = fs.readdirSync(pastaLogs);

  const ultimoLog = arquivos
    .filter(f => f.startsWith('log-'))
    .sort()
    .pop();

  if (ultimoLog) {
    const caminhoLog = path.join(pastaLogs, ultimoLog);
    await enviarLogDrive(caminhoLog, ultimoLog);
  } else {
    log("ERROR", "📂 Nenhum arquivo de log encontrado para envio");
  }

  log("SUCCESS", `🏁 Execução finalizada em ${tempo(inicioTotal)}`);

  await browser.close();
})();

// ===============================
// ENVIO DE RELATORIO PARA O DRIVE
// ===============================
async function enviarRelatorioDrive(caminhoArquivo, nomeArquivo) {
  try {
    const url = process.env.GOOGLE_SCRIPT_URL;

    if (!url) {
      log("ERROR", "☁️ GOOGLE_SCRIPT_URL não definida");
      return;
    }

    const fileBuffer = fs.readFileSync(caminhoArquivo);
    const base64 = fileBuffer.toString('base64');

    log("INFO", "☁️ Enviando relatório para o Drive...");

    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        type: "relatorio",
        fileName: nomeArquivo,
        fileContent: base64
      })
    });

    await response.text();
  } catch (erro) {
    log("ERROR", `☁️ Erro ao enviar relatório para o Drive: ${erro.message}`);
  }
}

// ===============================
// ENVIO DE LOG PARA O DRIVE
// ===============================
async function enviarLogDrive(caminhoLog, nomeArquivo) {
  try {
    const url = process.env.GOOGLE_SCRIPT_URL;
    const fileBuffer = fs.readFileSync(caminhoLog);
    const base64 = fileBuffer.toString('base64');

    log("INFO", "📁 ☁️ Enviando LOG para o Drive...");

    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        type: "log",
        fileName: nomeArquivo,
        fileContent: base64
      })
    });

    await response.text();
  } catch (erro) {
    log("ERROR", `☁️ Erro ao enviar LOG para o Drive: ${erro.message}`);
  }
}

// =============================
// LER O STATUS
// =============================
async function obterStatusRemoto() {
  for (let i = 1; i <= 2; i++) {
    try {
      const res = await fetch(process.env.GOOGLE_SCRIPT_URL);
      const data = await res.json();
      return data.status;
    } catch (erro) {
      if (i === 2) {
        log("ERROR", "📡 Erro ao obter status remoto");
        return "desconhecido";
      }

      await new Promise(r => setTimeout(r, 2000));
    }
  }
}

// =============================
// SALVAR STATUS
// =============================
async function salvarStatusRemoto(status) {
  try {
    const url = process.env.GOOGLE_SCRIPT_URL;

    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        type: "state",
        status
      })
    });

    await response.text();
  } catch (erro) {
    log("ERROR", `📡 Erro ao salvar status remoto: ${erro.message}`);
  }
}
