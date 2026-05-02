// =======================================================
// ROBO ASSIM → ORBITA & SUPABASE
// Versão: 3.4
// =======================================================

// =========================
// IMPORTS
// =========================
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
// FUNÇÃO AJUSTAR BANCO
// ========================
 function transformarParaBanco(registros) {
      return registros
        .filter(r => r.guia) // garante chave
        .map(r => {
          const [matricula, nome] = (r.beneficiario || '').split('\n');
    
          return {
            guia: r.guia,
            matricula: matricula || null,
            paciente_nome: nome || null,
            data_execucao: converterData(r.dataHora),
            status: extrairStatus(r),
            codigo_tuss: r.soli || null,
            codigo_erro: null,
            descricao_erro: r.justificativa || null,
            teve_token: !!(r.token && r.token.trim())
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

    const currentUrl = page.url();

    if (currentUrl.includes('preresultado')) {
      const url = currentUrl.replace('preresultado', 'resultadosempaginacao');

      await page.goto(url, {
        waitUntil: 'domcontentloaded',
        timeout: 60000
      });
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

    const registros = await extractionContext.evaluate(() => {
      const linhas = Array.from(document.querySelectorAll('tr'));
      const dados = [];
      let ultimoRegistro = {};

      linhas.forEach(linha => {
        const pre = linha.querySelector('pre');
        if (!pre) return;

        const td = linha.querySelectorAll('td');

        const dataHora = td[0]?.innerText?.trim();
        const sv = td[1]?.innerText?.trim();
        const nat = td[2]?.innerText?.trim();
        const beneficiarioCell = td[3]?.innerHTML;

        let matricula;
        let beneficiario;

        if (beneficiarioCell) {
          const partes = beneficiarioCell.split('<br>');
          matricula = partes[0]?.trim();
          beneficiario = partes[1]?.trim();
        }

        const token = td[4]?.innerText?.trim();
        const justificativa = td[5]?.innerText?.trim();
        const processo = td[6]?.innerText?.trim();
        const guia = td[7]?.innerText?.trim();
        const soli = td[8]?.innerText?.trim();
        const especialidade = td[9]?.innerText?.trim();

        const texto = pre.innerText.trim().split(/\s+/);
        const codigo = texto[0];
        const status = texto.slice(1).join(' ');

        const registro = {
          dataHora: dataHora ?? ultimoRegistro.dataHora,
          sv: sv ?? ultimoRegistro.sv,
          nat: nat ?? ultimoRegistro.nat,
          matricula: matricula ?? ultimoRegistro.matricula,
          beneficiario: beneficiario ?? ultimoRegistro.beneficiario,
          token: token ?? ultimoRegistro.token,
          justificativa: justificativa ?? ultimoRegistro.justificativa,
          processo: processo ?? ultimoRegistro.processo,
          guia: guia ?? ultimoRegistro.guia,
          soli: soli ?? ultimoRegistro.soli,
          especialidade: especialidade ?? ultimoRegistro.especialidade,
          codigo,
          status
        };

        dados.push(registro);
        ultimoRegistro = registro;
      });

      return dados;
    });

    const registrosFiltrados = registros.filter(r =>
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
        "Codigo        Status     Sit": codigoStatusSit || undefined
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
  const frame = page.frameLocator('#iframe_item_89');

  await page.mouse.move(5, 300);
  await page.locator('#nav_list > li:nth-child(2)').waitFor();
  await page.locator('#nav_list > li:nth-child(2)').click();
  await page.locator('#nav_list > li:nth-child(2)').waitFor();
  await page.locator('#submenu-item_29 > li:nth-child(8)').click();
  await page.locator('#nav_list > li:nth-child(2)').waitFor();
  await page.locator('#submenu-item_88 > li:nth-child(1) > a').click();
  await page.waitForLoadState('domcontentloaded');

  try {
    const upload = await page.waitForSelector('input[type=file]', { timeout: 5000 });
    console.log("🔍 Campo encontrado na página principal");
    await upload.setInputFiles(arquivoExcel);
  } catch {
    await frame.locator('input[type=file]').setInputFiles(arquivoExcel);

    const [dia, mes, ano] = dataHoje.split('/');
    const dataInput = `${ano}-${mes}-${dia}`;

    await frame.locator('body > div > form > div.row > div:nth-child(2) > input[type=date]').fill(dataInput);
    await frame.locator('body > div > form > div.row > div:nth-child(3) > input[type=date]').fill(dataInput);
    await frame.locator('text=Carregar, visualizar e linkar').click();
    await page.waitForLoadState('networkidle');

    console.log("🚀 Upload concluído");
  }

  const botaoConfirmar = frame.locator('button:has-text("Confirmar Assinaturas")');

    if (await botaoConfirmar.count() > 0) {
      await botaoConfirmar.waitFor({ state: 'visible', timeout: 15000 });
      await botaoConfirmar.click({ force: true });
    
      console.log("🏁 Confirmação realizada");
    } else {
      console.log("⚠️ Nenhum botão de confirmação encontrado (possivelmente sem dados)");
    }
  
  await page.waitForLoadState('networkidle');

  console.log("🏁 Confirmação realizada");
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
  console.log("⏳ Aguardando", atraso / 1000, "segundos...");
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

                  const ontem = new Date();
                  ontem.setDate(hoje.getDate() - 1);
                  
                  const diaOntem = ontem.getDate().toString().padStart(2, '0');
                  const mesOntem = (ontem.getMonth() + 1).toString().padStart(2, '0');
                  const anoOntem = ontem.getFullYear();
                  
                  const dataOntem = `${diaOntem}/${mesOntem}/${anoOntem}`;

  
  const dataArquivo = `${dia}-${mes}-${ano}`;

  const urlNormal =
    `https://sirius.assim.com.br/assimcsp/autorizador/preresultado.csp?idHospital=52345&DataIni=${dataOntem}&DataFim=${dataHoje}&executor=T&natservico=T&servico=T&especialidade=T&amb=&prefeitura=0&tuss=`;

  const urlPrefeitura =
    `https://sirius.assim.com.br/assimcsp/autorizador/preresultado.csp?idHospital=52345&DataIni=${dataHoje}&DataFim=${dataHoje}&executor=T&natservico=T&servico=T&especialidade=T&amb=&prefeitura=1&tuss=`;

  const registrosNormal = await extrairRelatorio(page, urlNormal);
  const registrosPrefeitura = await extrairRelatorio(page, urlPrefeitura);

  console.log("Registros de Atendimento Normal:", Math.max(registrosNormal.length - 1, 0));
  console.log("Registros de Atendimento da Prefeitura:", Math.max(registrosPrefeitura.length - 1, 0));

  const registrosTodos = [...registrosNormal, ...registrosPrefeitura];

  const dadosBanco = transformarParaBanco(registrosTodos);

  log("INFO", `📦 Enviando ${dadosBanco.length} registros em lotes`);
  
  await enviarEmLotes(dadosBanco);

  if (dadosBanco.length === 0) {
    log("INFO", "📭 Nenhum dado encontrado. Pulando envio para Órbita.");
  
    await browser.close();
    process.exit(0);
  }
  
  const worksheet = XLSX.utils.json_to_sheet(registrosTodos);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Relatorio");

  const nomeArquivo = `relatorio_assim_${dataArquivo}.xlsx`;
  const pastaRelatorios = path.join(__dirname, 'relatorios');

  if (!fs.existsSync(pastaRelatorios)) {
    fs.mkdirSync(pastaRelatorios, { recursive: true });
  }

  const caminhoArquivo = path.join(pastaRelatorios, nomeArquivo);
  XLSX.writeFile(workbook, caminhoArquivo);

  await enviarRelatorioDrive(caminhoArquivo, nomeArquivo);
  console.log("📊 Excel gerado:", nomeArquivo);

  const userOrbita = process.env.ORBITA_USER;
  const passOrbita = process.env.ORBITA_PASS;

  if (!userOrbita || !passOrbita) {
    log("ERROR", "🔐 Credenciais do Órbita não encontradas");
    await browser.close();
    process.exit(1);
  }

  await loginOrbita(page, userOrbita, passOrbita);
  await enviarExcelOrbita(page, caminhoArquivo, dataHoje);

  console.log("Tempo TOTAL:", tempo(inicioTotal));

  const pastaLogs = path.join(__dirname, 'logs');
  const arquivos = fs.readdirSync(pastaLogs);

  const ultimoLog = arquivos
    .filter(f => f.startsWith('log-'))
    .sort()
    .pop();

  if (ultimoLog) {
    const caminhoLog = path.join(pastaLogs, ultimoLog);
    log("INFO", "📁 Log selecionado: " + ultimoLog);
    await enviarLogDrive(caminhoLog, ultimoLog);
  } else {
    log("ERROR", "📂 Nenhum arquivo de log encontrado para envio");
  }

  log("SUCCESS", `🏁 Execução finalizada com sucesso em ${tempo(inicioTotal)}`);

  await browser.close();
  console.log("✅ Execução finalizada com sucesso");
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

    const text = await response.text();
    log("INFO", "☁️ Resposta do Drive: " + text);
  } catch (erro) {
    log("ERROR", "📊 Erro ao enviar relatório");
    log("ERROR", `❌ ${erro.message}`);
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

    const text = await response.text();
    log("INFO", "☁️ Resposta do Drive (LOG): " + text);
  } catch (erro) {
    log("ERROR", "📁 ☁️ Erro ao enviar LOG");
    log("ERROR", "❌ " + erro.message);
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

    const text = await response.text();
    log("INFO", "📡 Resposta do STATE: " + text);

    } catch (erro) {
    log("ERROR", "📡 Erro ao salvar status remoto");
    log("ERROR", "❌ " + erro.message);
    }
}
