// =======================================================
// ROBO ASSIM → ORBITA
// Versão: 3.0
// =======================================================

// =========================
// IMPORTS
// =========================

const { chromium } = require('playwright');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

function log(tipo, mensagem) {
  const agora = new Date();

  const data = agora.toISOString().slice(0,10);
  const hora = agora.toTimeString().slice(0,8);

  const linha = `[${data} ${hora}] [${tipo}] ${mensagem}\n`;

  const pastaLogs = path.join(__dirname, 'logs');

  if (!fs.existsSync(pastaLogs)) {
    fs.mkdirSync(pastaLogs, { recursive: true });
  }

  const caminhoLog = path.join(pastaLogs, `log-${data}.txt`);

  fs.appendFileSync(caminhoLog, linha);

  console.log(linha.trim());
}

function tempo(inicio){
  return ((Date.now() - inicio) / 1000).toFixed(2) + "s";
}

function lerStatus() {

  const caminho = path.join(__dirname, 'state', 'status-site.json');

  if (!fs.existsSync(caminho)) return null;

  try {
    return JSON.parse(fs.readFileSync(caminho, 'utf-8')).status;
  } catch (erro) {
    log("ERROR", "Erro ao ler status-site.json");
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

    log("INFO", "Mensagem enviada ao Slack");

  } catch (erro) {

    if (erro.name === 'AbortError') {
      log("ERROR", "Timeout ao enviar mensagem para Slack");
    } else {
      log("ERROR", "Erro ao enviar mensagem para Slack");
      log("ERROR", erro.message);
    }

  }
}


async function acessarComRetry(page, url, tentativas = 3) {

  for (let i = 1; i <= tentativas; i++) {

    try {

      log("INFO", `🌐 Tentativa ${i} de acesso ao site...`);

      await page.goto(url, {
        waitUntil: 'domcontentloaded',
        timeout: 15000
      });

      log("SUCCESS", "✅ Acesso realizado com sucesso");
      return true;

    } catch (erro) {

      log("ERROR", `❌ Erro na tentativa ${i}: ${erro.message}`);

      if (i === tentativas) {
        log("ERROR", "🚨 Falha total ao acessar o site");
        return false;
      }

      log("INFO", "⏳ Aguardando 5s para nova tentativa...");
      await page.waitForTimeout(5000);
    }
  }
}

// =========================
// EXTRAIR RELATÓRIO
// =========================
async function extrairRelatorio(page, urlConsulta){

  try {

    const tRelatorio = Date.now();

    await page.goto(urlConsulta,{
      waitUntil:'domcontentloaded',
      timeout:60000
    });

  const currentUrl = page.url();

  if(currentUrl.includes('preresultado')){
    const url = currentUrl.replace('preresultado','resultadosempaginacao');

    await page.goto(url,{
      waitUntil:'domcontentloaded',
      timeout:60000
    });
  }

  await page.waitForSelector('pre, table',{ timeout:120000 });

  let extractionContext = page;

  for(const f of page.frames()){
    const hasPre = await f.locator('pre').count().catch(()=>0);

    if(hasPre > 0){
      extractionContext = f;
      break;
    }
  }

  const registros = await extractionContext.evaluate(()=>{

    const linhas = Array.from(document.querySelectorAll('tr'));
    const dados = [];
    let ultimoRegistro = {};

    linhas.forEach(linha=>{

      const pre = linha.querySelector('pre');
      if(!pre) return;

      const td = linha.querySelectorAll('td');

      const dataHora = td[0]?.innerText?.trim();
      const sv = td[1]?.innerText?.trim();
      const nat = td[2]?.innerText?.trim();

      const beneficiarioCell = td[3]?.innerHTML;

      let matricula;
      let beneficiario;

      if(beneficiarioCell){
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
    !( (r.matricula || '').trim() === '' && (r.beneficiario || '').trim() === '' )
  );
  
function normalizarDataHora(dataHora) {

  if (!dataHora) return undefined;

  // Já tem ano
  if (dataHora.match(/\d{2}\/\d{2}\/\d{4}/)) {

    if (dataHora.length === 16) {
      return dataHora + ":00";
    }

    return dataHora;
  }

  // Não tem ano (ex: 17/03 08:03)
  const hoje = new Date();
  const ano = hoje.getFullYear();

  let novaData = `${dataHora.slice(0,5)}/${ano} ${dataHora.slice(6)}`;

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
    nomeCompleto = `${matricula}\n ${beneficiario}`.trim();
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

}
   catch (erro) {

    log("ERROR", `ERRO ao acessar relatório: ${urlConsulta}`);
	log("ERROR", `Detalhe: ${erro.message}`);

    return []; // 👈 evita quebrar o fluxo

  }
}

// =========================
// LOGIN ORBITA
// =========================

async function loginOrbita(page, usuario, senha){

  await page.goto("https://cronogramauniversoaba.com.br/app_Login/");

  await page.fill('input[placeholder="Usuário"]', usuario);
  await page.fill('input[placeholder="Senha"]', senha);

  await Promise.all([
    page.waitForNavigation({ waitUntil:'domcontentloaded' }),
    page.click('text=Entrar')
  ]);

  await page.waitForLoadState('networkidle');

}

// =========================
// UPLOAD EXCEL ORBITA
// =========================

async function enviarExcelOrbita(page, arquivoExcel, dataHoje){

	const frame = page.frameLocator('#iframe_item_89');
	
  await page.mouse.move(5,300);
  await page.locator('#nav_list > li:nth-child(2)').waitFor();

  await page.locator('#nav_list > li:nth-child(2)').click();

  await page.locator('#nav_list > li:nth-child(2)').waitFor();

  await page.locator('#submenu-item_29 > li:nth-child(8)').click();

  await page.locator('#nav_list > li:nth-child(2)').waitFor();

  await page.locator('#submenu-item_88 > li:nth-child(1) > a').click();

  await page.waitForLoadState('domcontentloaded');

  try {

    const upload = await page.waitForSelector('input[type=file]', { timeout: 5000 });

    console.log("Campo encontrado na página principal");

    await upload.setInputFiles(arquivoExcel);

  } catch {

    // upload do arquivo
    await frame.locator('input[type=file]').setInputFiles(arquivoExcel);

	const [dia, mes, ano] = dataHoje.split('/');
	const dataInput = `${ano}-${mes}-${dia}`;

	await frame.locator('body > div > form > div.row > div:nth-child(2) > input[type=date]').fill(dataInput);

	await frame.locator('body > div > form > div.row > div:nth-child(3) > input[type=date]').fill(dataInput);

    await frame.locator('text=Carregar, visualizar e linkar').click();

    await page.waitForLoadState('networkidle');

    console.log("Upload concluído");
	
  }

	// espera botão aparecer
const botaoConfirmar = frame.locator('button:has-text("Confirmar Assinaturas")');

await botaoConfirmar.waitFor({ state: 'visible', timeout: 15000 });

	// força o clique (importante)
await botaoConfirmar.click({ force: true });

await page.waitForLoadState('networkidle');

console.log("Confirmação realizada");
  
await page.waitForTimeout(2000);

}

// =========================
// EXECUÇÃO PRINCIPAL
// =========================

(async () => {

const inicioTotal = Date.now();

const browser = await chromium.launch({
  headless:true
});

const context = await browser.newContext();

// BLOQUEAR IMAGENS / FONTES / CSS
const bloqueados = new Set(['image', 'font', 'stylesheet', 'media']);

await context.route('**/*', route => {
  const tipo = route.request().resourceType();
  return bloqueados.has(tipo) ? route.abort() : route.continue();
});

// LOGIN ASSIM
const page = await context.newPage();

const sucesso = await acessarComRetry(
  page,
  'https://sirius.assim.com.br/assimcsp/autorizador/login.csp'
);

const statusAnterior = await obterStatusRemoto();

// 🔴 SITE FORA
if (!sucesso) {

  if (statusAnterior !== "offline") {

    log("ERROR", "Site ficou OFFLINE");

    await enviarSlack(`🔴 *INDISPONIBILIDADE DETECTADA*\nO site do Autorizador da Assim está FORA do ar.\n⏰ ${new Date().toLocaleString('pt-BR')}`);

  }

  await salvarStatusRemoto("offline");

  process.exit(1);
}

// 🟢 SITE VOLTOU
if (sucesso) {

  if (statusAnterior === "offline") {

    log("SUCCESS", "Site voltou ao normal");

    await enviarSlack(`🟢 *SERVIÇO RESTABELECIDO*\nO site do Autorizador da Assim voltou a operar normalmente.\n⏰ ${new Date().toLocaleString('pt-BR')}`);

  }

  await salvarStatusRemoto("online");
}

await page.selectOption('select','52345');
await page.fill('input[type="password"]', process.env.SENHA);

await Promise.all([
  page.waitForNavigation(),
  page.click('text=Entrar')
]);

// DATA

await page.waitForSelector('select[name="DiaFim"]');

const hoje = new Date();

const dia = hoje.getDate().toString().padStart(2,'0');
const mes = (hoje.getMonth()+1).toString().padStart(2,'0');
const ano = hoje.getFullYear();

const dataHoje = `${dia}/${mes}/${ano}`;
const dataInput = `${ano}-${mes}-${dia}`;
const dataArquivo = `${dia}-${mes}-${ano}`;

// URLS

const urlNormal =
`https://sirius.assim.com.br/assimcsp/autorizador/preresultado.csp?idHospital=52345&DataIni=${dataHoje}&DataFim=${dataHoje}&executor=T&natservico=T&servico=T&especialidade=T&amb=&prefeitura=0&tuss=`;

const urlPrefeitura =
`https://sirius.assim.com.br/assimcsp/autorizador/preresultado.csp?idHospital=52345&DataIni=${dataHoje}&DataFim=${dataHoje}&executor=T&natservico=T&servico=T&especialidade=T&amb=&prefeitura=1&tuss=`;

// EXTRAÇÃO

const registrosNormal = await extrairRelatorio(page, urlNormal);
const registrosPrefeitura = await extrairRelatorio(page, urlPrefeitura);

console.log(
  "Registros normal:",
  Math.max(registrosNormal.length - 1, 0)
);

console.log(
  "Registros prefeitura:",
  Math.max(registrosPrefeitura.length - 1, 0)
);
const registrosTodos = [  ...registrosNormal,  ...registrosPrefeitura];

// EXCEL

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
	
console.log("Excel gerado:", nomeArquivo);

// LOGIN ORBITA

await loginOrbita(page,"caiovinicius","C@io1309");

// UPLOAD

await enviarExcelOrbita(page,caminhoArquivo,dataHoje);


console.log("Tempo TOTAL:", tempo(inicioTotal));
	
const pastaLogs = path.join(__dirname, 'logs');

const arquivos = fs.readdirSync(pastaLogs);

const ultimoLog = arquivos
  .filter(f => f.startsWith('log-'))
  .sort()
  .pop();

const caminhoLog = path.join(pastaLogs, ultimoLog);

log("INFO", "Log selecionado: " + ultimoLog);

await enviarLogDrive(caminhoLog, `log-${dataHojeISO}.txt`);
	
await browser.close();

console.log("Execução finalizada com sucesso");
})();

// ===============================
// ENVIO DE RELATORIO PARA O DRIVE
// ===============================
	
async function enviarRelatorioDrive(caminhoArquivo, nomeArquivo) {
  try {
    const url = process.env.GOOGLE_SCRIPT_URL;

    if (!url) {
      log("ERROR", "GOOGLE_SCRIPT_URL não definida");
      return;
    }

    const fileBuffer = fs.readFileSync(caminhoArquivo);
    const base64 = fileBuffer.toString('base64');

    log("INFO", "Enviando relatório para o Drive...");

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

    log("INFO", "Resposta do Drive: " + text);

  } catch (erro) {
    log("ERROR", "Erro ao enviar relatório");
    log("ERROR", erro.message);
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

    log("INFO", "Enviando LOG para o Drive...");

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

    log("INFO", "Resposta do Drive (LOG): " + text);

  } catch (erro) {
    log("ERROR", "Erro ao enviar LOG");
    log("ERROR", erro.message);
  }
}

// =============================
// LER O STATUS
// =============================

async function obterStatusRemoto() {
	try {
    const url = process.env.GOOGLE_SCRIPT_URL;

    const res = await fetch(url);
    const data = await res.json();

    return data.status;

 	 } catch (erro) {
    log("ERROR", "Erro ao obter status remoto");
    return null;
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

    log("INFO", "Resposta do STATE: " + text);

  } catch (erro) {
    log("ERROR", "Erro ao salvar status remoto");
    log("ERROR", erro.message);
  }
}
