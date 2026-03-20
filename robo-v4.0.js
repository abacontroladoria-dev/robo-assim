const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

class RetryableError extends Error {
  constructor(message) {
    super(message);
    this.name = 'RetryableError';
  }
}

class FatalError extends Error {
  constructor(message) {
    super(message);
    this.name = 'FatalError';
  }
}

function logStructured(level, message, metadata = {}) {
  const logEntry = {
    timestamp: new Date().toISOString(),
    level,
    message,
    executionId: process.env.GITHUB_RUN_ID || 'local',
    ...metadata
  };

  console.log(JSON.stringify(logEntry));

  const logDir = 'logs';
  if (!fs.existsSync(logDir)) {
    fs.mkdirSync(logDir);
  }
  const logFile = path.join(logDir, `${new Date().toISOString().split('T')[0]}.jsonl`);
  fs.appendFileSync(logFile, JSON.stringify(logEntry) + '\n');
}

async function withRetry(fn, options = {}) {
  const {
    maxAttempts = 3,
    initialDelay = 1000,
    maxDelay = 30000,
    backoffMultiplier = 2,
    jitter = true,
    retryOn = (err) => true
  } = options;

  let lastError;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await fn();
    } catch (err) {
      lastError = err;
      if (attempt === maxAttempts || !retryOn(err)) {
        throw err;
      }

      let delay = initialDelay * Math.pow(backoffMultiplier, attempt - 1);
      if (jitter) delay *= (0.5 + Math.random());
      delay = Math.min(delay, maxDelay);

      logStructured('WARN', `Tentativa ${attempt}/${maxAttempts} falhou. Aguardando ${delay.toFixed(0)}ms...`, { error: err.message });
      await new Promise(r => setTimeout(r, delay));
    }
  }
  throw lastError;
}

class CircuitBreaker {
  constructor(threshold = 5, timeout = 300000) {
    this.failureCount = 0;
    this.threshold = threshold;
    this.timeout = timeout;
    this.state = 'CLOSED';
    this.lastFailureTime = null;
  }

  async execute(fn) {
    if (this.state === 'OPEN') {
      if (Date.now() - this.lastFailureTime > this.timeout) {
        this.state = 'HALF_OPEN';
        logStructured('INFO', 'Circuit breaker: tentando recuperar (HALF_OPEN)...');
      } else {
        const remainingTime = Math.ceil((this.timeout - (Date.now() - this.lastFailureTime)) / 1000);
        throw new Error(`Circuit breaker está ABERTO. Tente novamente em ${remainingTime}s.`);
      }
    }

    try {
      const result = await fn();
      this.onSuccess();
      return result;
    } catch (err) {
      this.onFailure();
      throw err;
    }
  }

  onSuccess() {
    if (this.state !== 'CLOSED') {
      logStructured('SUCCESS', 'Circuit breaker: recuperado e FECHADO.');
    }
    this.failureCount = 0;
    this.state = 'CLOSED';
  }

  onFailure() {
    this.failureCount++;
    this.lastFailureTime = Date.now();
    if (this.failureCount >= this.threshold) {
      this.state = 'OPEN';
      logStructured('ERROR', `Circuit breaker: ABERTO após ${this.failureCount} falhas consecutivas.`, { threshold: this.threshold });
    }
  }
}

class ExecutionMetrics {
  constructor() {
    this.startTime = Date.now();
    this.events = [];
  }

  recordEvent(name, duration, status = 'success', metadata = {}) {
    this.events.push({
      name,
      duration,
      status,
      timestamp: new Date().toISOString(),
      ...metadata
    });
  }

  summary() {
    const totalDuration = Date.now() - this.startTime;
    const successfulEvents = this.events.filter(e => e.status === 'success').length;
    const failedEvents = this.events.filter(e => e.status === 'failure').length;

    return {
      totalDuration: totalDuration,
      eventCount: this.events.length,
      successfulEvents,
      failedEvents,
      successRate: this.events.length > 0 ? successfulEvents / this.events.length : 0,
      events: this.events,
      slowestEvent: this.events.length > 0 ? this.events.reduce((a, b) => a.duration > b.duration ? a : b) : null
    };
  }
}

function validateEnv() {
  const required = ['SENHA', 'ORBITA_USER', 'ORBITA_PASS', 'GOOGLE_SCRIPT_URL', 'SLACK_WEBHOOK'];
  const missing = required.filter(v => !process.env[v]);

  if (missing.length > 0) {
    throw new FatalError(`Variáveis de ambiente ausentes: ${missing.join(', ')}. Por favor, configure-as.`);
  }

  const urlVars = ['GOOGLE_SCRIPT_URL', 'SLACK_WEBHOOK'];
  for (const key of urlVars) {
    try {
      new URL(process.env[key]);
    } catch (e) {
      throw new FatalError(`Variável de ambiente ${key} não é uma URL válida: ${process.env[key]}.`);
    }
  }

  logStructured('INFO', 'Todas as variáveis de ambiente validadas com sucesso.');
}

class AssimService {
  constructor(page, metrics) {
    this.page = page;
    this.metrics = metrics;
    this.baseUrl = 'https://sirius.assim.com.br/assimcsp/autorizador/';
  }

  async login(usuario, senha) {
    const start = Date.now();
    try {
      logStructured('INFO', 'Iniciando login no Assim...');
      await withRetry(async () => {
        await this.page.goto(this.baseUrl + 'login.csp', { waitUntil: 'domcontentloaded', timeout: 60000 });
        await this.page.fill('input[type="email"]', usuario);
        await this.page.fill('input[type="password"]', senha);
        await this.page.click('button[type="submit"]');
        await this.page.waitForLoadState('networkidle', { timeout: 60000 });

        const currentUrl = this.page.url();
        if (currentUrl.includes('login.csp')) {
          throw new RetryableError('Login Assim falhou. Credenciais inválidas ou página de login persistiu.');
        }
      }, {
        maxAttempts: 3,
        retryOn: (err) => err instanceof RetryableError || err.message.includes('timeout')
      });
      logStructured('SUCCESS', 'Login no Assim realizado com sucesso.');
      this.metrics.recordEvent('loginAssim', Date.now() - start, 'success');
    } catch (error) {
      this.metrics.recordEvent('loginAssim', Date.now() - start, 'failure', { error: error.message });
      throw new FatalError(`Falha fatal no login do Assim: ${error.message}`);
    }
  }

  async extrairRelatorio(urlConsulta) {
    const start = Date.now();
    let dados = [];
    try {
      logStructured('INFO', `Iniciando extração de relatório da Assim: ${urlConsulta}`);
      await withRetry(async () => {
        await this.page.goto(urlConsulta, { waitUntil: 'domcontentloaded', timeout: 90000 });

        const preResultadoLocator = this.page.locator('text=Pré-Resultado');
        if (await preResultadoLocator.isVisible()) {
          logStructured('INFO', 'Página de pré-resultado detectada. Navegando para paginação...');
          await this.page.goto(this.baseUrl + 'preresultado.csp?pagina=1', { waitUntil: 'domcontentloaded', timeout: 90000 });
        }

        await this.page.waitForSelector('table.table-striped', { timeout: 60000 });

        const rows = await this.page.$$('table.table-striped tbody tr');
        if (rows.length === 0) {
          logStructured('WARN', 'Nenhum registro encontrado na tabela.', { url: urlConsulta });
          return;
        }

        for (const row of rows) {
          const cols = await row.$$('td');
          if (cols.length >= 12) {
            const dataHora = await cols[0].textContent();
            const sv = await cols[1].textContent();
            const nat = await cols[2].textContent();
            const matricula = await cols[3].textContent();
            const beneficiario = await cols[4].textContent();
            const token = await cols[5].textContent();
            const justificativa = await cols[6].textContent();
            const processo = await cols[7].textContent();
            const guia = await cols[8].textContent();
            const soli = await cols[9].textContent();
            const especialidade = await cols[10].textContent();
            const codigo = await cols[11].textContent();
            const status = await cols[12].textContent();

            dados.push({
              dataHora: dataHora ? dataHora.trim() : '',
              sv: sv ? sv.trim() : '',
              nat: nat ? nat.trim() : '',
              matricula: matricula ? matricula.trim() : '',
              beneficiario: beneficiario ? beneficiario.trim() : '',
              token: token ? token.trim() : '',
              justificativa: justificativa ? justificativa.trim() : '',
              processo: processo ? processo.trim() : '',
              guia: guia ? guia.trim() : '',
              soli: soli ? soli.trim() : '',
              especialidade: especialidade ? especialidade.trim() : '',
              codigo: codigo ? codigo.trim() : '',
              status: status ? status.trim() : ''
            });
          }
        }
        logStructured('SUCCESS', `Extração de relatório concluída. ${dados.length} registros encontrados.`, { url: urlConsulta });
      }, {
        maxAttempts: 3,
        retryOn: (err) => err.message.includes('timeout') || err.message.includes('network')
      });
      this.metrics.recordEvent('extrairRelatorio', Date.now() - start, 'success', { url: urlConsulta, records: dados.length });
      return dados;
    } catch (error) {
      this.metrics.recordEvent('extrairRelatorio', Date.now() - start, 'failure', { url: urlConsulta, error: error.message });
      if (error.message.includes('404')) {
        throw new FatalError(`Página de relatório não encontrada: ${urlConsulta}. Erro: ${error.message}`);
      }
      throw new RetryableError(`Falha na extração do relatório da Assim: ${error.message}`);
    }
  }
}

class OrbitaService {
  constructor(page, metrics) {
    this.page = page;
    this.metrics = metrics;
    this.baseUrl = 'https://cronogramauniversoaba.com.br/';
  }

  async login(usuario, senha) {
    const start = Date.now();
    try {
      logStructured('INFO', 'Iniciando login no Orbita...');
      await withRetry(async () => {
        await this.page.goto(this.baseUrl + 'app_Login/', { waitUntil: 'domcontentloaded', timeout: 60000 });
        await this.page.fill('input[name="usuario"]', usuario);
        await this.page.fill('input[name="senha"]', senha);
        await this.page.click('button');
        await this.page.waitForLoadState('networkidle', { timeout: 60000 });

        const currentUrl = this.page.url();
        if (currentUrl.includes('app_Login')) {
          throw new RetryableError('Login Orbita falhou. Credenciais inválidas ou página de login persistiu.');
        }
      }, {
        maxAttempts: 3,
        retryOn: (err) => err instanceof RetryableError || err.message.includes('timeout')
      });
      logStructured('SUCCESS', 'Login no Orbita realizado com sucesso.');
      this.metrics.recordEvent('loginOrbita', Date.now() - start, 'success');
    } catch (error) {
      this.metrics.recordEvent('loginOrbita', Date.now() - start, 'failure', { error: error.message });
      throw new FatalError(`Falha fatal no login do Orbita: ${error.message}`);
    }
  }

  async enviarExcel(filePath, dataRelatorio) {
    const start = Date.now();
    try {
      logStructured('INFO', `Iniciando envio do arquivo Excel para Orbita: ${filePath}`);
      await withRetry(async () => {
        await this.page.goto(this.baseUrl + 'app_Upload_Relatorio_Assinatura/', { waitUntil: 'domcontentloaded', timeout: 60000 });

        await this.page.locator('text=Carregar, visualizar e linkar').click();
        await this.page.waitForLoadState('networkidle', { timeout: 30000 });

        const frameLocator = this.page.frameLocator('iframe[name="app_Upload_Relatorio_Assinatura_iframe"]');
        await frameLocator.locator('input[type="file"]').waitFor({ state: 'visible', timeout: 30000 });

        await frameLocator.locator('input[type="file"]').setInputFiles(filePath);
        logStructured('INFO', 'Arquivo anexado no Orbita.');

        await frameLocator.locator('button:has-text("Upload")').click();
        await this.page.waitForLoadState('networkidle', { timeout: 90000 });

        const successMessageLocator = this.page.locator('text=Arquivo enviado com sucesso!');
        if (!(await successMessageLocator.isVisible())) {
          throw new RetryableError('Upload para Orbita falhou ou mensagem de sucesso não apareceu.');
        }
        logStructured('SUCCESS', 'Arquivo enviado com sucesso para Orbita.');

        const botaoConfirmar = this.page.locator('button:has-text("Confirmar Assinaturas")');
        await botaoConfirmar.waitFor({ state: 'visible', timeout: 15000 });
        await botaoConfirmar.click();
        await this.page.waitForLoadState('networkidle', { timeout: 60000 });

        const confirmSuccessLocator = this.page.locator(`text=Assinaturas do dia ${dataRelatorio} confirmadas com sucesso!`);
        if (!(await confirmSuccessLocator.isVisible())) {
          throw new RetryableError('Confirmação de assinaturas no Orbita falhou ou mensagem de sucesso não apareceu.');
        }
        logStructured('SUCCESS', `Assinaturas do dia ${dataRelatorio} confirmadas com sucesso no Orbita.`);

      }, {
        maxAttempts: 3,
        retryOn: (err) => err instanceof RetryableError || err.message.includes('timeout')
      });
      this.metrics.recordEvent('enviarExcelOrbita', Date.now() - start, 'success', { filePath });
    } catch (error) {
      this.metrics.recordEvent('enviarExcelOrbita', Date.now() - start, 'failure', { filePath, error: error.message });
      throw new FatalError(`Falha fatal no envio para Orbita: ${error.message}`);
    }
  }
}

class DriveService {
  constructor(googleScriptUrl, metrics) {
    this.googleScriptUrl = googleScriptUrl;
    this.metrics = metrics;
  }

  async uploadFile(filePath, fileName, folderName) {
    const start = Date.now();
    try {
      logStructured('INFO', `Iniciando upload do arquivo para Google Drive: ${fileName} na pasta ${folderName}`);
      const fileContent = fs.readFileSync(filePath).toString('base64');

      const payload = {
        action: 'uploadFile',
        fileName: fileName,
        folderName: folderName,
        fileContent: fileContent
      };

      await withRetry(async () => {
        const response = await fetch(this.googleScriptUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
          timeout: 120000
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new RetryableError(`Erro no upload para Drive: ${response.status} - ${errorText}`);
        }
        logStructured('SUCCESS', `Arquivo ${fileName} enviado com sucesso para o Google Drive.`);
      }, {
        maxAttempts: 3,
        retryOn: (err) => err instanceof RetryableError || err.message.includes('timeout')
      });
      this.metrics.recordEvent('uploadDriveFile', Date.now() - start, 'success', { fileName, folderName });
    } catch (error) {
      this.metrics.recordEvent('uploadDriveFile', Date.now() - start, 'failure', { fileName, folderName, error: error.message });
      throw new FatalError(`Falha fatal no upload para Google Drive: ${error.message}`);
    }
  }

  async uploadLog(logFilePath) {
    const start = Date.now();
    try {
      logStructured('INFO', `Iniciando upload do arquivo de log para Google Drive: ${logFilePath}`);
      const fileName = path.basename(logFilePath);
      const folderName = 'Logs Robô Assim';
      const fileContent = fs.readFileSync(logFilePath).toString('base64');

      const payload = {
        action: 'uploadFile',
        fileName: fileName,
        folderName: folderName,
        fileContent: fileContent
      };

      await withRetry(async () => {
        const response = await fetch(this.googleScriptUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
          timeout: 60000
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new RetryableError(`Erro no upload do log para Drive: ${response.status} - ${errorText}`);
        }
        logStructured('SUCCESS', `Arquivo de log ${fileName} enviado com sucesso para o Google Drive.`);
      }, {
        maxAttempts: 3,
        retryOn: (err) => err instanceof RetryableError || err.message.includes('timeout')
      });
      this.metrics.recordEvent('uploadDriveLog', Date.now() - start, 'success', { fileName });
    } catch (error) {
      this.metrics.recordEvent('uploadDriveLog', Date.now() - start, 'failure', { fileName, error: error.message });
      logStructured('ERROR', `Falha fatal no upload do log para Google Drive: ${error.message}`);
    }
  }

  async saveRemoteStatus(status) {
    const start = Date.now();
    try {
      logStructured('INFO', `Salvando status remoto do Assim: ${status}`);
      const payload = {
        action: 'saveStatus',
        status: status
      };

      await withRetry(async () => {
        const response = await fetch(this.googleScriptUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
          timeout: 30000
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new RetryableError(`Erro ao salvar status remoto: ${response.status} - ${errorText}`);
        }
        logStructured('SUCCESS', `Status remoto '${status}' salvo com sucesso.`);
      }, {
        maxAttempts: 3,
        retryOn: (err) => err instanceof RetryableError || err.message.includes('timeout')
      });
      this.metrics.recordEvent('saveRemoteStatus', Date.now() - start, 'success', { status });
    } catch (error) {
      this.metrics.recordEvent('saveRemoteStatus', Date.now() - start, 'failure', { status, error: error.message });
      logStructured('ERROR', `Falha ao salvar status remoto: ${error.message}`);
    }
  }

  async getRemoteStatus() {
    const start = Date.now();
    try {
      logStructured('INFO', 'Obtendo status remoto do Assim...');
      let status = 'unknown';
      await withRetry(async () => {
        const payload = { action: 'getStatus' };
        const response = await fetch(this.googleScriptUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
          timeout: 30000
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new RetryableError(`Erro ao obter status remoto: ${response.status} - ${errorText}`);
        }
        const data = await response.json();
        status = data.status || 'unknown';
        logStructured('SUCCESS', `Status remoto obtido: ${status}.`);
      }, {
        maxAttempts: 3,
        retryOn: (err) => err instanceof RetryableError || err.message.includes('timeout')
      });
      this.metrics.recordEvent('getRemoteStatus', Date.now() - start, 'success', { status });
      return status;
    } catch (error) {
      this.metrics.recordEvent('getRemoteStatus', Date.now() - start, 'failure', { error: error.message });
      logStructured('ERROR', `Falha ao obter status remoto: ${error.message}`);
      return 'unknown';
    }
  }
}

const slackCircuitBreaker = new CircuitBreaker(3, 300000);

async function sendSlackNotification(message) {
  try {
    await slackCircuitBreaker.execute(async () => {
      logStructured('INFO', 'Enviando notificação para o Slack...');
      const response = await fetch(process.env.SLACK_WEBHOOK, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text: message }),
        timeout: 10000
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Erro ao enviar para Slack: ${response.status} - ${errorText}`);
      }
      logStructured('SUCCESS', 'Notificação enviada para o Slack com sucesso.');
    });
  } catch (error) {
    logStructured('ERROR', `Falha ao enviar notificação para o Slack (Circuit Breaker): ${error.message}`);
  }
}

function dentroDoHorario() {
  const now = new Date();
  const day = now.getDay();
  const hour = now.getHours();
  const minutes = now.getMinutes();

  if (day < 1 || day > 5) {
    logStructured('INFO', 'Fora do horário: Não é dia de semana.', { dayOfWeek: day });
    return false;
  }

  if (hour < 8 || (hour === 8 && minutes < 30) || hour >= 18) {
    logStructured('INFO', 'Fora do horário: Horário de execução não permitido.', { currentHour: hour, currentMinutes: minutes });
    return false;
  }

  logStructured('INFO', 'Dentro do horário de execução permitido.');
  return true;
}

async function verificarMudancaStatus(driveService) {
  const assimUrl = 'https://sirius.assim.com.br/assimcsp/autorizador/login.csp';
  let currentStatus = 'offline';

  try {
    logStructured('INFO', 'Verificando status atual do site Assim...');
    const response = await fetch(assimUrl, { timeout: 10000 });
    if (response.ok) {
      currentStatus = 'online';
    }
  } catch (error) {
    logStructured('WARN', `Erro ao verificar status do Assim: ${error.message}`);
    currentStatus = 'offline';
  }

  const previousStatus = await driveService.getRemoteStatus();
  logStructured('INFO', `Status anterior: ${previousStatus}, Status atual: ${currentStatus}`);

  if (currentStatus !== previousStatus) {
    if (currentStatus === 'online') {
      await sendSlackNotification('✅ O site da Assim está ONLINE novamente!');
      logStructured('SUCCESS', 'Site Assim voltou a ficar ONLINE.');
    } else {
      await sendSlackNotification('❌ O site da Assim está OFFLINE!');
      logStructured('ERROR', 'Site Assim ficou OFFLINE.');
    }
    await driveService.saveRemoteStatus(currentStatus);
  } else {
    logStructured('INFO', `Status do site Assim permanece ${currentStatus}.`);
  }
}

async function robo() {
  const metrics = new ExecutionMetrics();
  let browser;
  let page;
  let excelFilePath = '';
  const logFileName = `${new Date().toISOString().split('T')[0]}.jsonl`;
  const logFilePath = path.join('logs', logFileName);

  try {
    logStructured('INFO', 'Iniciando execução do robô Universo ABA.', { version: '4.0' });

    validateEnv();
    metrics.recordEvent('validateEnv', Date.now() - metrics.startTime, 'success');

    if (!dentroDoHorario()) {
      logStructured('INFO', 'Robô fora do horário de execução permitido. Encerrando.');
      process.exit(0);
    }

    const driveService = new DriveService(process.env.GOOGLE_SCRIPT_URL, metrics);

    await verificarMudancaStatus(driveService);
    metrics.recordEvent('verificarMudancaStatus', Date.now() - metrics.startTime, 'success');

    const currentAssimStatus = await driveService.getRemoteStatus();
    if (currentAssimStatus === 'offline') {
      logStructured('WARN', 'Site Assim está OFFLINE. Não será possível extrair relatórios. Encerrando.');
      await sendSlackNotification('⚠️ Robô encerrado: Site Assim está OFFLINE. Não foi possível extrair relatórios.');
      process.exit(0);
    }

    browser = await chromium.launch({ headless: true });
    page = await browser.newPage();
    await page.route('**/*', route => {
      const resourceType = route.request().resourceType();
      if (resourceType === 'image' || resourceType === 'stylesheet' || resourceType === 'font') {
        route.abort();
      } else {
        route.continue();
      }
    });
    logStructured('INFO', 'Navegador Playwright iniciado e configurado.');
    metrics.recordEvent('initPlaywright', Date.now() - metrics.startTime, 'success');

    const assimService = new AssimService(page, metrics);
    const orbitaService = new OrbitaService(page, metrics);

    await assimService.login(process.env.ORBITA_USER, process.env.SENHA);
    logStructured('SUCCESS', 'Login no Assim concluído.');

    const today = new Date();
    const formattedDate = `${today.getDate().toString().padStart(2, '0')}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getFullYear()}`;
    const urlNormal = `https://sirius.assim.com.br/assimcsp/autorizador/preresultado.csp?data=${formattedDate}&prefeitura=0`;
    const urlPrefeitura = `https://sirius.assim.com.br/assimcsp/autorizador/preresultado.csp?data=${formattedDate}&prefeitura=1`;

    const relatorioNormal = await assimService.extrairRelatorio(urlNormal);
    const relatorioPrefeitura = await assimService.extrairRelatorio(urlPrefeitura);

    const allReports = [...relatorioNormal, ...relatorioPrefeitura];
    logStructured('INFO', `Total de registros extraídos: ${allReports.length}`);

    if (allReports.length === 0) {
      logStructured('WARN', 'Nenhum relatório encontrado para o dia. Encerrando.');
      await sendSlackNotification('⚠️ Robô encerrado: Nenhum relatório encontrado para o dia.');
      process.exit(0);
    }

    const ws = XLSX.utils.json_to_sheet(allReports);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'RelatorioAssim');
    const excelFileName = `relatorio_assim_${formattedDate}.xlsx`;
    excelFilePath = path.join(__dirname, excelFileName);
    XLSX.writeFile(wb, excelFilePath);
    logStructured('SUCCESS', `Arquivo Excel gerado: ${excelFilePath}`);
    metrics.recordEvent('generateExcel', Date.now() - metrics.startTime, 'success', { records: allReports.length });

    await driveService.uploadFile(excelFilePath, excelFileName, 'Relatórios Assim');
    logStructured('SUCCESS', 'Relatório enviado para Google Drive.');

    await orbitaService.login(process.env.ORBITA_USER, process.env.ORBITA_PASS);
    logStructured('SUCCESS', 'Login no Orbita concluído.');

    await orbitaService.enviarExcel(excelFilePath, formattedDate);
    logStructured('SUCCESS', 'Relatório enviado para Orbita e assinaturas confirmadas.');

    await driveService.uploadLog(logFilePath);

    logStructured('SUCCESS', 'Robô executado com sucesso!');
    await sendSlackNotification('✅ Robô executado com sucesso! Relatórios extraídos, enviados para Drive e Orbita.');

  } catch (error) {
    logStructured('ERROR', `Erro durante a execução do robô: ${error.name} - ${error.message}`, { stack: error.stack });
    await sendSlackNotification(`❌ Erro na execução do robô: ${error.name} - ${error.message}`);

    if (logFilePath && fs.existsSync(logFilePath)) {
      const tempDriveService = new DriveService(process.env.GOOGLE_SCRIPT_URL, metrics);
      await tempDriveService.uploadLog(logFilePath);
    }

    metrics.recordEvent('roboExecution', Date.now() - metrics.startTime, 'failure', { error: error.message });
    process.exit(1);
  } finally {
    if (browser) {
      await browser.close();
      logStructured('INFO', 'Navegador Playwright fechado.');
    }
    if (excelFilePath && fs.existsSync(excelFilePath)) {
      fs.unlinkSync(excelFilePath);
      logStructured('INFO', `Arquivo Excel temporário removido: ${excelFilePath}`);
    }

    const finalSummary = metrics.summary();
    logStructured('INFO', 'Resumo final das métricas de execução:', finalSummary);

    if (finalSummary.totalDuration > 300000) {
      await sendSlackNotification(`⚠️ Alerta de Performance: Execução do robô demorou ${(finalSummary.totalDuration / 1000).toFixed(2)}s.`);
    }
    if (finalSummary.failedEvents > 0) {
      await sendSlackNotification(`⚠️ Alerta de Falhas: ${finalSummary.failedEvents} eventos falharam durante a execução.`);
    }
  }
}

robo();
