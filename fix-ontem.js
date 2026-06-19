// Script pontual: busca dados de ontem e sobe para o Órbita
require('dotenv').config();
const { chromium } = require('playwright');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

function log(tipo, msg) {
  const hora = new Date().toLocaleTimeString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  console.log(`[${hora}] [${tipo}] ${msg}`);
}

async function extrairRelatorio(page, urlConsulta) {
  try {
    await page.goto(urlConsulta, { waitUntil: 'domcontentloaded', timeout: 60000 });

    let currentUrl = page.url();
    log("INFO", `URL final: ${currentUrl}`);

    if (currentUrl.includes('preresultado')) {
      const urlSemPaginacao = currentUrl.replace('preresultado', 'resultadosempaginacao');
      await page.goto(urlSemPaginacao, { waitUntil: 'domcontentloaded', timeout: 60000 });
      currentUrl = page.url();
    }

    await page.waitForSelector('pre, table', { timeout: 60000 });

    const resultado = await page.evaluate(() => {
      const allTrs = Array.from(document.querySelectorAll('tr'));
      let columnHeaderRow = null;
      for (const tr of allTrs) {
        const ths = Array.from(tr.querySelectorAll('th'));
        const texts = ths.map(th => (th.innerText || '').trim().toLowerCase());
        if (texts.some(t => t === 'guia') && texts.some(t => t === 'sv' || t === 'nat')) {
          columnHeaderRow = tr;
          break;
        }
      }
      if (!columnHeaderRow) {
        for (const tr of allTrs) {
          if (tr.querySelectorAll('th').length > 8) { columnHeaderRow = tr; break; }
        }
      }

      const headers = columnHeaderRow
        ? Array.from(columnHeaderRow.querySelectorAll('th')).map(th => (th.innerText || '').trim())
        : [];

      const headerMap = {};
      headers.forEach((h, i) => { headerMap[h.toLowerCase().replace(/\s+/g, ' ')] = i; });

      function resolverCol(...cs) {
        for (const c of cs) {
          const idx = headerMap[c.toLowerCase().replace(/\s+/g, ' ')];
          if (idx !== undefined) return idx;
        }
        return -1;
      }

      const colDataHora  = resolverCol('data hora', 'data/hora', 'data');
      const colSv        = resolverCol('sv', 's.v.');
      const colNat       = resolverCol('nat', 'nat.');
      const colBenefic   = resolverCol('beneficiário', 'beneficiario', 'paciente');
      const colBiofacial = resolverCol('biofacial', 'bio facial');
      const colToken     = resolverCol('token', 'senha');
      const colJustif    = resolverCol('justificativa', 'justif.');
      const colProcesso  = resolverCol('processo', 'n° processo');
      const colGuia      = resolverCol('guia', 'n° guia');
      const colSoli      = resolverCol('soli', 'solicitante');
      const colEspec     = resolverCol('especialidade', 'espec.');
      const colCodigo    = colEspec >= 0 ? colEspec + 1 : 11;
      const colStatus    = colEspec >= 0 ? colEspec + 2 : 12;

      const dados = [];
      let ultimo = {};

      Array.from(document.querySelectorAll('tr')).forEach(linha => {
        const td = linha.querySelectorAll('td');
        if (td.length === 0) return;

        function get(idx, fb) {
          const i = idx >= 0 ? idx : fb;
          return (i >= 0 && i < td.length) ? td[i]?.innerText?.trim() || undefined : undefined;
        }

        const dataHora = get(colDataHora, 0);
        const sv       = get(colSv, 1);
        const nat      = get(colNat, 2);
        const benefRaw = get(colBenefic, 3) || "";
        const partes   = benefRaw.split('\n').map(p => p.trim()).filter(Boolean);
        const matricula    = partes[0] || null;
        const beneficiario = partes[1] || partes[0] || null;
        const biofacial    = get(colBiofacial, 4);
        const token        = get(colToken, 5);
        const justificativa = get(colJustif, 6);
        const processo     = get(colProcesso, 7);
        const guia         = get(colGuia, 8);
        if (!guia || !/^\d+$/.test(guia)) return;
        const soli         = get(colSoli, 9);
        const especialidade = get(colEspec, 10);
        const codigoRaw    = get(colCodigo, 11) || "";
        const partesCod    = codigoRaw.split(/\s+/);
        const codigo       = partesCod[0] || null;
        let status         = get(colStatus, 12);
        if (!status) status = partesCod.slice(1).join(' ') || null;

        const r = {
          dataHora:      dataHora      ?? ultimo.dataHora,
          sv:            sv            ?? ultimo.sv,
          nat:           nat           ?? ultimo.nat,
          matricula:     matricula     ?? ultimo.matricula,
          beneficiario:  beneficiario  ?? ultimo.beneficiario,
          biofacial:     biofacial     ?? ultimo.biofacial,
          token:         token         ?? ultimo.token,
          justificativa: justificativa ?? ultimo.justificativa,
          processo:      processo      ?? ultimo.processo,
          guia,
          soli:          soli          ?? ultimo.soli,
          especialidade: especialidade ?? ultimo.especialidade,
          codigo,
          status
        };
        dados.push(r);
        ultimo = r;
      });

      return dados;
    });

    return resultado.filter(r =>
      !((r.matricula || '').trim() === '' && (r.beneficiario || '').trim() === '')
    );
  } catch (err) {
    log("ERROR", `Erro ao extrair: ${err.message}`);
    return [];
  }
}

(async () => {
  const ontem = new Date();
  ontem.setDate(ontem.getDate() - 1);
  const dia  = ontem.getDate().toString().padStart(2, '0');
  const mes  = (ontem.getMonth() + 1).toString().padStart(2, '0');
  const ano  = ontem.getFullYear();
  const dataOntem   = `${dia}/${mes}/${ano}`;
  const dataArquivo = `${dia}-${mes}-${ano}`;

  log("INFO", `Buscando dados de ontem: ${dataOntem}`);

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();
  await context.route('**/*', route => {
    const tipo = route.request().resourceType();
    return ['image', 'font', 'stylesheet', 'media'].includes(tipo) ? route.abort() : route.continue();
  });
  const page = await context.newPage();

  // Login ASSIM
  log("INFO", "Fazendo login na ASSIM...");
  await page.goto('https://sirius.assim.com.br/assimcsp/autorizador/login.csp', {
    waitUntil: 'domcontentloaded', timeout: 30000
  });
  await page.selectOption('select', '52345');
  await page.fill('input[type="password"]', process.env.SENHA);
  await Promise.all([page.waitForNavigation(), page.click('text=Entrar')]);
  await page.waitForSelector('select[name="DiaFim"]');
  log("SUCCESS", "Login ASSIM OK");

  // Extração
  const urlNormal =
    `https://sirius.assim.com.br/assimcsp/autorizador/preresultado.csp?idHospital=52345&DataIni=${dataOntem}&DataFim=${dataOntem}&executor=52345&natservico=T&servico=T&especialidade=T&amb=&prefeitura=0&tuss=`;
  const urlPrefeitura =
    `https://sirius.assim.com.br/assimcsp/autorizador/preresultado.csp?idHospital=52345&DataIni=${dataOntem}&DataFim=${dataOntem}&executor=52345&natservico=T&servico=T&especialidade=T&amb=&prefeitura=1&tuss=`;

  const normal     = await extrairRelatorio(page, urlNormal);
  const prefeitura = await extrairRelatorio(page, urlPrefeitura);
  const todos      = [...normal, ...prefeitura];

  log("INFO", `Registros encontrados: ${todos.length} (normal: ${normal.length}, prefeitura: ${prefeitura.length})`);

  if (todos.length === 0) {
    log("INFO", "Nenhum registro encontrado para ontem. Encerrando.");
    await browser.close();
    process.exit(0);
  }

  // Gera Excel sem biofacial
  const dadosExcel = todos.map(({ biofacial, ...r }) => r);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dadosExcel), "Relatorio");

  const pastaRelatorios = path.join(__dirname, 'relatorios');
  if (!fs.existsSync(pastaRelatorios)) fs.mkdirSync(pastaRelatorios, { recursive: true });

  const nomeArquivo  = `relatorio_assim_${dataArquivo}.xlsx`;
  const caminhoArquivo = path.join(pastaRelatorios, nomeArquivo);
  XLSX.writeFile(wb, caminhoArquivo);
  log("SUCCESS", `Excel gerado: ${nomeArquivo}`);

  // Login Órbita
  log("INFO", "Fazendo login no Órbita...");
  await page.goto('https://cronogramauniversoaba.com.br/app_Login/');
  await page.fill('input[placeholder="Usuário"]', process.env.ORBITA_USER);
  await page.fill('input[placeholder="Senha"]', process.env.ORBITA_PASS);
  await Promise.all([page.waitForNavigation({ waitUntil: 'domcontentloaded' }), page.click('text=Entrar')]);
  await page.waitForLoadState('networkidle');
  log("SUCCESS", "Login Órbita OK");

  // Upload para Órbita
  log("INFO", `Enviando ${nomeArquivo} para Órbita com data ${dataOntem}...`);
  await page.goto('https://cronogramauniversoaba.com.br/blank_upload_registros_assim/', { waitUntil: 'networkidle' });
  await page.locator('input[type="file"]').setInputFiles(caminhoArquivo);

  const dataISO = `${ano}-${mes}-${dia}`;
  await page.locator('input[name="data_inicial"]').fill(dataISO);
  await page.locator('input[name="data_final"]').fill(dataISO);

  await page.getByRole('button', { name: 'Carregar, visualizar e linkar' }).click();
  await page.waitForLoadState('networkidle');
  await page.waitForTimeout(2000);
  log("SUCCESS", "Upload concluído");

  const botaoConfirmar = page.getByRole('button').filter({ hasText: /confirmar|finalizar|processar/i }).first();
  if (await botaoConfirmar.isVisible()) {
    await botaoConfirmar.click({ force: true });
    await page.waitForTimeout(2000);
    log("SUCCESS", "Confirmação realizada");
  } else {
    log("INFO", "Botão de confirmação não encontrado — verificar se há dados para confirmar");
  }

  await browser.close();
  log("SUCCESS", `Concluído. ${todos.length} registros de ${dataOntem} enviados ao Órbita.`);
})();
