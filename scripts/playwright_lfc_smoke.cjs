const path = require("path");

process.env.PLAYWRIGHT_BROWSERS_PATH = path.join(__dirname, "..", ".playwright-browsers");
const outputDir = path.join(__dirname, "..", "output", "playwright");
const tmpDir = path.join(outputDir, "tmp");
process.env.TEMP = tmpDir;
process.env.TMP = tmpDir;

const { chromium } = require("playwright");
const http = require("http");
const fs = require("fs");

async function main() {
  const root = path.resolve(__dirname, "..");
  const docsRoot = path.join(root, "docs");
  const screenshotPath = path.join(outputDir, "iota-lfc-smoke.png");

  fs.mkdirSync(outputDir, { recursive: true });
  fs.mkdirSync(tmpDir, { recursive: true });

  const server = await startStaticServer(docsRoot);
  const target = `http://127.0.0.1:${server.port}/iota-development-workspace.html#lfc`;

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  const consoleErrors = [];
  page.on("console", (message) => {
    if (message.type() === "error") consoleErrors.push(message.text());
  });
  page.on("pageerror", (error) => consoleErrors.push(error.message));

  await page.goto(target, { waitUntil: "load" });
  await page.waitForSelector("#financeView:not([hidden]) #marketLenderFilter");
  await page.screenshot({ path: screenshotPath, fullPage: false });

  async function readMarketState() {
    await page.waitForFunction(() => document.querySelectorAll("#marketLenderFilter option").length > 0);
    return page.evaluate(() => ({
      meta: document.querySelector("#marketNewsMeta")?.textContent?.trim() || "",
      options: [...document.querySelectorAll("#marketLenderFilter option")].map((option) =>
        option.textContent.trim()
      ),
      visibleLenders: [...document.querySelectorAll("#marketWatchGrid .row-title")].map((item) =>
        item.textContent.trim()
      ),
    }));
  }

  const totalMarket = await readMarketState();
  await page.click('.project-btn[data-project="427"]');
  await page.waitForTimeout(250);
  const oneMarket = await readMarketState();
  await page.click('.project-btn[data-project="816"]');
  await page.waitForTimeout(250);
  const twoMarket = await readMarketState();

  await page.fill('#lfcLogForm textarea[name="text"]', `QA LFC log ${Date.now()}`);
  await page.click('#lfcLogForm button[type="submit"]');
  await page.waitForTimeout(250);
  const logSubmitOk = await page.evaluate(() => {
    const text = document.querySelector("#lfcLogList")?.textContent || "";
    return text.includes("QA LFC log");
  });

  await page.click("#loanTable tr[data-loan-id]");
  await page.waitForTimeout(250);
  const loanDrawerOpen = await page.evaluate(() => {
    const drawer = document.querySelector("#financeDrawerBackdrop");
    return drawer?.classList.contains("open") && (document.querySelector("#financeDrawerBody")?.textContent || "").includes("Loan 조건");
  });
  await page.keyboard.press("Escape");
  await page.waitForTimeout(100);

  const result = await page.evaluate(() => {
    const cards = [...document.querySelectorAll(".lfc-cell-card")];
    const rects = cards.map((card) => card.getBoundingClientRect());
    const firstTop = rects[0] ? Math.round(rects[0].top) : null;
    return {
      title: document.querySelector("#financeView h1")?.textContent?.trim() || "",
      cardCount: cards.length,
      cardsOneRow: rects.length === 4 && rects.every((rect) => Math.abs(Math.round(rect.top) - firstTop) <= 2),
      horizontalOverflow: document.documentElement.scrollWidth > document.documentElement.clientWidth,
      forbiddenTextFound: ["Target Vehicle", "약정조건·만기 관리", "대주별 관리 및 이슈"].some((text) =>
        document.body.innerText.includes(text)
      ),
    };
  });

  await browser.close();
  await server.close();

  const marketFilter = {
    totalCount: totalMarket.options.length,
    oneCount: oneMarket.options.length,
    twoCount: twoMarket.options.length,
    oneOptions: oneMarket.options,
    twoOptions: twoMarket.options,
    oneVisibleLenders: oneMarket.visibleLenders,
    twoVisibleLenders: twoMarket.visibleLenders,
    oneHas427Lender: oneMarket.options.includes("KB국민은행"),
    twoHas816Lender: twoMarket.options.includes("메리츠화재") || twoMarket.options.includes("NH투자증권"),
    twoExcludes427OnlyLender: !twoMarket.options.includes("KB국민은행"),
    listsDiffer: JSON.stringify(oneMarket.options) !== JSON.stringify(twoMarket.options),
  };

  const interactions = { logSubmitOk, loanDrawerOpen };
  const payload = { ...result, marketFilter, interactions, consoleErrors, screenshotPath };
  console.log(JSON.stringify(payload, null, 2));

  if (!result.title.includes("LFC")) process.exitCode = 1;
  if (result.cardCount !== 4) process.exitCode = 1;
  if (!result.cardsOneRow) process.exitCode = 1;
  if (result.horizontalOverflow) process.exitCode = 1;
  if (result.forbiddenTextFound) process.exitCode = 1;
  if (!marketFilter.oneHas427Lender) process.exitCode = 1;
  if (!marketFilter.twoHas816Lender) process.exitCode = 1;
  if (!marketFilter.twoExcludes427OnlyLender) process.exitCode = 1;
  if (!marketFilter.listsDiffer) process.exitCode = 1;
  if (!interactions.logSubmitOk) process.exitCode = 1;
  if (!interactions.loanDrawerOpen) process.exitCode = 1;
  if (consoleErrors.length) process.exitCode = 1;
}

function startStaticServer(root) {
  const mime = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".png": "image/png",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".zip": "application/zip",
    ".md": "text/markdown; charset=utf-8",
  };
  const server = http.createServer((req, res) => {
    const rawUrl = new URL(req.url || "/", "http://127.0.0.1");
    const cleanPath = rawUrl.pathname === "/" ? "/iota-development-workspace.html" : rawUrl.pathname;
    const candidate = path.resolve(root, `.${decodeURIComponent(cleanPath)}`);
    if (!candidate.startsWith(root)) {
      res.writeHead(403);
      res.end("Forbidden");
      return;
    }
    fs.readFile(candidate, (error, data) => {
      if (error) {
        res.writeHead(404);
        res.end("Not found");
        return;
      }
      res.writeHead(200, { "content-type": mime[path.extname(candidate).toLowerCase()] || "application/octet-stream" });
      res.end(data);
    });
  });
  return new Promise((resolve, reject) => {
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      resolve({
        port: server.address().port,
        close: () => new Promise((done) => server.close(done)),
      });
    });
  });
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
