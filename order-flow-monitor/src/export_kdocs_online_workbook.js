const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright-core");

const WORKDIR = "C:\\Users\\Administrator\\Documents\\Codex\\2026-05-28\\sqlserver";
const DEFAULT_URL = "https://www.kdocs.cn/l/cjC1KSaw5Uw5";
const EDGE = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function argValue(name, fallback) {
  const index = process.argv.indexOf(name);
  return index >= 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}

async function clickFirst(page, candidates, timeout = 4000) {
  for (const candidate of candidates) {
    const locator =
      candidate.kind === "text"
        ? page.getByText(candidate.value, { exact: candidate.exact ?? false })
        : candidate.kind === "role"
          ? page.getByRole(candidate.role, { name: candidate.value })
          : page.locator(candidate.value);
    try {
      await locator.first().waitFor({ state: "visible", timeout });
      if (candidate.hover) {
        await locator.first().hover({ timeout });
      } else {
        await locator.first().click({ timeout });
      }
      return candidate;
    } catch (_) {
      // Try the next known WPS/KDocs UI variant.
    }
  }
  return null;
}

async function openMainMenu(page) {
  const clicked = await clickFirst(page, [
    { kind: "css", value: 'button:has-text("☰")' },
    { kind: "css", value: '[aria-label*="菜单"]' },
    { kind: "css", value: '.docs-icon-menu, .menu, [class*="menu"]' },
    { kind: "css", value: 'div[role="button"]:near(:text("订单流转"))' },
    { kind: "text", value: "文件" },
    { kind: "role", role: "button", value: /文件/ },
    { kind: "css", value: '[aria-label*="文件"]' },
  ], 5000);
  if (clicked) return clicked;

  await page.mouse.click(101, 19);
  await page.waitForTimeout(1200);
  return { kind: "coordinate", value: "101,19" };
}

async function exportWorkbook() {
  const url = process.env.KDOCS_URL || argValue("--url", DEFAULT_URL);
  const exportDir = process.env.KDOCS_EXPORT_DIR || argValue("--export-dir", path.join(WORKDIR, "order_flow_exports"));
  const output = process.env.KDOCS_EXPORT_OUTPUT || argValue("--output", "");
  const userDataDir = process.env.KDOCS_BROWSER_PROFILE || argValue("--profile", path.join(WORKDIR, ".kdocs-playwright-profile"));
  const headless = !(process.argv.includes("--visible") || process.env.KDOCS_VISIBLE === "1");

  ensureDir(exportDir);
  ensureDir(userDataDir);
  if (output) ensureDir(path.dirname(output));

  const context = await chromium.launchPersistentContext(userDataDir, {
    headless,
    executablePath: EDGE,
    acceptDownloads: true,
    downloadsPath: exportDir,
    viewport: { width: 1440, height: 900 },
    locale: "zh-CN",
    args: ["--no-proxy-server", "--disable-system-proxy"],
  });

  try {
    const page = context.pages()[0] || await context.newPage();
    page.setDefaultTimeout(15000);
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });
    await page.waitForLoadState("networkidle", { timeout: 30000 }).catch(() => {});

    const loginHints = await page.getByText(/微信快捷登录|扫码|立即登录|账号登录/).count().catch(() => 0);
    if (loginHints > 0) {
      throw new Error("kdocs_login_required");
    }

    console.log("save_status=saving");
    await page.keyboard.press("Control+s").catch(() => {});
    await page.waitForTimeout(3000);
    console.log("save_status=waited");

    await openMainMenu(page);
    await page.waitForTimeout(600);

    const clicked = await clickFirst(page, [
      { kind: "text", value: "导出" },
      { kind: "text", value: "导出为" },
      { kind: "text", value: "下载" },
      { kind: "text", value: "另存为" },
      { kind: "css", value: '[role="menuitem"]:has-text("导出")' },
      { kind: "css", value: '[role="menuitem"]:has-text("下载")' },
    ], 6000);
    if (!clicked) {
      const screenshot = path.join(exportDir, "kdocs_export_no_download_menu.png");
      await page.screenshot({ path: screenshot, fullPage: true }).catch(() => {});
      throw new Error(`kdocs_download_menu_not_found screenshot=${screenshot}`);
    }
    await page.waitForTimeout(400);

    const downloadPromise = page.waitForEvent("download", { timeout: 120000 }).catch(() => null);
    const xlsxClicked = await clickFirst(page, [
      { kind: "text", value: "Excel" },
      { kind: "text", value: "xlsx" },
      { kind: "text", value: "XLSX" },
      { kind: "text", value: "导出为 Excel" },
      { kind: "text", value: "下载为 Excel" },
    ], 5000);
    if (!xlsxClicked) {
      await page.keyboard.press("Enter").catch(() => {});
    }

    const download = await downloadPromise;
    const suggested = download.suggestedFilename();
    const target = output || path.join(
      exportDir,
      suggested && suggested.toLowerCase().endsWith(".xlsx")
        ? suggested
        : `kdocs_export_${new Date().toISOString().replace(/[:.]/g, "-")}.xlsx`,
    );
    await download.saveAs(target);
    console.log(`export_status=downloaded`);
    console.log(`export_path=${target}`);
    console.log(`export_suggested_filename=${suggested}`);
    return target;
  } finally {
    await context.close();
  }
}

exportWorkbook().then(() => {
  process.exit(0);
}).catch((error) => {
  console.error("export_status=failed");
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
