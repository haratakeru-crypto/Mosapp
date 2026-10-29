const { chromium } = require("playwright");

async function measure(url, pick3s) {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  await page.goto(url, { waitUntil: "networkidle" });
  const meta = await page.evaluate(() => {
    const html = document.documentElement.innerHTML;
    return {
      timeLimit: (html.match(/TIME_LIMIT\s*=\s*(\d+)/) || [])[1] || null,
      timeLimitLet: (html.match(/let timeLimit = (\d+)/) || [])[1] || null,
      feedback: (html.match(/FEEDBACK_MS\s*=\s*(\d+)/) || [])[1] || null
    };
  });
  if (pick3s) {
    await page.click("#time-3");
  }
  await page.click("#start-btn");
  await page.waitForTimeout(80);
  const started = await page.evaluate(() => performance.now());
  await page.waitForFunction(() => {
    const hold = document.getElementById("badge-hold");
    const verdict = document.getElementById("verdict");
    if (hold && /タイムアップ/.test(hold.textContent || "") && hold.style.opacity === "1") return true;
    if (verdict && verdict.classList.contains("show") && /タイムアップ/.test(verdict.textContent || "")) return true;
    return false;
  }, { timeout: 15000 });
  const ended = await page.evaluate(() => performance.now());
  await browser.close();
  return { url, ...meta, timeoutMs: Math.round(ended - started) };
}

(async () => {
  const rows = [];
  rows.push(await measure("https://mos-excel-3sec-judge.netlify.app/", false));
  rows.push(await measure("https://mos-excel-3sec-snipe.netlify.app/", true));
  rows.push(await measure("https://mos-word-3sec-judge.netlify.app/", false));
  console.log(JSON.stringify(rows, null, 2));
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
