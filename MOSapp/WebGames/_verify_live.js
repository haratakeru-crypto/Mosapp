const { chromium } = require("playwright");

const sites = [
  { url: "https://mos-word-3sec-judge.netlify.app/", title: "3秒ジャッジ！ Word", start: true },
  { url: "https://mos-ppt-3sec-judge.netlify.app/", title: "3秒ジャッジ！ PowerPoint", start: true },
  { url: "https://mos-word-3sec-snipe.netlify.app/", title: "3秒スナイプ！ Word", start: true, timePick: true },
  { url: "https://mos-ppt-3sec-snipe.netlify.app/", title: "3秒スナイプ！ PowerPoint", start: true, timePick: true }
];

(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  for (const s of sites) {
    await page.goto(s.url, { waitUntil: "networkidle" });
    const t = await page.title();
    const startVisible = await page.locator("#start-screen").isVisible();
    console.log(JSON.stringify({ url: s.url, title: t, startVisible }));
    if (t !== s.title) throw new Error("title mismatch " + t);
    await page.click("#start-btn");
    await page.waitForTimeout(500);
    const q = (await page.locator("#card-question").textContent() || "").trim();
    const progress = (await page.locator("#progress-text").textContent() || "").trim();
    const startGone = !(await page.locator("#start-screen").isVisible());
    console.log("  after start", { startGone, progress, q: q.slice(0, 80) });
    if (!q || q === "問題文がロードされます") throw new Error("no question " + s.url);
  }
  await browser.close();
  console.log("ok");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
