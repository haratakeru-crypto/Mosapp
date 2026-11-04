const { chromium } = require("playwright");

const sites = [
  "https://mos-excel-3sec-snipe.netlify.app/",
  "https://mos-word-3sec-snipe.netlify.app/",
  "https://mos-ppt-3sec-snipe.netlify.app/"
];

(async () => {
  const browser = await chromium.launch({ headless: true });
  for (const url of sites) {
    const page = await browser.newPage();
    const stamp = Date.now();
    await page.goto(url + "?v=" + stamp, { waitUntil: "networkidle" });
    const intro = await page.locator("#start-screen").innerText();
    if (!intro.includes("1周は10問")) throw new Error(url + " intro missing");
    const qs = await page.evaluate(async (base) => {
      const text = await (await fetch(base + "questions.js?v=" + Date.now())).text();
      return JSON.parse(text.replace(/^[\s\S]*?const QUESTIONS = /, "").replace(/;\s*$/, ""));
    }, url);
    await page.click("#start-btn");
    await page.waitForFunction(() => document.getElementById("progress-text").textContent.includes("1 / 10"));
    const seen = new Set();
    const kinds = [];
    for (let i = 0; i < 10; i++) {
      const qText = (await page.locator("#card-question").innerText()).replace(/\s+/g, " ").trim();
      const left = (await page.locator("#hint-left").innerText()).trim();
      const right = (await page.locator("#hint-right").innerText()).trim();
      const matches = qs.filter((c) =>
        c.q.replace(/\s+/g, " ").trim() === qText &&
        ((c.correct === left && c.wrong === right) || (c.correct === right && c.wrong === left))
      );
      if (matches.length !== 1) {
        throw new Error(url + " q" + (i + 1) + " matches " + matches.length + " L=" + left + " R=" + right);
      }
      const card = matches[0];
      if (seen.has(card.sourceId)) throw new Error(url + " duplicate " + card.sourceId);
      seen.add(card.sourceId);
      kinds.push(card.id.endsWith("-btn") ? "btn" : "tab");
      const correctSide = left === card.correct ? "#choice-left" : "#choice-right";
      const wrongSide = correctSide === "#choice-left" ? "#choice-right" : "#choice-left";
      await page.click(i < 3 ? wrongSide : correctSide);
      if (i < 9) {
        const next = i + 2;
        await page.waitForFunction((n) => document.getElementById("progress-text").textContent.startsWith(n + " /"), next);
      }
    }
    await page.waitForFunction(() => {
      const el = document.getElementById("round-screen");
      return el && el.style.display !== "none" && document.getElementById("round-desc").textContent.includes("新しい");
    });
    const desc = (await page.locator("#round-desc").innerText()).trim();
    if (desc !== "外した 3 問、新しい 7 問") throw new Error(url + " banner " + desc);
    console.log("ok", url, desc, kinds.join(","));
    await page.close();
  }
  await browser.close();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
