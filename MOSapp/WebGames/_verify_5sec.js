const { chromium } = require("playwright");

const sites = [
  ["judge", "https://mos-excel-3sec-judge.netlify.app/"],
  ["judge", "https://mos-word-3sec-judge.netlify.app/"],
  ["judge", "https://mos-ppt-3sec-judge.netlify.app/"],
  ["snipe", "https://mos-excel-3sec-snipe.netlify.app/"],
  ["snipe", "https://mos-word-3sec-snipe.netlify.app/"],
  ["snipe", "https://mos-ppt-3sec-snipe.netlify.app/"]
];

(async () => {
  const browser = await chromium.launch({ headless: true });
  for (const [kind, url] of sites) {
    const page = await browser.newPage();
    await page.addInitScript(() => localStorage.setItem("mos-snipe-time-limit", "3000"));
    await page.goto(url + "?v=" + Date.now(), { waitUntil: "networkidle" });
    const intro = await page.locator("#start-screen").innerText();
    if (!intro.includes("たったの5秒")) throw new Error(url + " copy\n" + intro);
    if (intro.includes("3秒")) throw new Error(url + " still says 3秒\n" + intro);
    if (await page.locator(".time-btn").count()) throw new Error(url + " still has time buttons");
    await page.click("#start-btn");
    await page.waitForFunction(() => {
      const fill = document.getElementById("fuse-fill");
      return fill && fill.style.transition.includes("5000ms");
    });
    const transition = await page.locator("#fuse-fill").evaluate((el) => el.style.transition);
    if (kind === "judge" && url.includes("excel")) {
      const t0 = Date.now();
      await page.waitForFunction(() => document.getElementById("badge-hold").textContent.includes("タイムアップ"));
      const elapsed = Date.now() - t0;
      if (elapsed < 4700 || elapsed > 5600) throw new Error(url + " timeout " + elapsed);
      console.log("ok", url, transition, "wall", elapsed);
    } else if (kind === "snipe" && url.includes("excel")) {
      const t0 = Date.now();
      await page.waitForFunction(() => document.getElementById("verdict").textContent.includes("タイムアップ"));
      const elapsed = Date.now() - t0;
      if (elapsed < 4700 || elapsed > 5600) throw new Error(url + " timeout " + elapsed);
      console.log("ok", url, transition, "wall", elapsed);
    } else {
      console.log("ok", url, transition);
    }
    await page.close();
  }
  await browser.close();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
