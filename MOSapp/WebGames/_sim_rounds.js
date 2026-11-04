const fs = require("fs");
const vm = require("vm");

function load(dir) {
  const code = fs.readFileSync(require("path").join(__dirname, dir, "questions.js"), "utf8");
  const ctx = {};
  vm.runInNewContext(code.replace("const QUESTIONS", "QUESTIONS"), ctx);
  return ctx.QUESTIONS;
}

function shuffle(arr) {
  return arr.slice().sort(() => Math.random() - 0.5);
}

function sim(name) {
  let pool = shuffle(load(name));
  function take(blocked, count) {
    const picked = [];
    const rest = [];
    const blockedIds = new Set(blocked.map((q) => q.sourceId));
    pool.forEach((card) => {
      const clash = blockedIds.has(card.sourceId) || picked.some((q) => q.sourceId === card.sourceId);
      if (picked.length < count && !clash) picked.push(card);
      else rest.push(card);
    });
    pool = rest;
    return picked;
  }
  let round = take([], 10);
  let seen = new Set();
  let n = 1;
  while (round.length) {
    const ids = round.map((q) => q.sourceId);
    if (new Set(ids).size !== ids.length) throw new Error(name + " duplicate in round " + n);
    round.forEach((q) => seen.add(q.id));
    const misses = n === 1 ? round.slice(0, 3) : [];
    const fresh = take(misses, Math.max(0, 10 - misses.length));
    if (misses.length === 0 && fresh.length === 0) break;
    round = misses.concat(fresh);
    n++;
    if (n > 80) throw new Error("too many rounds");
  }
  const all = load(name);
  if (seen.size !== all.length) throw new Error(name + " seen " + seen.size + " of " + all.length);
  console.log(name, "rounds", n, "cards", all.length, "ok");
}

["excel-3sec-snipe", "word-3sec-snipe", "ppt-3sec-snipe"].forEach(sim);
