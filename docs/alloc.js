// Випадковий розподіл з урахуванням допусків. Людина має поле allowed: Set індексів секторів.
function rnd(n) {
  const a = new Uint32Array(1), lim = Math.floor(0x100000000 / n) * n;
  do crypto.getRandomValues(a); while (a[0] >= lim);
  return a[0] % n;
}
export function shuffle(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) { const j = rnd(i + 1); [a[i], a[j]] = [a[j], a[i]]; }
  return a;
}

// Чи можна розкласти людей по секторах, не перевищивши місткість (парування з місткістю).
function feasible(people, caps) {
  const asg = caps.map(() => []);
  const place = (p, seen) => {
    for (const s of p.allowed) {
      if (seen.has(s)) continue;
      seen.add(s);
      if (asg[s].length < caps[s]) { asg[s].push(p); return true; }
      for (let i = 0; i < asg[s].length; i++) if (place(asg[s][i], seen)) { asg[s][i] = p; return true; }
    }
    return false;
  };
  return people.every(p => place(p, new Set()));
}

// tiers: [обов'язкові, пріоритет 1, пріоритет 2] — люди {allowed:Set, ...}
export function allocate(tiers, sectorCount, sessions, perSector) {
  const caps = Array(sectorCount).fill(sessions * perSector);
  const total = sectorCount * sessions * perSector;
  const order = tiers.flatMap((t, i) => shuffle(t).map(p => ({ p, tier: i })));
  const selected = [], rest = [];
  for (const { p, tier } of order) {
    p.tier = tier;
    if (selected.length < total && feasible([...selected, p], caps)) selected.push(p); else rest.push(p);
  }
  // випадково обираємо сектор кожній людині, зберігаючи можливість розставити решту
  const left = caps.slice();
  const bySector = caps.map(() => []);
  const todo = shuffle(selected);
  todo.forEach((p, i) => {
    for (const s of shuffle([...p.allowed].filter(x => left[x] > 0))) {
      left[s]--;
      if (feasible(todo.slice(i + 1), left)) { bySector[s].push(p); return; }
      left[s]++;
    }
  });
  // grid[сесія][сектор] = масив із perSector місць (null — вільне)
  const grid = Array.from({ length: sessions }, () => Array.from({ length: sectorCount }, () => Array(perSector).fill(null)));
  bySector.forEach((list, s) => shuffle(list).forEach((p, i) => { grid[i % sessions][s][Math.floor(i / sessions)] = p; }));
  return { grid, rest, selectedCount: selected.length, total };
}
