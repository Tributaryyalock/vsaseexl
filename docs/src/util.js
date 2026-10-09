// Раздел с автонумерацией пунктов: строки получают номер «n.k»,
// { p: 'текст' } — абзац без номера, остальные объекты передаются как есть.
const S = (n, title, items, opts = {}) => {
  let k = 0;
  const blocks = items.map((it) => {
    if (typeof it === 'string') { k += 1; return [`${n}.${k}`, it]; }
    if (it && it.p) return it.p;
    return it;
  });
  return { title: n ? `${n}. ${title}` : title, blocks, ...opts };
};

module.exports = { S };
