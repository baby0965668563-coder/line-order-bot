function parseOrders(text) {
  const rawLines = String(text || '').split('\n');

  const rows = [];
  let blankBefore = true;

  for (const raw of rawLines) {
    const line = raw.trim();

    if (!line) {
      blankBefore = true;
      continue;
    }

    rows.push({ line, blankBefore });
    blankBefore = false;
  }

  let currentItem = '';
  let currentPrice = 0;
  let lastPrice = 0;
  let priceMap = {};

  const itemCount = {};
  const userTotal = {};
  const details = [];

  function cleanItemName(text) {
    return clean(text)
      .replace(/\d+顆/g, '')
      .replace(/^[$＄💰]\d+$/, '')
      .trim();
  }

  function add(item, price, name, qty = 1, note = '') {
    item = cleanItemName(item);
    name = clean(name);
    note = clean(note);

    if (!item || !name || !qty) return;

    const finalItem = note ? `${item}（${note}）` : item;
    const total = Number(price || 0) * Number(qty || 1);

    itemCount[finalItem] = (itemCount[finalItem] || 0) + qty;
    userTotal[name] = (userTotal[name] || 0) + total;

    details.push({
      item: finalItem,
      rawItem: item,
      name,
      qty,
      note,
      price: Number(price || 0),
      total
    });
  }

  function getPrice(rawQty) {
    if (priceMap[rawQty] !== undefined) return priceMap[rawQty];
    if (rawQty === '半' && priceMap['0.5'] !== undefined) return priceMap['0.5'];
    return currentPrice || lastPrice || 0;
  }

  for (let i = 0; i < rows.length; i++) {
    const { line, blankBefore } = rows[i];

    if (isTrashLine(line) && !/[💰$＄]\s*\d/.test(line)) continue;

    const priceOnly = line.match(/^[💰$＄]\s*(\d{1,5})$/);
    if (priceOnly) {
      lastPrice = Number(priceOnly[1]);
      continue;
    }

    const itemSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})$/);
    if (itemSymbol) {
      currentItem = itemSymbol[1].trim();
      currentPrice = Number(itemSymbol[2]);
      lastPrice = currentPrice;
      priceMap = {};
      continue;
    }

    const inlineGroup = line.match(/^(.+?)\s*[+＋*＊]\s*(半|0\.5|\.5|\d+)\s*(.+)$/);
    if (inlineGroup && (currentPrice > 0 || lastPrice > 0)) {
      const item = inlineGroup[1].trim();
      const qtyRaw = inlineGroup[2];
      const namesText = inlineGroup[3].trim();

      const looksLikeNoteOnly = /^[（(].*[）)]$/.test(namesText);
      const hasNameSeparator = /[、,，\s]/.test(namesText);

      if (!looksLikeNoteOnly && hasNameSeparator) {
        const names = splitNames(namesText);
        const qty = parseQty(qtyRaw);
        const price = getPrice(qtyRaw);

        currentItem = item;
        currentPrice = price || lastPrice || currentPrice;
        lastPrice = currentPrice || lastPrice;

        names.forEach((n, index) => {
          const parsed = extractParenNote(n);
          add(currentItem, currentPrice, parsed.main, index === 0 ? qty : 1, parsed.note);
        });

        continue;
      }
    }

    const orderMatch = line.match(/^(.+?)[+＋*＊]\s*(半|0\.5|\.5|\d+)(.*)$/);
    if (orderMatch && currentItem) {
      const nameRaw = orderMatch[1].trim();
      const rawQty = orderMatch[2];
      const extraRaw = String(orderMatch[3] || '').trim();

      const parsedName = extractParenNote(nameRaw);
      const parsedExtra = extractParenNote(extraRaw);

      add(
        currentItem,
        getPrice(rawQty),
        parsedName.main,
        parseQty(rawQty),
        parsedName.note || parsedExtra.note || clean(extraRaw)
      );

      continue;
    }

    if (!/[+＋*＊]/.test(line)) {
      const parsed = extractParenNote(line);

      if (
        blankBefore &&
        lastPrice > 0 &&
        parsed.main.length <= 12 &&
        !isTrashLine(parsed.main)
      ) {
        currentItem = parsed.main;
        currentPrice = lastPrice;
        continue;
      }

      if (currentItem && (currentPrice > 0 || lastPrice > 0) && !isTrashLine(line)) {
        add(currentItem, currentPrice || lastPrice, parsed.main, 1, parsed.note);
        continue;
      }

      continue;
    }
  }

  return { itemCount, userTotal, details };
}