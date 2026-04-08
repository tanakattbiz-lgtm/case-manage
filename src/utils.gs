function genId(prefix) {
  return prefix + '-' + Date.now().toString(36).toUpperCase();
}

function nowStr() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

function todayStr() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
}

function sheetToObjects(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  return data.slice(1).map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => {
      obj[h] = (row[j] instanceof Date)
        ? Utilities.formatDate(row[j], 'Asia/Tokyo', 'yyyy/MM/dd')
        : row[j];
    });
    return obj;
  });
}
