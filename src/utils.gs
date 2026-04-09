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
      if (row[j] instanceof Date) {
        const dateValue = row[j];
        const hasTime = dateValue.getHours() !== 0
          || dateValue.getMinutes() !== 0
          || dateValue.getSeconds() !== 0
          || dateValue.getMilliseconds() !== 0;
        obj[h] = Utilities.formatDate(
          dateValue,
          'Asia/Tokyo',
          hasTime ? 'yyyy/MM/dd HH:mm:ss' : 'yyyy/MM/dd'
        );
        return;
      }
      obj[h] = row[j];
    });
    return obj;
  });
}
