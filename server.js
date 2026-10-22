const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const path = require('path');

const PORT = process.env.PORT || 3000;
const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });
const db = new Database(path.join(process.env.DATA_DIR || __dirname, 'data.db'));
db.pragma('journal_mode = WAL'); // cho phep nhieu nguoi doc/ghi cung luc

// Moi dong Excel luu dang JSON (data) + chuoi de tim kiem (search) -> file Excel co cot nao cung duoc
db.exec(`
  CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);
  CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    data TEXT NOT NULL,
    search TEXT NOT NULL
  );
`);

// Chuyen ve chu thuong, bo dau tieng Viet -> go "nhom" van tim ra "Nhôm"
const fold = (s) =>
  String(s)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd');

function getColumns() {
  const r = db.prepare("SELECT value FROM meta WHERE key = 'columns'").get();
  return r ? JSON.parse(r.value) : [];
}
function setColumns(cols) {
  db.prepare("INSERT INTO meta (key, value) VALUES ('columns', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").run(JSON.stringify(cols));
}

function cellToText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof Date) return isNaN(v) ? '' : v.toISOString().slice(0, 10);
  return String(v).trim();
}

const insertItem = db.prepare('INSERT INTO items (data, search) VALUES (?, ?)');
function addRow(columns, obj) {
  const values = columns.map((c) => cellToText(obj[c]));
  const record = {};
  columns.forEach((c, i) => (record[c] = values[i]));
  insertItem.run(JSON.stringify(record), fold(values.join(' ')));
}

// ---- API ----

// Tim kiem: de trong = tat ca; nhieu tu = dong chua TAT CA cac tu do (o bat ky cot nao)
app.get('/api/products', (req, res) => {
  const columns = getColumns();
  const terms = fold(req.query.q || '').split(/\s+/).filter(Boolean);
  const limit = Math.min(parseInt(req.query.limit) || 200, 1000);
  const offset = parseInt(req.query.offset) || 0;
  const where = terms.map(() => `search LIKE ? ESCAPE '\\'`).join(' AND ');
  const params = terms.map((t) => '%' + t.replace(/[\\%_]/g, (m) => '\\' + m) + '%');
  const whereSql = where ? 'WHERE ' + where : '';
  const total = db.prepare(`SELECT COUNT(*) AS n FROM items ${whereSql}`).get(...params).n;
  const rows = db
    .prepare(`SELECT data FROM items ${whereSql} ORDER BY id LIMIT ? OFFSET ?`)
    .all(...params, limit, offset)
    .map((r) => JSON.parse(r.data));
  res.json({ columns, total, rows });
});

// Them dong moi
app.post('/api/products', express.json(), (req, res) => {
  const columns = getColumns();
  if (!columns.length) return res.status(400).json({ error: 'Chưa có dữ liệu. Hãy import file Excel trước.' });
  const body = req.body || {};
  if (!columns.some((c) => cellToText(body[c]))) {
    return res.status(400).json({ error: 'Hãy nhập ít nhất một ô.' });
  }
  addRow(columns, body);
  res.json({ ok: true });
});

// Import Excel: thay toan bo du lieu bang noi dung file moi (cot lay tu dong dau tien)
app.post('/api/import', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Chưa chọn file.' });
  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });

    const isEmptyRow = (r) => !r.some((v) => cellToText(v) !== '');
    const headIdx = data.findIndex((r) => !isEmptyRow(r));
    if (headIdx < 0) return res.status(400).json({ error: 'File rỗng.' });

    const headerRow = data[headIdx];
    const body = data.slice(headIdx + 1).filter((r) => !isEmptyRow(r));
    const width = Math.max(headerRow.length, ...body.map((r) => r.length), 0);

    // Giu cot neu co ten cot hoac co du lieu
    const keep = [];
    for (let i = 0; i < width; i++) {
      if (cellToText(headerRow[i]) !== '' || body.some((r) => cellToText(r[i]) !== '')) keep.push(i);
    }
    // Dat ten cot, tranh trung ten
    const used = {};
    const columns = keep.map((i) => {
      let name = cellToText(headerRow[i]) || `Cột ${i + 1}`;
      if (used[name]) {
        used[name]++;
        name = `${name}_${used[name]}`;
      } else used[name] = 1;
      return name;
    });
    if (!columns.length) return res.status(400).json({ error: 'Không tìm thấy cột nào.' });

    db.transaction(() => {
      db.exec('DELETE FROM items');
      setColumns(columns);
      for (const r of body) {
        const obj = {};
        keep.forEach((srcIdx, k) => (obj[columns[k]] = r[srcIdx]));
        addRow(columns, obj);
      }
    })();
    res.json({ ok: true, imported: body.length, columns: columns.length });
  } catch (e) {
    res.status(400).json({ error: 'Không đọc được file: ' + e.message });
  }
});

// Tai ve file Excel hien tai
app.get('/api/export', (req, res) => {
  const columns = getColumns();
  const rows = db.prepare('SELECT data FROM items ORDER BY id').all().map((r) => JSON.parse(r.data));
  const ws = XLSX.utils.json_to_sheet(rows, { header: columns });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename="san-pham.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// Xoa toan bo du lieu (trang web goi API nay khi duoc tai lai, neu RESET_ON_LOAD = true)
app.post('/api/reset', (req, res) => {
  db.exec('DELETE FROM items');
  db.exec("DELETE FROM meta WHERE key = 'columns'");
  res.json({ ok: true });
});

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Chay tai http://localhost:${PORT}  (may khac trong mang LAN: http://<IP-may-nay>:${PORT})`);
});
