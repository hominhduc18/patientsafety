// Nhập biểu mẫu từ file Excel (.xlsx/.xls/.csv), PDF, Word (.docx)
// Trả về chuỗi theo định dạng ô soạn biểu mẫu:
//   Dòng 1: tên biểu mẫu
//   Mỗi dòng sau: Nội dung | Sự cố | Mức ACCI | Phòng ngừa | Khắc phục
(() => {
  const LIBS = {
    XLSX: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    pdfjsLib: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js',
    mammoth: 'https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js',
    qrcode: 'https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.min.js'
  };
  const PDF_WORKER = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';

  const load = g => window[g] ? Promise.resolve(window[g]) : new Promise((ok, no) => {
    const s = document.createElement('script');
    s.src = LIBS[g];
    s.onload = () => ok(window[g]);
    s.onerror = () => no(new Error('Không tải được thư viện đọc file (cần kết nối mạng)'));
    document.head.appendChild(s);
  });

  const norm = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[đĐ]/g, 'd').toLowerCase().replace(/\s+/g, ' ').trim();
  const clean = s => String(s ?? '').replace(/\|/g, '/').replace(/\s+/g, ' ').trim();

  // Nhận diện cột theo tiêu đề (thứ tự kiểm tra quan trọng)
  const RULES = [
    ['stt', /^(stt|tt|so tt|so thu tu|#)$/],
    ['acci', /acci|muc do|cap do|phan loai|phan muc/],
    ['prev', /phong ngua|du phong|bien phap phong/],
    ['fix', /khac phuc|xu ly|xu tri|hanh dong/],
    ['inc', /su co|nguy co|rui ro|sai sot/],
    ['t', /noi dung|buoc|tieu chi|hang muc|cong viec|muc kiem|quy trinh|yeu cau|thao tac|hoat dong|kiem tra/],
    ['acci', /^muc$/]
  ];
  const fieldOf = text => {
    const n = norm(text);
    if (!n || n.length > 60) return null;
    for (const [f, re] of RULES) if (re.test(n)) return f;
    return null;
  };

  // rows: [[{x, w, text}]] — x là vị trí cột (Excel/Word: chỉ số cột, PDF: toạ độ)
  function findHeader(rows) {
    let best = null;
    rows.slice(0, 40).forEach((row, ri) => {
      const cols = {};
      row.forEach(c => { const f = fieldOf(c.text); if (f && !(f in cols)) cols[f] = c.x + (c.w || 0) / 2; });
      const score = Object.keys(cols).filter(k => k !== 'stt').length;
      if (score >= 2 && (!best || score > best.score)) best = { ri, cols, score };
    });
    return best;
  }

  const stripNum = s => s.replace(/^\s*(\d+(\.\d+)*|[a-z])\s*[.)/:-]\s*/i, '').trim();
  const isNumberLead = s => /^\s*\d+(\.\d+)*\s*[.)/:-]?\s+\S/.test(s);

  function build(rows, fallbackName) {
    rows = rows.map(r => r.map(c => ({ ...c, text: clean(c.text) })).filter(c => c.text)).filter(r => r.length);
    if (!rows.length) throw new Error('File không có nội dung chữ');
    const h = findHeader(rows);
    let name = fallbackName, items = [];

    if (h) {
      const above = rows.slice(0, h.ri).map(r => r.map(c => c.text).join(' ')).filter(s => s.length > 4);
      if (above.length) name = above.sort((a, b) => /bang kiem|quy trinh|bieu mau/.test(norm(b)) - /bang kiem|quy trinh|bieu mau/.test(norm(a)))[0];
      const centers = Object.entries(h.cols);
      const colOf = c => {
        const cx = c.x + (c.w || 0) / 2;
        let bf = null, bd = Infinity;
        for (const [f, x] of centers) { const d = Math.abs(x - cx); if (d < bd) { bd = d; bf = f; } }
        return bf;
      };
      const data = rows.slice(h.ri + 1).map(r => {
        const o = {};
        r.forEach(c => { const f = colOf(c); o[f] = o[f] ? o[f] + ' ' + c.text : c.text; });
        return o;
      }).filter(o => !Object.values(o).every(v => fieldOf(v))); // bỏ dòng tiêu đề lặp lại ở trang sau
      const useStt = 'stt' in h.cols && data.some(o => o.stt && /^\d/.test(o.stt));
      let cur = null;
      for (const o of data) {
        const onlyStt = o.stt && !o.t && !o.inc && !o.acci && !o.prev && !o.fix;
        if (onlyStt) continue; // dòng tiêu đề nhóm (I, II, ...)
        if (useStt && o.stt && !/^\d/.test(o.stt)) { cur = null; continue; } // "I. CHUẨN BỊ" = tiêu đề nhóm
        const isNew = useStt ? !!(o.stt && /^\d/.test(o.stt)) : !!o.t;
        if (useStt && !isNew && !cur) continue;
        if (isNew || !cur) {
          cur = { t: o.t || '', inc: o.inc || '', acci: o.acci || '', prev: o.prev || '', fix: o.fix || '' };
          items.push(cur);
        } else {
          for (const f of ['t', 'inc', 'acci', 'prev', 'fix']) if (o[f]) cur[f] = (cur[f] ? cur[f] + ' ' : '') + o[f];
        }
      }
    } else {
      // Không có bảng tiêu đề: mỗi dòng đánh số là 1 mục, dòng không số nối vào mục trước
      const lines = rows.map(r => r.map(c => c.text).join(' '));
      const numbered = lines.some(isNumberLead);
      let cur = null;
      lines.forEach((ln, i) => {
        if (i === 0 && !isNumberLead(ln)) { name = ln; return; }
        if (!numbered || isNumberLead(ln) || !cur) { cur = { t: ln, inc: '', acci: '', prev: '', fix: '' }; items.push(cur); }
        else cur.t += ' ' + ln;
      });
    }

    items = items.map(i => ({ ...i, t: stripNum(i.t), acci: i.acci.replace(/^m[uứ]c\s*/i, '') })).filter(i => i.t);
    if (!items.length) throw new Error('Không tìm thấy mục kiểm nào trong file');
    return clean(name) + '\n' + items.map(i => [i.t, i.inc, i.acci, i.prev, i.fix].join(' | ')).join('\n');
  }

  async function fromExcel(file, base) {
    const X = await load('XLSX');
    const wb = X.read(await file.arrayBuffer(), { type: 'array' });
    const sheets = wb.SheetNames.map(n => X.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '', raw: false })
      .map(r => r.map((text, x) => ({ x, text }))));
    const pick = sheets.find(r => findHeader(r.map(x => x.filter(c => clean(c.text))))) || sheets.find(r => r.length) || [];
    return build(pick, base);
  }

  async function fromPdf(file, base) {
    const P = await load('pdfjsLib');
    P.GlobalWorkerOptions.workerSrc = PDF_WORKER;
    const doc = await P.getDocument({ data: await file.arrayBuffer() }).promise;
    const rows = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const items = (await page.getTextContent()).items.filter(i => i.str && i.str.trim());
      const lines = [];
      items.forEach(i => {
        const y = i.transform[5], x = i.transform[4], hgt = Math.abs(i.transform[3]) || 10;
        let ln = lines.find(l => Math.abs(l.y - y) < hgt * 0.5);
        if (!ln) lines.push(ln = { y, hgt, its: [] });
        ln.its.push({ x, w: i.width, s: i.str });
      });
      lines.sort((a, b) => b.y - a.y).forEach(l => {
        l.its.sort((a, b) => a.x - b.x);
        const cells = [];
        l.its.forEach(it => {
          const last = cells[cells.length - 1];
          if (last && it.x - (last.x + last.w) < l.hgt * 1.2) { last.text += (it.x - (last.x + last.w) > 1 ? ' ' : '') + it.s; last.w = it.x + it.w - last.x; }
          else cells.push({ x: it.x, w: it.w, text: it.s });
        });
        rows.push(cells);
      });
    }
    if (!rows.length) throw new Error('PDF dạng ảnh scan, không đọc được chữ. Hãy dùng file Excel/Word hoặc PDF có chữ.');
    return build(rows, base);
  }

  async function fromWord(file, base) {
    const M = await load('mammoth');
    const { value } = await M.convertToHtml({ arrayBuffer: await file.arrayBuffer() });
    const dom = new DOMParser().parseFromString(value, 'text/html');
    const tables = [...dom.querySelectorAll('table')].map(t => [...t.querySelectorAll('tr')]
      .map(tr => [...tr.children].map((td, x) => ({ x, text: td.textContent }))));
    const title = [...dom.body.children].find(e => e.tagName !== 'TABLE' && clean(e.textContent).length > 4);
    const tbl = tables.find(r => findHeader(r.map(x => x.filter(c => clean(c.text)))));
    if (tbl) return build((title ? [[{ x: 0, text: title.textContent }]] : []).concat(tbl), base);
    return build([...dom.querySelectorAll('p,li,h1,h2,h3,h4')].map(e => [{ x: 0, text: e.textContent }]), base);
  }

  window.loadQR = () => load('qrcode');

  window.parseFormFile = async file => {
    const base = file.name.replace(/\.[^.]+$/, '');
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    if (['xlsx', 'xls', 'xlsm', 'csv', 'ods'].includes(ext)) return fromExcel(file, base);
    if (ext === 'pdf') return fromPdf(file, base);
    if (ext === 'docx') return fromWord(file, base);
    if (ext === 'doc') throw new Error('File .doc cũ: hãy mở bằng Word và lưu lại thành .docx');
    throw new Error('Chỉ hỗ trợ Excel (.xlsx, .xls, .csv), PDF, Word (.docx)');
  };
})();
