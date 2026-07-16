const {
  Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType,
  BorderStyle, WidthType, ShadingType,
} = require('docx');

const CW = 9026; // A4 content width with 1" margins

const border = { style: BorderStyle.SINGLE, size: 1, color: '999999' };
const borders = { top: border, bottom: border, left: border, right: border };

function t(text, opts = {}) { return new TextRun({ text, ...opts }); }

function p(text, opts = {}) {
  return new Paragraph({ children: [t(text, opts.run || {})], spacing: { after: 120 }, ...(opts.para || {}) });
}

function pRuns(runs, opts = {}) {
  return new Paragraph({ children: runs, spacing: { after: 120 }, ...(opts.para || {}) });
}

function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [t(text)] }); }
function h4(text) { return new Paragraph({ heading: HeadingLevel.HEADING_4, children: [t(text)] }); }

function bullet(text) {
  return new Paragraph({ numbering: { reference: 'bullets', level: 0 }, children: [t(text)], spacing: { after: 80 } });
}

function cell(content, width, opts = {}) {
  const paras = Array.isArray(content) ? content : [new Paragraph({ children: [t(String(content), { bold: !!opts.bold, size: opts.size || 20 })], spacing: { after: 0 } })];
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    children: paras,
  });
}

function table(headers, rows, widths) {
  const w = widths || headers.map(() => Math.floor(CW / headers.length));
  const diff = CW - w.reduce((a, b) => a + b, 0);
  w[w.length - 1] += diff;
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: w,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((hd, i) => cell(hd, w[i], { bold: true, fill: 'D5E8F0' })),
      }),
      ...rows.map((r) => new TableRow({ children: r.map((c, i) => cell(c, w[i])) })),
    ],
  });
}

function spacer() { return new Paragraph({ children: [], spacing: { after: 120 } }); }

function labelPara(label, text) {
  return new Paragraph({
    children: [t(label + ': ', { bold: true }), t(text)],
    spacing: { before: 120, after: 80 },
  });
}

function renderUC(uc, brCounter) {
  const out = [];
  out.push(h3(uc.id + ': ' + uc.name));
  out.push(labelPara('Mục tiêu', uc.objective));
  out.push(labelPara('Tác nhân', uc.actors));
  out.push(labelPara('Kích hoạt', uc.trigger));
  out.push(labelPara('Điều kiện tiên quyết', uc.pre));
  out.push(labelPara('Kết quả sau thực hiện', uc.post));

  out.push(p('Luồng sự kiện chính:', { run: { bold: true } }));
  out.push(table(
    ['Bước', 'Mô tả'],
    uc.steps.map((s, i) => ['(' + (i + 1) + ')', s]),
    [900, CW - 900],
  ));
  out.push(spacer());

  const excKw = ['msg', 'từ chối', 'lỗi', 'chặn', 'không hợp lệ', 'hết hạn', 'không đúng',
    'không khớp', 'vô hiệu', 'trống', 'trùng', 'sai', 'không thể', 'không tìm thấy', 'không được'];
  const isExc = function (d) {
    const dl = d.toLowerCase();
    return excKw.some(function (k) { return dl.indexOf(k) !== -1; });
  };
  let exFlows = [];
  if (uc.altFlows) { exFlows = uc.altFlows.slice(); }
  if (uc.brs && uc.brs.length) {
    uc.brs.forEach(function (b) { if (isExc(b.desc)) { exFlows.push({ step: b.step, desc: b.desc }); } });
  }
  out.push(p('Luồng thay thế / ngoại lệ:', { run: { bold: true } }));
  if (exFlows.length) {
    out.push(table(
      ['Tại bước', 'Điều kiện & cách xử lý'],
      exFlows.map(function (e) { return ['(' + e.step + ')', e.desc]; }),
      [1100, CW - 1100],
    ));
  } else {
    out.push(p('Không có luồng ngoại lệ đặc biệt; lỗi đầu vào xử lý theo quy tắc chung.'));
  }
  out.push(spacer());

  out.push(p('Quy tắc nghiệp vụ:', { run: { bold: true } }));
  if (uc.brs && uc.brs.length) {
    out.push(table(
      ['Bước', 'Mã BR', 'Mô tả'],
      uc.brs.map(function (b) { return ['(' + b.step + ')', 'BR' + (brCounter.n++), b.desc]; }),
      [900, 1100, CW - 2000],
    ));
  } else {
    out.push(p('Không có.'));
  }
  out.push(spacer());
  return out;
}

module.exports = { CW, t, p, pRuns, h1, h2, h3, h4, bullet, cell, table, spacer, labelPara, renderUC };
