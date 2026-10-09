// Общий генератор документов Akcent Role Play в советском стиле.
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, Header, Footer, PageNumber, LevelFormat, PageBreak,
} = require('docx');

const RED = 'A51C1C';
const GREY = '5A5A5A';
const FONT = 'Times New Roman';
const LETTERS = 'абвгдежзиклмнопрстуфхцчшщэюя';

const para = (children, opts = {}) => new Paragraph({
  spacing: { after: 90, line: 288 },
  alignment: opts.align || AlignmentType.JUSTIFIED,
  indent: opts.indent || { firstLine: 567 },
  keepNext: opts.keepNext,
  children: typeof children === 'string' ? [new TextRun(children)] : children,
});

const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: border, bottom: border, left: border, right: border };

const table = ({ w, head, rows, size = 21, boldFirst = true }) => new Table({
  width: { size: w.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  columnWidths: w,
  rows: [
    new TableRow({
      tableHeader: true,
      children: head.map((t, i) => new TableCell({
        width: { size: w[i], type: WidthType.DXA }, borders,
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: RED },
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, color: 'FFFFFF', size })] })],
      })),
    }),
    ...rows.map((r, ri) => new TableRow({
      cantSplit: true,
      children: r.map((t, i) => new TableCell({
        width: { size: w[i], type: WidthType.DXA }, borders,
        shading: ri % 2 ? { type: ShadingType.CLEAR, color: 'auto', fill: 'F5EFEA' } : undefined,
        margins: { top: 45, bottom: 45, left: 100, right: 100 },
        children: [new Paragraph({ children: [new TextRun({ text: String(t), bold: i === 0 && boldFirst, size })] })],
      })),
    })),
  ],
});

// Блоки содержимого раздела:
//   'текст'                       — абзац
//   ['1.1', 'текст']              — пункт с номером
//   { h2: 'подзаголовок' }
//   { list: [...] }               — маркированный список (звёздочки)
//   { abc: [...] }                — список «а) б) в)»
//   { table: { w, head, rows } }
//   { note: 'текст' }             — примечание
//   { quote: [...] }              — выделенный текст (клятва, образец)
const renderBlock = (b) => {
  if (typeof b === 'string') return [para(b)];
  if (Array.isArray(b)) {
    const [num, text] = b;
    return [para([new TextRun({ text: `${num}. `, bold: true, color: RED }), new TextRun(text)], { indent: { firstLine: 567 } })];
  }
  if (b.h2) {
    return [new Paragraph({ keepNext: true, spacing: { before: 200, after: 90 },
      children: [new TextRun({ text: b.h2, bold: true, size: 25 })] })];
  }
  if (b.list) {
    return b.list.map((t) => new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      spacing: { after: 40, line: 276 }, alignment: AlignmentType.JUSTIFIED,
      children: [new TextRun(t)],
    }));
  }
  if (b.abc) {
    return b.abc.map((t, i) => new Paragraph({
      spacing: { after: 40, line: 276 }, alignment: AlignmentType.JUSTIFIED,
      indent: { left: 1000, hanging: 400 },
      children: [new TextRun({ text: `${LETTERS[i]}) `, bold: true }), new TextRun(t)],
    }));
  }
  if (b.table) return [table(b.table), new Paragraph({ spacing: { after: 100 }, children: [] })];
  if (b.note) {
    return [new Paragraph({ spacing: { after: 110 }, indent: { firstLine: 567 }, alignment: AlignmentType.JUSTIFIED,
      children: [new TextRun({ text: 'Примечание. ', bold: true, size: 21 }), new TextRun({ text: b.note, size: 21 })] })];
  }
  if (b.quote) {
    return b.quote.map((t, i) => new Paragraph({
      spacing: { after: 60, line: 288 }, alignment: AlignmentType.JUSTIFIED,
      indent: { left: 850, right: 850 },
      border: i === 0 ? { top: { style: BorderStyle.SINGLE, size: 6, color: RED, space: 6 } } : undefined,
      children: [new TextRun({ text: t, italics: true })],
    })).concat([new Paragraph({ spacing: { after: 120 }, indent: { left: 850, right: 850 },
      border: { top: { style: BorderStyle.SINGLE, size: 6, color: RED, space: 6 } }, children: [] })]);
  }
  throw new Error('unknown block ' + JSON.stringify(b).slice(0, 80));
};

const build = async (spec) => {
  const front = [];
  (spec.approved || ['УТВЕРЖДЕНО', 'Политбюро проекта', 'Akcent Role Play', '«___» ____________ 2026 г.']).forEach((t, i) =>
    front.push(new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 40 },
      children: [new TextRun({ text: t, bold: i === 0, size: 22 })] })));
  front.push(new Paragraph({ spacing: { before: 1200 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: '★', size: 72, color: RED })] }));
  front.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 120 },
    children: [new TextRun({ text: spec.title, bold: true, size: spec.titleSize || 52, color: RED, characterSpacing: 40 })] }));
  spec.subtitle.forEach((t) => front.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
    children: [new TextRun({ text: t, size: 30 })] })));
  (spec.footnote || []).forEach((t, i) => front.push(new Paragraph({ alignment: AlignmentType.CENTER,
    spacing: { before: i === 0 ? 600 : 0, after: 60 }, children: [new TextRun({ text: t, italics: true, color: GREY, size: 22 })] })));
  front.push(new Paragraph({ children: [new PageBreak()] }));
  front.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: 'Содержание', bold: true, size: 32, color: RED })] }));
  spec.sections.forEach((s) => front.push(new Paragraph({ spacing: { after: 70 },
    children: [new TextRun({ text: '★  ', color: RED, size: 20 }), new TextRun(s.title)] })));

  const body = [];
  spec.sections.forEach((s, si) => {
    body.push(new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: si === 0 || s.newPage !== false,
      spacing: { before: s.newPage === false ? 300 : 0, after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RED, space: 4 } },
      children: [new TextRun(s.title)],
    }));
    s.blocks.forEach((b) => body.push(...renderBlock(b)));
  });
  if (spec.signature) {
    body.push(new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: spec.signature, italics: true, color: GREY })] }));
  }

  const doc = new Document({
    creator: 'Akcent Role Play',
    title: spec.docTitle || spec.title,
    styles: {
      default: { document: { run: { font: FONT, size: 24 } } },
      paragraphStyles: [{ id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: RED },
        paragraph: { spacing: { before: 0, after: 200 }, outlineLevel: 0 } }],
    },
    numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '★',
      alignment: AlignmentType.LEFT, style: { run: { color: RED }, paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
    sections: [{
      properties: { page: { margin: { top: 1134, bottom: 1134, left: 1418, right: 850 } }, titlePage: true },
      headers: {
        default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: spec.header, size: 18, color: GREY })] })] }),
        first: new Header({ children: [new Paragraph({ children: [] })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: GREY })] })] }),
        first: new Footer({ children: [new Paragraph({ children: [] })] }),
      },
      children: [...front, ...body],
    }],
  });
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(spec.file, buf);
  console.log('written', spec.file);
};

module.exports = { build };
