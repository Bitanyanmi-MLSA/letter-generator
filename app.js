let names = [];
const $ = id => document.getElementById(id);

function show(text, cls) { const m = $('msg'); m.textContent = text; m.className = cls; }

// Minimal CSV parser supporting quoted fields
function parseCsv(text) {
  const rows = []; let row = [], cur = '', q = false;
  text = text.replace(/^\uFEFF/, '');
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"' && text[i + 1] === '"') { cur += '"'; i++; }
      else if (c === '"') q = false;
      else cur += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(cur); cur = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cur); rows.push(row); row = []; cur = '';
    } else cur += c;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows.filter(r => r.some(x => x.trim() !== ''));
}

function extractNames(text) {
  const rows = parseCsv(text);
  if (!rows.length) throw new Error('The CSV file is empty.');
  if (rows.some(r => r.length > 1 && r.slice(1).some(x => x.trim()))) {
    throw new Error('The CSV must contain names only: a single column.');
  }
  let list = rows.map(r => r[0].trim());
  if (/^(name|names|full ?name)$/i.test(list[0])) list.shift();
  const bad = list.find(n => !/^[\p{L}\p{M}][\p{L}\p{M}\s.'’-]*$/u.test(n));
  if (bad !== undefined) throw new Error(`"${bad}" does not look like a name. Only names are accepted.`);
  if (!list.length) throw new Error('No names found.');
  return list;
}

async function loadFile(f) {
  names = []; $('go').disabled = true; $('preview').innerHTML = '';
  if (!f) return;
  if (!/\.csv$/i.test(f.name)) return show('Only .csv files are accepted.', 'err');
  try {
    names = extractNames(await f.text());
    $('fileinfo').textContent = `${f.name}: ${names.length} name(s)`;
    names.slice(0, 20).forEach(n => { const li = document.createElement('li'); li.textContent = n; $('preview').append(li); });
    if (names.length > 20) { const li = document.createElement('li'); li.textContent = '…'; $('preview').append(li); }
    $('go').disabled = false; show('', '');
  } catch (e) { $('fileinfo').textContent = ''; show(e.message, 'err'); }
}

function buildDoc(name) {
  const { Document, Paragraph, TextRun, AlignmentType } = docx;
  const sp = { after: 200 };
  const paras = [
    new Paragraph({ children: [new TextRun({ text: new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' }) })], spacing: sp }),
    new Paragraph({ children: [new TextRun({ text: `Dear ${name},` })], spacing: sp }),
    new Paragraph({ children: [new TextRun({ text: $('subject').value, bold: true, underline: {} })], alignment: AlignmentType.CENTER, spacing: sp }),
    ...$('body').value.split(/\n\s*\n/).filter(p => p.trim()).map(p =>
      new Paragraph({ children: [new TextRun(p.trim())], alignment: AlignmentType.JUSTIFIED, spacing: sp })),
    new Paragraph({ children: [new TextRun($('sign').value)], spacing: { before: 300 } }),
  ];
  return new Document({ sections: [{ children: paras }] });
}

async function generate() {
  $('go').disabled = true; show('Generating…', 'ok');
  try {
    const zip = new JSZip(), folder = zip.folder('letters'), used = {};
    for (const n of names) {
      let base = n.replace(/[^\p{L}\p{N}]+/gu, '_').replace(/^_|_$/g, '') || 'letter';
      used[base] = (used[base] || 0) + 1;
      const file = `Letter_${base}${used[base] > 1 ? '_' + used[base] : ''}.docx`;
      folder.file(file, await docx.Packer.toBlob(buildDoc(n)));
    }
    const blob = await zip.generateAsync({ type: 'blob' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = 'letters.zip'; a.click();
    URL.revokeObjectURL(a.href);
    show(`Done: ${names.length} letter(s) saved in letters.zip`, 'ok');
  } catch (e) { show('Failed: ' + e.message, 'err'); }
  $('go').disabled = false;
}

const drop = $('drop');
drop.onclick = () => $('file').click();
$('file').onchange = e => loadFile(e.target.files[0]);
drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
drop.ondragleave = () => drop.classList.remove('over');
drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); loadFile(e.dataTransfer.files[0]); };
$('go').onclick = generate;
