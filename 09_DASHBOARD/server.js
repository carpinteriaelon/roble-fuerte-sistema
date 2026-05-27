const express = require('express');
const xlsx    = require('xlsx');
const path    = require('path');
const multer  = require('multer');
const fs      = require('fs');

const app  = express();
const PORT = 3000;
const BASE = path.join(__dirname, '..');

app.use(express.static(__dirname));
app.use(express.json());

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'Dashboard_Executivo.html'));
});

function lerSheet(ws) {
  const raw = xlsx.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Procura cabeçalho nas primeiras 12 linhas; tenta limiar 4, depois 3
  let headerIdx = -1;
  for (const limiar of [4, 3]) {
    for (let i = 0; i < Math.min(raw.length, 12); i++) {
      const strings = raw[i].filter(c => typeof c === 'string' && c.trim() !== '');
      if (strings.length >= limiar) { headerIdx = i; break; }
    }
    if (headerIdx !== -1) break;
  }

  // Extrai resumo das linhas antes do cabeçalho (pares "Rótulo: valor_numérico")
  const resumo = {};
  for (let i = 0; i < (headerIdx === -1 ? raw.length : headerIdx); i++) {
    const row = raw[i];
    for (let j = 0; j < row.length - 1; j++) {
      const label = String(row[j] || '').trim().replace(/:$/, '');
      const valor = row[j + 1];
      if (label && typeof valor === 'number') resumo[label] = valor;
    }
  }

  if (headerIdx === -1) return { resumo, dados: [] };

  // Monta cabeçalhos: nomeia col vazia após IVA como TOTAL, ignora duplicados
  const seen = new Set();
  const cabecalhos = raw[headerIdx].map((h, i, arr) => {
    let s = String(h || '').trim();
    if (!s) {
      const prev = String(arr[i - 1] || '').trim();
      if (/iva|tax/i.test(prev) && !seen.has('TOTAL')) s = 'TOTAL';
      else return '';
    }
    if (seen.has(s)) {
      const s2 = s + '_2';
      if (!seen.has(s2)) { seen.add(s2); return s2; }
      return '';
    }
    seen.add(s);
    return s;
  });

  const dados = [];
  for (let i = headerIdx + 1; i < raw.length; i++) {
    const row = raw[i];
    if (row.every(c => c === '' || c == null)) continue;
    const obj = {};
    cabecalhos.forEach((h, j) => { if (h) obj[h] = row[j] ?? ''; });
    if (Object.values(obj).some(v => v !== '')) dados.push(obj);
  }

  return { resumo, dados };
}

const _cache = {};

function lerExcel(caminhoArquivo) {
  try {
    const mtime = fs.statSync(caminhoArquivo).mtimeMs;
    if (_cache[caminhoArquivo]?.mtime === mtime) return _cache[caminhoArquivo].dados;
    const wb = xlsx.readFile(caminhoArquivo);
    const resultado = {};
    wb.SheetNames.forEach(nome => { resultado[nome] = lerSheet(wb.Sheets[nome]); });
    _cache[caminhoArquivo] = { mtime, dados: resultado };
    return resultado;
  } catch (e) {
    return { erro: e.message };
  }
}

const MESES_RE = {
  JANEIRO:   /jan|ene/i,
  FEVEREIRO: /fev|feb/i,
  MARCO:     /mar[çcz]?o/i,
  ABRIL:     /abr/i,
  MAIO:      /\bmai/i,
  JUNHO:     /jun/i,
  JULHO:     /jul/i,
  AGOSTO:    /ago/i,
  SETEMBRO:  /set/i,
  OUTUBRO:   /out|oct/i,
  NOVEMBRO:  /nov/i,
  DEZEMBRO:  /dez|dic/i
};

function detectarMes(filename) {
  for (const [mes, re] of Object.entries(MESES_RE)) {
    if (re.test(filename)) return mes;
  }
  return null;
}

app.get('/api/fechamento', (req, res) => {
  const arquivos = fs.readdirSync(pastaUploads)
    .filter(f => /fechamento/i.test(f) && /\.(xlsx|xls)$/i.test(f));
  if (!arquivos.length) return res.json({});
  const resultado = {};
  arquivos.forEach(f => {
    const mes = detectarMes(f);
    if (mes) resultado[mes] = lerExcel(path.join(pastaUploads, f));
  });
  res.json(resultado);
});

app.get('/api/extrato', (req, res) => {
  const arquivos = fs.readdirSync(pastaUploads)
    .filter(f => /extrato/i.test(f) && /\.(xlsx|xls)$/i.test(f));
  const resultado = {};
  arquivos.forEach(f => {
    const mes = detectarMes(f);
    if (mes) resultado[mes] = lerExcel(path.join(pastaUploads, f));
  });
  res.json(resultado);
});

app.get('/api/financeiro', (req, res) => {
  res.json(lerExcel(path.join(BASE, '01_FINANCEIRO', 'Controle_Financeiro_Master.xlsx')));
});

app.get('/api/obras', (req, res) => {
  res.json(lerExcel(path.join(BASE, '02_OBRAS', 'CRM_Projetos.xlsx')));
});

app.put('/api/obra-campo', express.json(), (req, res) => {
  const { id, campo, valor } = req.body || {};
  if (!id || !campo) return res.status(400).json({ erro: 'id e campo obrigatórios' });
  try {
    const filePath = path.join(BASE, '02_OBRAS', 'CRM_Projetos.xlsx');
    const wb = xlsx.readFile(filePath);
    const ws = wb.Sheets['Pipeline'];
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const headerRow = rows.findIndex(r => r.includes('ID'));
    if (headerRow === -1) return res.status(500).json({ erro: 'Cabeçalho não encontrado' });
    const headers = rows[headerRow];
    const colIdx  = headers.indexOf(campo);
    if (colIdx === -1) return res.status(400).json({ erro: `Campo "${campo}" não existe` });
    const dataRowIdx = rows.findIndex((r, i) => i > headerRow && r[0] === id);
    if (dataRowIdx === -1) return res.status(404).json({ erro: `ID "${id}" não encontrado` });
    const cellAddr = xlsx.utils.encode_cell({ r: dataRowIdx, c: colIdx });
    ws[cellAddr] = { t: 's', v: valor };
    xlsx.writeFile(wb, filePath);
    delete _cache[filePath];
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ erro: e.message });
  }
});

app.get('/api/producao', (req, res) => {
  res.json(lerExcel(path.join(BASE, '05_PRODUCAO', 'Producao_Funcionarios.xlsx')));
});

app.get('/api/cadastros', (req, res) => {
  res.json(lerExcel(path.join(BASE, '00_Cadastros_Mestres.xlsx')));
});

// ── Upload de arquivos Excel ────────────────────────────
const pastaUploads = path.join(BASE, '08_DOCUMENTOS', 'Uploads');
if (!fs.existsSync(pastaUploads)) fs.mkdirSync(pastaUploads, { recursive: true });

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, pastaUploads),
  filename:    (req, file, cb) => {
    const ts   = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const nome = Buffer.from(file.originalname, 'latin1').toString('utf8');
    cb(null, ts + '_' + nome);
  }
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ok = /\.(xlsx|xls|csv|pdf|docx|doc|png|jpg|jpeg)$/i.test(file.originalname);
    cb(null, ok);
  },
  limits: { fileSize: 20 * 1024 * 1024 }
});

app.post('/api/upload', upload.single('arquivo'), (req, res) => {
  if (!req.file) return res.status(400).json({ erro: 'Arquivo não recebido ou tipo não permitido.' });

  const resposta = {
    ok: true,
    nome: req.file.originalname,
    salvo: req.file.filename,
    pasta: pastaUploads,
    tamanho: (req.file.size / 1024).toFixed(1) + ' KB'
  };

  // Se for Excel, retorna prévia dos dados
  if (/\.(xlsx|xls)$/i.test(req.file.originalname)) {
    resposta.previa = lerExcel(req.file.path);
  }

  res.json(resposta);
});

app.get('/api/ler-upload/:nome', (req, res) => {
  const arquivo = path.join(pastaUploads, req.params.nome);
  if (!require('fs').existsSync(arquivo)) return res.status(404).json({ erro: 'Não encontrado' });
  res.json(lerExcel(arquivo));
});

app.get('/api/uploads', (req, res) => {
  try {
    const arquivos = fs.readdirSync(pastaUploads).map(f => {
      const stat = fs.statSync(path.join(pastaUploads, f));
      return { nome: f, tamanho: (stat.size / 1024).toFixed(1) + ' KB', data: stat.mtime.toISOString().slice(0, 10) };
    }).sort((a, b) => b.data.localeCompare(a.data));
    res.json(arquivos);
  } catch { res.json([]); }
});

// ── Documentos de Obras (Presupuesto / Projeto) ─────────
const pastaObras = path.join(BASE, '08_DOCUMENTOS', 'Obras');
if (!fs.existsSync(pastaObras)) fs.mkdirSync(pastaObras, { recursive: true });

const storageObra = multer.diskStorage({
  destination: (req, file, cb) => cb(null, pastaObras),
  filename: (req, file, cb) => {
    const id   = String(req.body.id   || 'obra').replace(/[^a-zA-Z0-9_-]/g, '');
    const tipo = String(req.body.tipo || 'doc').replace(/[^a-zA-Z0-9_-]/g, '');
    const ext  = path.extname(file.originalname);
    cb(null, `${id}_${tipo}_${Date.now()}${ext}`);
  }
});
const uploadObra = multer({
  storage: storageObra,
  fileFilter: (req, file, cb) => cb(null, /\.(pdf|png|jpg|jpeg)$/i.test(file.originalname)),
  limits: { fileSize: 20 * 1024 * 1024 }
});

app.post('/api/obra-doc', uploadObra.single('arquivo'), (req, res) => {
  if (!req.file) return res.status(400).json({ erro: 'Arquivo não recebido ou tipo não permitido.' });
  res.json({ ok: true, nome: req.file.filename });
});

app.get('/api/obra-docs', (req, res) => {
  try {
    const arquivos = fs.readdirSync(pastaObras).map(f => {
      const stat = fs.statSync(path.join(pastaObras, f));
      return { nome: f, tamanho: (stat.size / 1024).toFixed(1) + ' KB' };
    });
    res.json(arquivos);
  } catch { res.json([]); }
});

app.get('/api/obra-doc/:nome', (req, res) => {
  const arquivo = path.join(pastaObras, req.params.nome);
  if (!fs.existsSync(arquivo)) return res.status(404).json({ erro: 'Não encontrado' });
  res.sendFile(arquivo);
});

app.delete('/api/obra-doc/:nome', (req, res) => {
  const arquivo = path.join(pastaObras, req.params.nome);
  if (!fs.existsSync(arquivo)) return res.status(404).json({ erro: 'Não encontrado' });
  fs.unlinkSync(arquivo);
  res.json({ ok: true });
});

// ── Checklist de Obras ─────────────────────────────────
const CHECKLIST_FILE = path.join(BASE, '08_DOCUMENTOS', 'obras_checklist.json');

function lerChecklist() {
  try {
    if (!fs.existsSync(CHECKLIST_FILE)) return {};
    return JSON.parse(fs.readFileSync(CHECKLIST_FILE, 'utf8'));
  } catch { return {}; }
}

app.get('/api/obras-checklist', (req, res) => res.json(lerChecklist()));

app.put('/api/obras-checklist/:id/notas', express.json(), (req, res) => {
  const { notas } = req.body || {};
  if (notas === undefined) return res.status(400).json({ erro: 'notas obrigatório' });
  const data = lerChecklist();
  if (!data[req.params.id]) data[req.params.id] = {};
  data[req.params.id]['_notas'] = notas;
  fs.writeFileSync(CHECKLIST_FILE, JSON.stringify(data, null, 2), 'utf8');
  res.json({ ok: true });
});

app.put('/api/obras-checklist/:id/acoes', express.json(), (req, res) => {
  const { acoes } = req.body || {};
  if (!Array.isArray(acoes)) return res.status(400).json({ erro: 'acoes deve ser array' });
  const data = lerChecklist();
  if (!data[req.params.id]) data[req.params.id] = {};
  data[req.params.id]['_acoes'] = acoes;
  fs.writeFileSync(CHECKLIST_FILE, JSON.stringify(data, null, 2), 'utf8');
  res.json({ ok: true });
});

app.put('/api/obras-checklist/:id/estrutura', express.json(), (req, res) => {
  const { estrutura } = req.body || {};
  if (!estrutura) return res.status(400).json({ erro: 'estrutura obrigatória' });
  const data = lerChecklist();
  if (!data[req.params.id]) data[req.params.id] = {};
  data[req.params.id]['_estrutura'] = estrutura;
  fs.writeFileSync(CHECKLIST_FILE, JSON.stringify(data, null, 2), 'utf8');
  res.json({ ok: true });
});

app.put('/api/obras-checklist/:id', express.json(), (req, res) => {
  const { item, checked } = req.body || {};
  if (!item) return res.status(400).json({ erro: 'item obrigatório' });
  const data = lerChecklist();
  if (!data[req.params.id]) data[req.params.id] = {};
  data[req.params.id][item] = !!checked;
  fs.writeFileSync(CHECKLIST_FILE, JSON.stringify(data, null, 2), 'utf8');
  res.json({ ok: true });
});

app.put('/api/obras-checklist/:id/financeiro', express.json(), (req, res) => {
  const { campo, valor } = req.body || {};
  if (!campo) return res.status(400).json({ erro: 'campo obrigatório' });
  const data = lerChecklist();
  if (!data[req.params.id]) data[req.params.id] = {};
  data[req.params.id][campo] = valor;
  fs.writeFileSync(CHECKLIST_FILE, JSON.stringify(data, null, 2), 'utf8');
  res.json({ ok: true });
});

// ── Facturas de Giro ────────────────────────────────────
const GIRO_FILE = path.join(BASE, '08_DOCUMENTOS', 'facturas_giro.json');

function lerGiro() {
  try {
    if (!fs.existsSync(GIRO_FILE)) return { facturas: [] };
    return JSON.parse(fs.readFileSync(GIRO_FILE, 'utf8'));
  } catch { return { facturas: [] }; }
}

function gravarGiro(data) {
  fs.writeFileSync(GIRO_FILE, JSON.stringify(data, null, 2), 'utf8');
}

app.get('/api/giro', (req, res) => res.json(lerGiro()));

app.post('/api/giro', (req, res) => {
  const data = lerGiro();
  const f = { ...req.body, id: Date.now().toString() };
  if (!f.dataVencimento && f.dataFactura && f.prazo) {
    const d = new Date(f.dataFactura);
    d.setDate(d.getDate() + parseInt(f.prazo));
    f.dataVencimento = d.toISOString().slice(0, 10);
  }
  if (!f.estado) f.estado = 'pendente';
  data.facturas.push(f);
  gravarGiro(data);
  res.json({ ok: true, factura: f });
});

app.put('/api/giro/:id', (req, res) => {
  const data = lerGiro();
  const idx = data.facturas.findIndex(f => f.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });
  data.facturas[idx] = { ...data.facturas[idx], ...req.body };
  gravarGiro(data);
  res.json({ ok: true, factura: data.facturas[idx] });
});

app.delete('/api/giro/:id', (req, res) => {
  const data = lerGiro();
  data.facturas = data.facturas.filter(f => f.id !== req.params.id);
  gravarGiro(data);
  res.json({ ok: true });
});

// ── Colunas Kanban ─────────────────────────────────────
const KANBAN_COLUNAS_FILE = path.join(BASE, '08_DOCUMENTOS', 'kanban_colunas.json');
const DEFAULT_COLUNAS = [
  { status: 'Aprovado', label: 'OBRA',    color: '#F59E0B' },
  { status: 'Produção', label: 'A FAZER', color: '#3B82F6' },
  { status: 'Lacado',   label: 'FAZENDO', color: '#8B5CF6' },
  { status: 'Montagem', label: 'FEITO',   color: '#10B981' },
];
function lerKanbanColunas() {
  try {
    if (!fs.existsSync(KANBAN_COLUNAS_FILE)) return DEFAULT_COLUNAS;
    return JSON.parse(fs.readFileSync(KANBAN_COLUNAS_FILE, 'utf8'));
  } catch { return DEFAULT_COLUNAS; }
}
app.get('/api/kanban-colunas', (req, res) => res.json(lerKanbanColunas()));
app.put('/api/kanban-colunas', express.json(), (req, res) => {
  if (!Array.isArray(req.body)) return res.status(400).json({ erro: 'array esperado' });
  fs.writeFileSync(KANBAN_COLUNAS_FILE, JSON.stringify(req.body, null, 2), 'utf8');
  res.json({ ok: true });
});

// ── Obras CRUD ─────────────────────────────────────────
app.delete('/api/obra/:id', (req, res) => {
  try {
    const filePath = path.join(BASE, '02_OBRAS', 'CRM_Projetos.xlsx');
    const wb = xlsx.readFile(filePath);
    const ws = wb.Sheets['Pipeline'];
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const headerRow = rows.findIndex(r => r.includes('ID'));
    if (headerRow === -1) return res.status(500).json({ erro: 'Cabeçalho não encontrado' });
    const dataRowIdx = rows.findIndex((r, i) => i > headerRow && String(r[0]) === req.params.id);
    if (dataRowIdx === -1) return res.status(404).json({ erro: `ID não encontrado` });
    rows.splice(dataRowIdx, 1);
    const newWs = xlsx.utils.aoa_to_sheet(rows);
    if (ws['!cols']) newWs['!cols'] = ws['!cols'];
    wb.Sheets['Pipeline'] = newWs;
    xlsx.writeFile(wb, filePath);
    delete _cache[filePath];
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ erro: e.message });
  }
});

app.post('/api/obra', express.json(), (req, res) => {
  try {
    const filePath = path.join(BASE, '02_OBRAS', 'CRM_Projetos.xlsx');
    const wb = xlsx.readFile(filePath);
    const ws = wb.Sheets['Pipeline'];
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const headerRow = rows.findIndex(r => r.includes('ID'));
    if (headerRow === -1) return res.status(500).json({ erro: 'Cabeçalho não encontrado' });
    const headers = rows[headerRow];
    const existingIds = rows.slice(headerRow + 1)
      .map(r => String(r[0] || ''))
      .filter(id => /^P-\d+$/.test(id))
      .map(id => parseInt(id.slice(2)));
    const nextNum = existingIds.length ? Math.max(...existingIds) + 1 : 1;
    const newId = `P-${String(nextNum).padStart(3, '0')}`;
    const newRow = headers.map((h, i) => {
      if (i === 0) return newId;
      return req.body[String(h)] !== undefined ? req.body[String(h)] : '';
    });
    rows.push(newRow);
    const newWs = xlsx.utils.aoa_to_sheet(rows);
    if (ws['!cols']) newWs['!cols'] = ws['!cols'];
    wb.Sheets['Pipeline'] = newWs;
    xlsx.writeFile(wb, filePath);
    delete _cache[filePath];
    res.json({ ok: true, id: newId });
  } catch (e) {
    res.status(500).json({ erro: e.message });
  }
});

// ── Tarefas Kanban (móveis por cliente) ─────────────────
const TAREFAS_FILE = path.join(BASE, '08_DOCUMENTOS', 'tarefas_kanban.json');

function lerTarefas() {
  try {
    if (!fs.existsSync(TAREFAS_FILE)) return { tarefas: [] };
    return JSON.parse(fs.readFileSync(TAREFAS_FILE, 'utf8'));
  } catch { return { tarefas: [] }; }
}

function gravarTarefas(data) {
  fs.writeFileSync(TAREFAS_FILE, JSON.stringify(data, null, 2), 'utf8');
}

app.get('/api/tarefas', (req, res) => res.json(lerTarefas()));

app.post('/api/tarefas', express.json(), (req, res) => {
  const { obraId, nome, status } = req.body || {};
  if (!obraId || !status) return res.status(400).json({ erro: 'obraId e status obrigatórios' });
  const data = lerTarefas();
  const t = { id: Date.now().toString(), obraId, nome: nome || '', status, acoes: [], notas: '' };
  data.tarefas.push(t);
  gravarTarefas(data);
  res.json({ ok: true, tarefa: t });
});

app.put('/api/tarefas/:id', express.json(), (req, res) => {
  const data = lerTarefas();
  const idx = data.tarefas.findIndex(t => t.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });
  data.tarefas[idx] = { ...data.tarefas[idx], ...req.body };
  gravarTarefas(data);
  res.json({ ok: true, tarefa: data.tarefas[idx] });
});

app.delete('/api/tarefas/:id', (req, res) => {
  const data = lerTarefas();
  data.tarefas = data.tarefas.filter(t => t.id !== req.params.id);
  gravarTarefas(data);
  res.json({ ok: true });
});

// ── Pagamentos de Funcionários ─────────────────────────
const PAGAMENTOS_FILE = path.join(BASE, '08_DOCUMENTOS', 'pagamentos_func.json');

function lerPagamentos() {
  try {
    if (!fs.existsSync(PAGAMENTOS_FILE)) return [];
    return JSON.parse(fs.readFileSync(PAGAMENTOS_FILE, 'utf8'));
  } catch { return []; }
}
function gravarPagamentos(data) {
  fs.writeFileSync(PAGAMENTOS_FILE, JSON.stringify(data, null, 2), 'utf8');
}

app.get('/api/pagamentos', (req, res) => res.json(lerPagamentos()));

app.post('/api/pagamentos', express.json(), (req, res) => {
  const { funcionario, mes, tipoSalario, dias, valorDiaria, valorFixo } = req.body || {};
  if (!funcionario || !mes || !tipoSalario) return res.status(400).json({ erro: 'campos obrigatórios' });
  const records = lerPagamentos();
  const record = { id: Date.now().toString(), funcionario, mes, tipoSalario,
    dias: dias || null, valorDiaria: valorDiaria || null, valorFixo: valorFixo || null, adelantos: [] };
  records.push(record);
  gravarPagamentos(records);
  res.json({ ok: true, record });
});

app.put('/api/pagamentos/:id', express.json(), (req, res) => {
  const records = lerPagamentos();
  const idx = records.findIndex(r => r.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });
  records[idx] = { ...records[idx], ...req.body };
  gravarPagamentos(records);
  res.json({ ok: true });
});

app.post('/api/pagamentos/duplicar-mes', express.json(), (req, res) => {
  const { mesFonte, mesDestino } = req.body || {};
  if (!mesFonte || !mesDestino) return res.status(400).json({ erro: 'mesFonte e mesDestino obrigatórios' });
  const records = lerPagamentos();
  const fonte = records.filter(r => r.mes === mesFonte);
  if (!fonte.length) return res.status(404).json({ erro: 'Nenhum registo no mês de origem' });
  const jaExistem = new Set(records.filter(r => r.mes === mesDestino).map(r => r.funcionario));
  const novos = fonte
    .filter(r => !jaExistem.has(r.funcionario))
    .map((r, i) => ({
      ...r,
      id: (Date.now() + i).toString(),
      mes: mesDestino,
      adel1Data: '', adel1Valor: null,
      adel2Data: '', adel2Valor: null,
      adel3Data: '', adel3Valor: null,
      adelantos: []
    }));
  if (!novos.length) return res.status(409).json({ erro: 'Todos os funcionários já existem nesse mês' });
  records.push(...novos);
  gravarPagamentos(records);
  res.json({ ok: true, criados: novos.length });
});

app.delete('/api/pagamentos/:id', (req, res) => {
  gravarPagamentos(lerPagamentos().filter(r => r.id !== req.params.id));
  res.json({ ok: true });
});

app.post('/api/pagamentos/:id/adelanto', express.json(), (req, res) => {
  const { data, valor } = req.body || {};
  if (!data || !valor) return res.status(400).json({ erro: 'data e valor obrigatórios' });
  const records = lerPagamentos();
  const idx = records.findIndex(r => r.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });
  records[idx].adelantos.push({ id: Date.now().toString(), data, valor: parseFloat(valor) });
  gravarPagamentos(records);
  res.json({ ok: true });
});

app.delete('/api/pagamentos/:id/adelanto/:adelId', (req, res) => {
  const records = lerPagamentos();
  const idx = records.findIndex(r => r.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });
  records[idx].adelantos = records[idx].adelantos.filter(a => a.id !== req.params.adelId);
  gravarPagamentos(records);
  res.json({ ok: true });
});

// ── Instagram Posts ────────────────────────────────────
const INSTAGRAM_FILE = path.join(BASE, '08_DOCUMENTOS', 'instagram_posts.json');

function lerInstagram() {
  try {
    if (!fs.existsSync(INSTAGRAM_FILE)) return [];
    return JSON.parse(fs.readFileSync(INSTAGRAM_FILE, 'utf8'));
  } catch { return []; }
}

function gravarInstagram(data) {
  fs.writeFileSync(INSTAGRAM_FILE, JSON.stringify(data, null, 2), 'utf8');
}

app.get('/api/instagram', (req, res) => res.json(lerInstagram()));

app.post('/api/instagram', express.json(), (req, res) => {
  const { data, descricao, tipo } = req.body || {};
  if (!data || !descricao || !tipo) return res.status(400).json({ erro: 'data, descricao e tipo obrigatórios' });
  const posts = lerInstagram();
  const post = { id: Date.now().toString(), data, descricao, tipo };
  posts.push(post);
  gravarInstagram(posts);
  res.json({ ok: true, post });
});

app.delete('/api/instagram/:id', (req, res) => {
  const posts = lerInstagram();
  gravarInstagram(posts.filter(p => p.id !== req.params.id));
  res.json({ ok: true });
});

app.listen(PORT, () => {
  console.log('\n  Dashboard Roble Fuerte rodando!');
  console.log(`  Abra no browser: http://localhost:${PORT}\n`);
});
