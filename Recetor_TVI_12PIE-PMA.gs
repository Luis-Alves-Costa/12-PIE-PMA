/**
 * Aula assíncrona P6 · Da taxa média à taxa instantânea · Matemática 12.º PIE-PMA (Classroom 3ºPMA/PIE)
 * Prof. Luís Alves da Costa · Escola Secundária Marquês de Pombal · 2026/2027
 *
 * prepararTVI() faz tudo de uma vez:
 *   1. cria o Google Forms próprio do glossário (as respostas ficam nesta folha);
 *   2. publica no Classroom, à hora marcada: o texto de apoio (material), o glossário (trabalho até às 17h00)
 *      e o mini-teste (trabalho de 20 pontos até às 18h30);
 *   3. liga o relógio (de 5 em 5 minutos) que fecha as entregas e envia os relatórios.
 * Às 17h00: fecha o Forms do glossário e envia ao professor o relatório dos percursos.
 * Às 18h30: fecha o mini-teste, lança e devolve as notas no Classroom, envia a cada aluno a resolução
 *           (classificação, erros assinalados, caminhos de melhoria) e ao professor o relatório dos erros.
 * Serviço avançado necessário: «Google Classroom API».
 */
const COURSE_ID = '886040969042';              // 3ºPMA/PIE 2026/2027
const PROF_EMAIL = 'luis.costa@esmp.pt';
const PAGES = 'https://luis-alves-costa.github.io/12-PIE-PMA/';
const F_TEXTO = 'P6_TVI_Texto_Apoio.html', F_GLOS = 'P6_TVI_Glossario.html', F_MT = 'P6_TVI_MiniTeste.html';
const FECHO_GLOS = new Date('2026-11-24T17:00:00Z'), FECHO_MT = new Date('2026-11-24T18:30:00Z');   // novembro: Lisboa = UTC
const CODIGO_MT = 'MT-P6-TVI-01';
const TITULO_MT = 'Mini-teste · Da taxa média à taxa instantânea (P6)';
const F_RES = 'Mini-teste', CAB = ['Recebido em', 'Nome', 'Turma', 'Email', 'Valores (0-20)', 'Menção', 'Dados (JSON)', 'Classroom', 'Resolução enviada em'];
const P = PropertiesService.getScriptProperties();

/* ================= preparação (correr uma vez) ================= */
function prepararTVI() {
  const publicar = new Date();                         // publica já; muda aqui para agendar, ex.: new Date('2026-11-24T12:30:00Z')
  folha_(F_RES, CAB);
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('verificar').timeBased().everyMinutes(5).create();
  const form = criarForm_();
  const topico = topicoId_();
  if (!P.getProperty('CW_TEXTO')) P.setProperty('CW_TEXTO', Classroom.Courses.CourseWorkMaterials.create(Object.assign({
      title: 'P6 · Aula assíncrona · Da taxa média à taxa instantânea (texto de apoio)',
      description: 'Lê o texto de apoio no telemóvel e mexe nos gráficos. Depois faz o glossário (até às 17h00) e o mini-teste (até às 18h30).',
      materials: [{ link: { url: PAGES + F_TEXTO + '?v=1' } }], topicId: topico || undefined }, agenda_(publicar)), COURSE_ID).id);
  if (!P.getProperty('CW_GLOS')) P.setProperty('CW_GLOS', Classroom.Courses.CourseWork.create(Object.assign({
      title: 'Glossário · Taxa instantânea (até às 17h00)', workType: 'ASSIGNMENT',
      description: 'Percorre os termos do glossário e verifica se os percebeste. No fim, o teu percurso é enviado automaticamente. ' +
                   'Depois carrega em «Entregar». As entregas fecham às 17h00.',
      materials: [{ link: { url: PAGES + F_GLOS + '?v=1' } }], topicId: topico || undefined,
      dueDate: dia_(FECHO_GLOS), dueTime: hora_(FECHO_GLOS) }, agenda_(publicar)), COURSE_ID).id);
  if (!P.getProperty('CW_MT')) P.setProperty('CW_MT', Classroom.Courses.CourseWork.create(Object.assign({
      title: TITULO_MT + ' (até às 18h30)', workType: 'ASSIGNMENT', maxPoints: 20,
      description: 'Faz o mini-teste no telemóvel; o resultado é enviado automaticamente. Depois carrega em «Entregar». ' +
                   'As entregas fecham às 18h30. A seguir recebes no email da escola a resolução, com a classificação e caminhos para melhorar.',
      materials: [{ link: { url: PAGES + F_MT + '?v=1' } }], topicId: topico || undefined,
      dueDate: dia_(FECHO_MT), dueTime: hora_(FECHO_MT) }, agenda_(publicar)), COURSE_ID).id);
  return JSON.stringify(form);
}
function agenda_(d) { return d.getTime() > Date.now() + 60000 ? { state: 'DRAFT', scheduledTime: d.toISOString() } : { state: 'PUBLISHED' }; }
function dia_(d) { return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() }; }
function hora_(d) { return { hours: d.getUTCHours(), minutes: d.getUTCMinutes() }; }

/* 1.º passo: cria só o Google Forms do glossário (para ligar o glossário antes de publicar) */
function criarFormulario() { return JSON.stringify(criarForm_()); }

/* Google Forms próprio do glossário: os campos que o glossário envia */
function criarForm_() {
  if (P.getProperty('FORM_CFG')) return JSON.parse(P.getProperty('FORM_CFG'));
  const f = FormApp.create('Glossário P6 · Taxa instantânea · 12.º PIE-PMA (percursos)');
  f.setDescription('Respostas enviadas automaticamente pelo glossário. Fecha às 17h00 de 24/11/2026.');
  try { f.setRequireLogin(false); } catch (e) {}
  f.setCollectEmail(false);
  f.setDestination(FormApp.DestinationType.SPREADSHEET, SpreadsheetApp.getActive().getId());
  const nomes = { nome: 'Nome', turma: 'Turma', lingua: 'Língua', termos: 'Termos percorridos', certas: 'Certas',
                  tempo: 'Tempo', fim: 'Fim', zonas: 'Zonas de dificuldade', ordem: 'Ordem do percurso', dados: 'Dados' };
  const itens = {};
  Object.keys(nomes).forEach(k => { itens[k] = (k === 'ordem' || k === 'dados' || k === 'zonas') ? f.addParagraphTextItem().setTitle(nomes[k]) : f.addTextItem().setTitle(nomes[k]); });
  let r = f.createResponse();
  Object.keys(itens).forEach(k => { r = r.withItemResponse(itens[k].createResponse('x' + k)); });
  const url = r.toPrefilledUrl(), campos = {};
  Object.keys(itens).forEach(k => { const m = url.match(new RegExp('entry\\.(\\d+)=x' + k + '(&|$)')); if (m) campos[k] = 'entry.' + m[1]; });
  const cfg = { url: f.getPublishedUrl().replace(/\/viewform.*$/, '/formResponse'), campos: campos, edit: f.getEditUrl(), id: f.getId() };
  P.setProperty('FORM_CFG', JSON.stringify(cfg));
  folha_('Config', ['Chave', 'Valor']).appendRow(['form', JSON.stringify(cfg)]);
  return cfg;
}

/* ================= receção do mini-teste ================= */
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (Date.now() > FECHO_MT.getTime() + 2 * 60000) return ContentService.createTextOutput('fechado');
    const d = JSON.parse(e.postData.contents);
    if (d.codigo !== CODIGO_MT) return ContentService.createTextOutput('codigo');
    const email = String(d.email || '').trim().toLowerCase(), sh = folha_(F_RES, CAB);
    const rows = sh.getLastRow() > 1 ? sh.getRange(2, 4, sh.getLastRow() - 1, 1).getValues().flat() : [];
    if (rows.map(x => String(x).toLowerCase()).indexOf(email) >= 0) return ContentService.createTextOutput('repetido');
    sh.appendRow([new Date(), d.nome, d.turma, email, Number(d.valores), d.mencao, JSON.stringify(d), '', '']);
    if (!P.getProperty('RESOLUCAO') && d.resolucao) P.setProperty('RESOLUCAO', JSON.stringify(d.resolucao));
    return ContentService.createTextOutput('ok');
  } finally { lock.releaseLock(); }
}
function doGet() { return ContentService.createTextOutput('Recetor P6 · 12.º PIE-PMA ativo.'); }

/* ================= relógio: fechos e relatórios ================= */
function verificar() {
  const agora = Date.now();
  if (agora >= FECHO_GLOS.getTime() && !P.getProperty('GLOS_FEITO')) { fecharGlossario_(); P.setProperty('GLOS_FEITO', new Date().toISOString()); }
  if (agora >= FECHO_MT.getTime() + 60000 && !P.getProperty('MT_FEITO')) { fecharMiniTeste_(); P.setProperty('MT_FEITO', new Date().toISOString()); }
  if (P.getProperty('GLOS_FEITO') && P.getProperty('MT_FEITO'))
    ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'verificar') ScriptApp.deleteTrigger(t); });
}

function fecharGlossario_() {
  const cfg = JSON.parse(P.getProperty('FORM_CFG') || '{}');
  if (!cfg.id) return;
  const f = FormApp.openById(cfg.id);
  f.setAcceptingResponses(false).setCustomClosedFormMessage('As entregas do glossário fecharam às 17h00.');
  const it = f.getItems(), resp = f.getResponses();
  const linhas = resp.map(r => { const o = {}; r.getItemResponses().forEach(x => o[x.getItem().getTitle()] = x.getResponse()); o._t = r.getTimestamp(); return o; });
  const faltam = semEntrega_(linhas.map(l => l['Nome']));
  const tab = linhas.map(l => '<tr><td>' + esc_(l['Nome']) + '</td><td>' + esc_(l['Língua']) + '</td><td>' + esc_(l['Termos percorridos']) + '/12</td><td>' +
      esc_(l['Certas']) + '</td><td>' + esc_(l['Tempo']) + '</td><td>' + esc_(l['Zonas de dificuldade']) + '</td><td style="font-size:12px">' + esc_(l['Ordem do percurso']) + '</td></tr>').join('');
  // termos mais falhados
  const falhas = {};
  linhas.forEach(l => String(l['Ordem do percurso'] || '').split(' | ').forEach(s => { if (/✗/.test(s)) { const t = s.replace(/^\d+\.\s*/, '').replace(' ✗', ''); falhas[t] = (falhas[t] || 0) + 1; } }));
  const top = Object.keys(falhas).sort((a, b) => falhas[b] - falhas[a]).slice(0, 6).map(t => '<li>' + esc_(t) + ' — ' + falhas[t] + ' aluno(s)</li>').join('') || '<li>nenhum</li>';
  MailApp.sendEmail({ to: PROF_EMAIL, subject: 'Glossário P6 (12.º PIE-PMA) · relatório dos percursos às 17h00 · ' + linhas.length + ' entrega(s)',
    htmlBody: cab_('Glossário P6 · Da taxa média à taxa instantânea', 'Relatório dos percursos (fecho às 17h00)') +
      '<p><b>Entregas:</b> ' + linhas.length + '</p><h3>Termos em que mais se erraram as verificações</h3><ul>' + top + '</ul>' +
      '<h3>Percursos</h3><table border="1" cellpadding="5" style="border-collapse:collapse;font-size:13px"><tr style="background:#dde8f5"><th>Aluno</th><th>Língua</th><th>Termos</th><th>Certas</th><th>Tempo</th><th>Zonas de dificuldade</th><th>Ordem (✗ = errou a verificação)</th></tr>' + tab + '</table>' +
      (faltam.length ? '<h3>Sem entrega</h3><p>' + faltam.map(esc_).join(', ') + '</p>' : '') +
      '<p style="color:#667">Respostas completas no Forms: ' + esc_(cfg.edit) + '</p>' });
}

function fecharMiniTeste_() {
  const sh = folha_(F_RES, CAB);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, CAB.length).getValues() : [];
  const R = JSON.parse(P.getProperty('RESOLUCAO') || '[]'), cw = P.getProperty('CW_MT');
  const stats = R.map(() => ({ primeira: 0, certa: 0, nunca: 0, ajuda: 0, erradas: {} }));
  const alunos = [];
  rows.forEach((r, k) => {
    const d = JSON.parse(r[6]); alunos.push(d);
    (d.respostas || []).forEach((x, i) => { if (!x || !stats[i]) return; const s = stats[i];
      if (x.ok) { s.certa++; if (x.tent === 1) s.primeira++; } else s.nunca++;
      if (x.ajuda) s.ajuda++;
      (x.esc || []).forEach(o => { if (R[i] && o !== R[i].c) s.erradas[o] = (s.erradas[o] || 0) + 1; }); });
    let c = '';
    try { c = classroomNota_(cw, String(r[3]), Number(r[4])); } catch (e) { c = 'Erro: ' + e.message; }
    try { MailApp.sendEmail({ to: String(r[3]), replyTo: PROF_EMAIL, name: 'Prof. Luís Alves da Costa', subject: 'Resolução · ' + TITULO_MT, htmlBody: emailAluno_(d, R) }); }
    catch (e) { c += ' · email: ' + e.message; }
    sh.getRange(k + 2, 8, 1, 2).setValues([[c, new Date()]]);
  });
  const n = alunos.length || 1;
  const porQ = R.map((q, i) => { const s = stats[i], err = Object.keys(s.erradas).sort((a, b) => s.erradas[b] - s.erradas[a]);
    return '<tr><td>P' + (i + 1) + '</td><td>' + Math.round(100 * s.primeira / n) + ' %</td><td>' + Math.round(100 * s.certa / n) + ' %</td><td>' + s.nunca + '</td><td>' + s.ajuda + '</td><td>' +
      (err.length ? err.slice(0, 2).map(o => '«' + esc_(o) + '» (' + s.erradas[o] + ')').join('<br>') : '—') + '</td></tr>'; }).join('');
  const porA = alunos.map(d => '<tr><td>' + esc_(d.nome) + '</td><td><b>' + String(d.valores).replace('.', ',') + '</b></td><td>' +
      (d.respostas || []).map((x, i) => !x ? '·' : (x.ok ? (x.tent === 1 ? '✓' : '✓' + x.tent) : '✗') + (x.ajuda ? '💡' : '')).join(' ') + '</td><td style="font-size:12px">' +
      (d.respostas || []).map((x, i) => x && x.esc && x.esc.some(o => o !== R[i].c) ? 'P' + (i + 1) + ': ' + x.esc.filter(o => o !== R[i].c).map(esc_).join(' → ') : '').filter(String).join('<br>') +
      '</td><td style="font-size:12px">' + esc_(d.duvidas || '') + '</td></tr>').join('');
  const faltam = semEntrega_(alunos.map(d => d.nome), alunos.map(d => String(d.email).toLowerCase()));
  const media = alunos.length ? alunos.reduce((s, d) => s + Number(d.valores), 0) / alunos.length : 0;
  MailApp.sendEmail({ to: PROF_EMAIL, subject: 'Mini-teste P6 (12.º PIE-PMA) · relatório dos erros · ' + alunos.length + ' entrega(s) · média ' + media.toFixed(1).replace('.', ','),
    htmlBody: cab_(TITULO_MT, 'Relatório dos erros e dos percursos (fecho às 18h30)') +
      '<p><b>Entregas:</b> ' + alunos.length + ' · <b>Média:</b> ' + media.toFixed(1).replace('.', ',') + ' valores</p>' +
      '<h3>Por pergunta</h3><table border="1" cellpadding="5" style="border-collapse:collapse;font-size:13px"><tr style="background:#dde8f5"><th>Pergunta</th><th>Certa à 1.ª</th><th>Certa no fim</th><th>Não acertaram</th><th>Usaram ajuda</th><th>Erros mais escolhidos</th></tr>' + porQ + '</table>' +
      '<h3>Por aluno</h3><p style="font-size:12px">✓ = certa à 1.ª · ✓2/✓3 = à 2.ª/3.ª tentativa · ✗ = não acertou · 💡 = consultou a ajuda</p>' +
      '<table border="1" cellpadding="5" style="border-collapse:collapse;font-size:13px"><tr style="background:#dde8f5"><th>Aluno</th><th>Valores</th><th>P1…P8</th><th>Escolhas erradas (por ordem)</th><th>Dúvidas</th></tr>' + porA + '</table>' +
      (faltam.length ? '<h3>Sem entrega</h3><p>' + faltam.map(esc_).join(', ') + '</p>' : '') +
      '<p>Cada aluno recebeu a sua resolução por email e a nota foi lançada e devolvida no Classroom.</p>' });
}

function emailAluno_(d, R) {
  const nome = String(d.nome || '').split(' ')[0];
  const blocos = R.map((q, i) => { const x = (d.respostas || [])[i] || {}, erradas = (x.esc || []).filter(o => o !== q.c);
    const est = x.ok ? (x.tent === 1 ? '✅ Certa à primeira.' : '✅ Certa à ' + x.tent + '.ª tentativa.') : '🔸 Para rever.';
    const melhoria = x.ok && x.tent === 1 ? '' :
      '<p style="margin:4px 0;background:#fdf3e4;border-radius:8px;padding:6px 10px"><b>Caminho de melhoria:</b> ' + q.a + ' Revê o ' + esc_(q.s) + ' do <a href="' + PAGES + F_TEXTO + '">texto de apoio</a>.</p>';
    return '<div style="border-left:4px solid ' + (x.ok && x.tent === 1 ? '#a8d5b5' : '#f2c38b') + ';padding:6px 12px;margin:12px 0">' +
      '<p style="margin:0 0 4px"><b>' + (i + 1) + '.</b> ' + q.t + '</p><p style="margin:0 0 4px">' + est + '</p>' +
      (erradas.length ? '<p style="margin:0 0 4px;color:#a14d2a">As tuas escolhas erradas: ' + erradas.map(o => '«' + esc_(o) + '»').join(', ') + '</p>' : '') +
      '<p style="margin:0 0 4px">Resposta certa: <b>' + esc_(q.c) + '</b></p><p style="margin:0;color:#455">' + q.x + '</p>' + melhoria + '</div>'; }).join('');
  return cab_(TITULO_MT, 'A tua resolução') + '<p>Olá, ' + esc_(nome) + '!</p>' +
    '<div style="background:#e9f6ee;border-radius:12px;padding:12px 16px;text-align:center;margin:12px 0"><div style="font-size:30px;font-weight:700">' +
    String(d.valores).replace('.', ',') + '</div><div>valores · ' + esc_(d.mencao) + '</div></div>' +
    '<p>Abaixo tens cada pergunta: o que escolheste, a resposta certa e, onde for útil, um caminho para melhorar. A ajuda e as tentativas fazem parte da aprendizagem.</p>' +
    blocos + '<p>Se ficaste com dúvidas, escreve-me ou traz-as para a próxima aula. Bom trabalho!</p><p>Prof. Luís Alves da Costa</p></div>';
}
function cab_(t, s) {
  return '<div style="font-family:Arial,sans-serif;max-width:680px;color:#223"><div style="background:#dde8f5;border-radius:12px;padding:12px 16px;margin-bottom:14px">' +
    '<b>Matemática · Turma 12.º PIE-PMA (2026/2027)</b><br>Prof. Luís Alves da Costa · Módulo P6 – Taxa de Variação</div><h2 style="color:#2b5d8a;margin:0 0 4px">' + esc_(t) + '</h2><p style="color:#667;margin:0 0 12px">' + esc_(s) + '</p>';
}

/* ================= Classroom ================= */
function classroomNota_(cw, email, v) {
  if (!cw || !email) return 'sem trabalho';
  const sub = (Classroom.Courses.CourseWork.StudentSubmissions.list(COURSE_ID, cw, { userId: email }).studentSubmissions || [])[0];
  if (!sub) return 'aluno não encontrado (' + email + ')';
  Classroom.Courses.CourseWork.StudentSubmissions.patch({ draftGrade: v, assignedGrade: v }, COURSE_ID, cw, sub.id, { updateMask: 'draftGrade,assignedGrade' });
  try { Classroom.Courses.CourseWork.StudentSubmissions['return']({}, COURSE_ID, cw, sub.id); return 'nota ' + v + ' lançada e devolvida'; }
  catch (e) { return 'nota ' + v + ' lançada (devolução: ' + e.message + ')'; }
}
function semEntrega_(nomes, emails) {
  try {
    const al = Classroom.Courses.Students.list(COURSE_ID, { pageSize: 100 }).students || [];
    const nz = nomes.map(n => norm_(n));
    return al.filter(s => { const e = String(s.profile.emailAddress || '').toLowerCase(), n = norm_(s.profile.name.fullName);
      if (emails && emails.indexOf(e) >= 0) return false;
      return !nz.some(x => x && (x === n || n.indexOf(x.split(' ')[0]) === 0 && x.split(' ').slice(-1)[0] === n.split(' ').slice(-1)[0])); }).map(s => s.profile.name.fullName);
  } catch (e) { return []; }
}
function norm_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim(); }
function topicoId_() { try { const t = (Classroom.Courses.Topics.list(COURSE_ID).topic || []).filter(x => /matem/i.test(x.name))[0]; return t ? t.topicId : null; } catch (e) { return null; } }
function folha_(nome, cab) {
  const ss = SpreadsheetApp.getActive(); let sh = ss.getSheetByName(nome);
  if (!sh) { sh = ss.insertSheet(nome); sh.appendRow(cab); sh.setFrozenRows(1); sh.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#dde8f5'); }
  return sh;
}
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]); }
