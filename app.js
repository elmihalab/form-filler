/*
 * مُعبِّئ الاستبيانات — الواجهة
 * شعبة الشؤون العلمية في حلب
 */
(function () {
  'use strict';
  const E = window.Engine;
  const S = window.Store;
  const $app = document.getElementById('app');

  const state = {
    view: 'home',     // home | editor | run
    tab: 'saved',     // saved | new | help
    profiles: [],
    draft: null,      // استبيان قيد الإعداد
    current: null,    // الاستبيان المفتوح للتوليد
    run: null,        // بيانات ملف الإكسل المفتوح
  };

  const DEFAULT_SETTINGS = {
    incompleteAction: 'guide',
    entity: 'مديرية أوقاف حلب',
    filePattern: '{التسلسل} - {الاسم}',
    emptyValue: '',
    markOn: '●',
    markOff: ' ',
    outputMode: 'zip',
    markDone: true,
  };

  const TYPE_OPTIONS = [
    ['text', 'نص'],
    ['date', 'تاريخ'],
    ['list-lines', 'قائمة (كل سطر عنصر)'],
    ['list-commas', 'قائمة (مفصولة بفواصل)'],
    ['list-numbers', 'قائمة (مرقّمة 1- 2- …)'],
    ['choice', 'اختيار من خيارات'],
  ];

  /* ================================================================== *
   *  أدوات مساعدة
   * ================================================================== */

  function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach((k) => {
        const v = attrs[k];
        if (v === null || v === undefined || v === false) return;
        if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
        else if (k === 'class') el.className = v;
        else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
        else if (['value', 'checked', 'disabled', 'selected', 'multiple', 'indeterminate'].includes(k)) el[k] = v;
        else if (v === true) el.setAttribute(k, '');
        else el.setAttribute(k, v);
      });
    }
    kids.flat(Infinity).forEach((c) => {
      if (c === null || c === undefined || c === false) return;
      el.appendChild(c instanceof Node ? c : document.createTextNode(String(c)));
    });
    return el;
  }

  function tag(t) { return h('span', { class: 'tag' }, '{' + t + '}'); }

  function fmtDate(ts) {
    if (!ts) return '';
    const d = new Date(ts);
    return String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear();
  }

  function stamp() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  function safeName(s) {
    return String(s || '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 90) || 'استبيان';
  }

  function newId() { return 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

  function toast(msg, kind) {
    let box = document.querySelector('.toasts');
    if (!box) { box = h('div', { class: 'toasts' }); document.body.appendChild(box); }
    const t = h('div', { class: 'toast' + (kind === 'err' ? ' err' : ''), role: 'status' }, msg);
    box.appendChild(t);
    setTimeout(() => t.remove(), kind === 'err' ? 7000 : 3500);
  }

  /** نافذة سؤال: تعيد قيمة الزر المضغوط */
  function modal({ title, body, buttons }) {
    return new Promise((resolve) => {
      const back = h('div', { class: 'modal-back' });
      const close = (v) => { back.remove(); resolve(v); };
      const box = h('div', { class: 'modal', role: 'dialog', 'aria-modal': 'true' },
        h('h3', null, title),
        typeof body === 'string' ? h('p', null, body) : body,
        h('div', { class: 'row' }, buttons.map((b) =>
          h('button', { class: 'btn' + (b.kind ? ' ' + b.kind : ''), onclick: () => close(b.value) }, b.label))));
      back.addEventListener('click', (e) => { if (e.target === back) close(null); });
      back.appendChild(box);
      document.body.appendChild(back);
      const first = box.querySelector('.btn'); if (first) first.focus();
    });
  }

  function pickFile(accept) {
    return new Promise((resolve) => {
      const input = document.getElementById('file-any');
      input.value = '';
      input.accept = accept || '';
      input.onchange = () => resolve(input.files[0] || null);
      input.click();
    });
  }

  function download(data, name, mime) {
    const blob = data instanceof Blob ? data : new Blob([data], { type: mime || 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = h('a', { href: url, download: name });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  async function copyText(text) {
    try { await navigator.clipboard.writeText(text); }
    catch (e) {
      const ta = h('textarea', { style: { position: 'fixed', opacity: '0' } }, text);
      document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove();
    }
    toast('نُسخ ' + text + ' — الصقه في ملف الوورد');
  }

  const tick = () => new Promise((r) => setTimeout(r, 0));

  let builtinCache = null;
  function builtin() {
    if (!builtinCache) {
      builtinCache = {
        auto: E.b64ToBuf(window.BUILTIN_TEMPLATES.auto),
        guide: E.b64ToBuf(window.BUILTIN_TEMPLATES.guide),
      };
    }
    return builtinCache;
  }

  function typeValue(f) {
    if (f.type === 'list') return 'list-' + (f.split || 'lines');
    return f.type || 'text';
  }

  function setType(f, v) {
    if (v.startsWith('list-')) { f.type = 'list'; f.split = v.slice(5); }
    else { f.type = v; delete f.split; }
  }

  function fieldTag(p, key) {
    const f = (p.fields || {})[key];
    return (f && f.tag) || '';
  }

  function sampleValue(table, key) {
    if (!table) return '';
    const r = table.rows.find((x) => String(x.values[key] || '').trim() !== '');
    return r ? String(r.values[key]).replace(/\n/g, ' ⏎ ') : '';
  }

  function distinctValues(table, key) {
    if (!table) return [];
    const set = new Set();
    table.rows.forEach((r) => { const v = String(r.values[key] || '').trim(); if (v && v.length < 80) set.add(v); });
    return Array.from(set).slice(0, 25);
  }

  function guessNameField(headers) {
    const prefer = headers.find((x) => /اسم/.test(x.col) && /(مقدم|الطالب|المتقدم|الكامل|الثلاثي|الإمام|الخطيب)/.test(x.col));
    const any = headers.find((x) => /^(الاسم|اسم)/.test(x.col.trim()) && !/(النشاط|المسجد|الجامع|الدورة)/.test(x.col));
    const anyName = headers.find((x) => /اسم/.test(x.col) && !/(النشاط|المسجد|الجامع|الدورة)/.test(x.col));
    return (prefer || any || anyName || headers[0] || {}).key || null;
  }

  function suggestName(fileName) {
    const base = String(fileName || '').replace(/\.[^.]+$/, '').replace(/[_\-]+/g, ' ').trim();
    if (!base || /^[0-9a-f\s]{6,}$/i.test(base) || base.length < 3) return 'استبيان جديد';
    return base;
  }

  async function refreshProfiles() {
    state.profiles = await S.all();
    updateFooter();
  }

  async function updateFooter() {
    const el = document.getElementById('storage-note');
    const u = await S.usage();
    let txt = 'تُحفظ الاستبيانات في متصفح هذا الجهاز فقط، ولا يُرفع أي ملف إلى الإنترنت';
    if (u !== null) txt += ' — المساحة المستخدمة: ' + (u / 1048576).toFixed(1) + ' ميغابايت';
    el.textContent = txt;
  }

  function rowName(p, row) {
    if (p.nameField) {
      const v = String(row.values[p.nameField] || '').trim();
      if (v) return v.split('\n')[0];
    }
    const first = p.headers.map((x) => row.values[x.key]).find((v) => String(v || '').trim());
    return first ? String(first).split('\n')[0].slice(0, 50) : ('طلب ' + row.seq);
  }

  /* ================================================================== *
   *  الرسم العام
   * ================================================================== */

  function render() {
    $app.innerHTML = '';
    if (state.view === 'editor') $app.appendChild(viewEditor());
    else if (state.view === 'run') $app.appendChild(viewRun());
    else $app.appendChild(viewHome());
    window.scrollTo({ top: Math.min(window.scrollY, document.body.scrollHeight) });
  }

  function goHome(tab) {
    state.view = 'home';
    state.draft = null;
    state.current = null;
    state.run = null;
    if (tab) state.tab = tab;
    render();
    window.scrollTo({ top: 0 });
  }

  /* ================================================================== *
   *  الصفحة الرئيسية
   * ================================================================== */

  function viewHome() {
    const tabBtn = (id, label, count) => h('button', { class: state.tab === id ? 'on' : '', onclick: () => { state.tab = id; render(); } },
      label, count !== undefined ? h('span', { class: 'count' }, count) : null);
    const wrap = h('div', null,
      h('div', { class: 'tabs', role: 'tablist' },
        tabBtn('saved', 'استبياناتي المحفوظة', state.profiles.length),
        tabBtn('new', 'استبيان جديد'),
        tabBtn('help', 'طريقة الاستخدام')));
    if (state.tab === 'new') wrap.appendChild(viewNew());
    else if (state.tab === 'help') wrap.appendChild(viewHelp());
    else wrap.appendChild(viewSaved());
    return wrap;
  }

  function viewSaved() {
    const box = h('div');
    box.appendChild(h('div', { class: 'row', style: { marginBottom: '14px' } },
      h('button', { class: 'btn', onclick: () => { state.tab = 'new'; render(); } }, '＋ استبيان جديد'),
      h('button', { class: 'btn ghost', onclick: importFlow }, 'استيراد استبيان جاهز'),
      h('span', { class: 'spacer' }),
      h('span', { class: 'muted' }, 'اسحب ملف الإكسل إلى أي مكان في الصفحة، وسيعرف البرنامج إلى أي استبيان ينتمي')));
    if (!state.profiles.length) {
      box.appendChild(h('div', { class: 'card empty' },
        h('div', { class: 'big-ico' }, '🗂'),
        h('h2', null, 'لا توجد استبيانات محفوظة بعد'),
        h('p', { class: 'lead' }, 'ابدأ بإضافة أول استبيان: ارفع ملف الإكسل الذي أخرجه برنامج الوزارة، ثم اختر ما يظهر في ملف الوورد.'),
        h('div', { class: 'row', style: { justifyContent: 'center' } },
          h('button', { class: 'btn big', onclick: () => { state.tab = 'new'; render(); } }, 'إضافة استبيان جديد'),
          h('button', { class: 'btn ghost big', onclick: importFlow }, 'استيراد استبيان جاهز'))));
      return box;
    }
    const grid = h('div', { class: 'grid' });
    state.profiles.forEach((p) => {
      const done = Object.values(p.processed || {});
      const nMain = done.filter((x) => x.k === 'main').length;
      const nGuide = done.filter((x) => x.k === 'guide').length;
      grid.appendChild(h('div', { class: 'survey' },
        h('h3', null, p.name),
        h('div', { class: 'kick' }, p.mode === 'auto' ? 'القالب: تلقائي (جدول بالمعلومات)' : 'القالب: ' + ((p.template && p.template.name) || '—')),
        h('div', { class: 'stats' },
          h('span', { class: 'chip green' }, 'طلبات مولَّدة: ' + nMain),
          (p.manualRows || []).length ? h('span', { class: 'chip manual' }, 'طلبات يدوية: ' + p.manualRows.length) : null,
          (p.rules || []).length ? h('span', { class: 'chip gold' }, 'خطابات توجيه: ' + nGuide) : null,
          h('span', { class: 'chip gray' }, p.headers.length + ' سؤالًا')),
        h('div', { class: 'muted' }, p.lastExcel ? ('آخر ملف إكسل: ' + p.lastExcel.name + ' — ' + fmtDate(p.lastExcel.date)) : (p.source === 'template' ? 'يُعبّأ يدويًا (لا يوجد ملف إكسل)' : 'لم يُرفع ملف إكسل بعد')),
        h('div', { class: 'row', style: { marginTop: 'auto' } },
          h('button', { class: 'btn', onclick: () => openRun(p.id) }, 'فتح وتوليد الملفات'),
          h('button', { class: 'btn ghost small', onclick: () => openEditor(p.id) }, 'الإعدادات'),
          h('button', { class: 'btn ghost small', onclick: () => exportFlow(p) }, 'تصدير'),
          h('button', { class: 'btn danger small', onclick: () => deleteFlow(p) }, 'حذف'))));
    });
    box.appendChild(grid);
    return box;
  }

  function dropZone(text, sub, accept, onFile, small) {
    const z = h('div', { class: 'drop' + (small ? ' small' : ''), tabindex: '0', role: 'button' },
      h('div', { class: 'big-ico' }, '⬆'),
      h('strong', null, text),
      sub ? h('div', { class: 'muted' }, sub) : null);
    const go = async () => { const f = await pickFile(accept); if (f) onFile(f); };
    z.addEventListener('click', go);
    z.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } });
    z.addEventListener('dragover', (e) => { e.preventDefault(); z.classList.add('over'); });
    z.addEventListener('dragleave', () => z.classList.remove('over'));
    z.addEventListener('drop', (e) => {
      e.preventDefault(); e.stopPropagation(); z.classList.remove('over'); document.body.classList.remove('dragging');
      const f = e.dataTransfer.files[0]; if (f) onFile(f);
    });
    return z;
  }

  function viewNew() {
    return h('div', null,
      h('div', { class: 'card' },
        h('h2', null, 'إضافة استبيان جديد'),
        h('p', { class: 'lead' }, 'ارفع ملف الإكسل الذي أخرجه برنامج الوزارة لهذا الاستبيان. إن كان الملف لاستبيان محفوظ عندك من قبل، فسينبّهك البرنامج ويفتح لك الطلبات الجديدة فقط.'),
        dropZone('اضغط هنا لاختيار ملف الإكسل، أو اسحبه وأفلته هنا', 'الصيغ المقبولة: ‎.xlsx‎ أو ‎.csv‎', '.xlsx,.xlsm,.csv', (f) => routeExcel(f))),
      h('div', { class: 'card' },
        h('h2', null, 'أو ابدأ من قالب وورد (للتعبئة اليدوية)'),
        h('p', { class: 'lead' }, 'للاستمارات التي لا يوجد لها ملف إكسل، مثل أوامر التكليف: ضع الوسوم في ملف الوورد مكان الخطوط المنقّطة، مثل ', tag('اسم الشيخ'), ' و', tag('المسجد'), '، ثم ارفعه هنا. سيصنع البرنامج من الوسوم نموذجًا على الشاشة تكتب فيه، ثم ينزّل لك ملف الوورد معبّأً.'),
        dropZone('اضغط هنا لاختيار قالب الوورد، أو اسحبه وأفلته هنا', 'صيغة ‎.docx‎', '.docx', (f) => startDraftFromTemplate(f), true)),
      h('div', { class: 'card' },
        h('h2', null, 'أو استورد استبيانًا جاهزًا'),
        h('p', { class: 'lead' }, 'إذا أرسلت لك شعبة أخرى ملف استبيان مُصدَّرًا (ينتهي بـ ‎.json‎) فاستورده هنا، وستجد فيه القالب والإعدادات جاهزة.'),
        h('button', { class: 'btn ghost', onclick: importFlow }, 'استيراد استبيان جاهز')));
  }

  function viewHelp() {
    const ex = (t, d) => h('tr', null, h('td', null, h('span', { class: 'tag' }, t)), h('td', null, d));
    return h('div', { class: 'card help' },
      h('h2', null, 'طريقة الاستخدام'),
      h('p', { class: 'lead' }, 'البرنامج يأخذ ملف الإكسل الذي يخرجه برنامج الوزارة، ويملأ منه ملف وورد لكل طلب. ولا يغادر أي ملف جهازك.'),

      h('h3', null, 'أول مرة مع استبيان جديد'),
      h('ol', null,
        h('li', null, 'من تبويب «استبيان جديد» ارفع ملف الإكسل.'),
        h('li', null, 'في جدول الحقول: أشّر على المعلومات التي تريدها، وأعطِ كل واحدة اسمًا قصيرًا (الوسم)، واختر نوعها.'),
        h('li', null, 'اختر القالب: «قالب وورد خاص بي» إن كان عندك ملف الاستبيان الرسمي، أو «قالب تلقائي» ليصنع البرنامج مستندًا مرتبًا بجدول.'),
        h('li', null, 'إن اخترت قالبك الخاص: افتح ملف الوورد، وضع مكان كل معلومة وسمها بين قوسين، مثل ', tag('اسم مقدم الطلب'), '، ثم ارفعه.'),
        h('li', null, 'إن أردت: أضف شروط «الطلب المكتمل» ليولّد البرنامج خطاب توجيه للطلبات الناقصة.'),
        h('li', null, 'احفظ الاستبيان. سيُحفظ القالب وملف الإكسل في متصفحك.')),

      h('h3', null, 'المرات التالية (الملف نفسه بعد إضافة طلبات جديدة)'),
      h('ol', null,
        h('li', null, 'أخرج ملف الإكسل الأحدث من برنامج الوزارة.'),
        h('li', null, 'اسحبه إلى الصفحة. سيتعرّف البرنامج على الاستبيان ويعرض الطلبات الجديدة فقط.'),
        h('li', null, 'اختر طريقة الإخراج ثم اضغط «توليد الملفات».')),

      h('h3', null, 'التعبئة اليدوية (دون ملف إكسل)'),
      h('ol', null,
        h('li', null, 'افتح ملف الاستمارة في الوورد، وضع مكان كل خط منقّط وسمًا بين قوسين، مثل ', tag('اسم الشيخ'), ' و', tag('المسجد'), '.'),
        h('li', null, 'للاختيار بين كلمتين (مثل إمام / خطيب) اكتب ', tag('الصفة=إمام'), ' قبل «إمام» و', tag('الصفة=خطيب'), ' قبل «خطيب»، فيصير في النموذج حقل اختيار.'),
        h('li', null, 'من تبويب «استبيان جديد» اختر «ابدأ من قالب وورد» وارفع الملف. سيصنع البرنامج من الوسوم نموذجًا على الشاشة.'),
        h('li', null, 'احفظ الاستبيان، ثم اكتب في النموذج واضغط «تنزيل ملف الوورد الآن»، أو أضف عدة طلبات إلى القائمة وولّدها معًا.'),
        h('li', null, 'وفي أي استبيان مبني على إكسل تجد أيضًا زر «إضافة طلب يدويًا» لطلب جاء على الورق أو بالهاتف.')),

      h('h3', null, 'كتابة الوسوم داخل ملف الوورد'),
      h('div', { class: 'table-wrap' }, h('table', { class: 't' },
        h('tr', null, h('th', null, 'ما تكتبه'), h('th', null, 'ما يظهر في الملف')),
        ex('{اسم مقدم الطلب}', 'الجواب كما هو.'),
        ex('{الأهداف 1}  {الأهداف 2}  {الأهداف 3}', 'العنصر الأول والثاني والثالث من حقل نوعه «قائمة». الخانة الأخيرة تجمع ما زاد عن عدد الخانات.'),
        ex('{الأهداف}', 'كل عناصر القائمة، كل عنصر في سطر.'),
        ex('{#الأهداف}{.}{/الأهداف}', 'يكرر الفقرة لكل عنصر في القائمة (ضع كل وسم من الثلاثة في فقرة مستقلة، و{.} في فقرة مرقّمة).'),
        ex('{#ملاحظات}نص{/ملاحظات}', 'يظهر «نص» فقط إذا كان للسؤال جواب.'),
        ex('{^ملاحظات}نص{/ملاحظات}', 'يظهر «نص» فقط إذا كان الجواب فارغًا.'),
        ex('{نوع النشاط=درس علمي}', 'تظهر علامة ● إذا كان الجواب «درس علمي»، وإلا تبقى فارغة.'),
        ex('{اليوم}', 'تاريخ اليوم بالتقويم الميلادي، مثل 12/11/2026.'),
        ex('{اليوم الهجري}', 'تاريخ اليوم بالتقويم الهجري (أم القرى)، مثل 02/06/1448.'),
        ex('{التسلسل}', 'رقم الطلب في ملف الإكسل.'),
        ex('{الاسم}', 'اسم صاحب الطلب (الحقل المختار في الإعدادات).'),
        ex('{الاستبيان}', 'اسم الاستبيان.'),
        ex('{الجهة}', 'اسم الجهة المكتوب في الإعدادات، مثل: مديرية أوقاف حلب.'),
        ex('{أسباب التوجيه}', 'ما ينقص الطلب (في خطاب التوجيه).'))),
      h('ul', null,
        h('li', null, 'اكتب القوسين { } من لوحة المفاتيح الإنجليزية، واكتب الوسم دفعة واحدة.'),
        h('li', null, 'لا يهم في المطابقة: الهمزات، والتاء المربوطة والهاء، والتشكيل، والنقطتان في آخر السؤال.'),
        h('li', null, 'بعد رفع القالب يعرض لك البرنامج كل وسم: ✔ معروف، أو ✖ غير معروف مع اقتراح التصحيح.')),

      h('h3', null, 'مشاركة استبيان مع شعبة أو دائرة أخرى'),
      h('p', null, 'من بطاقة الاستبيان اضغط «تصدير»، فيُنزَّل ملف ينتهي بـ ‎.json‎ فيه القالب والإعدادات دون إجابات الناس. أرسله لهم، فيستوردونه من «استيراد استبيان جاهز».'),

      h('h3', null, 'أين تُحفظ البيانات؟'),
      h('p', null, 'في متصفح هذا الجهاز فقط (قاعدة بيانات المتصفح). إن مُسحت بيانات المتصفح أو فتحت البرنامج من جهاز آخر فلن تجد استبياناتك هناك، لذلك احتفظ بنسخة مُصدَّرة من كل استبيان.'));
  }

  /* ================================================================== *
   *  توجيه ملف الإكسل إلى استبيانه
   * ================================================================== */

  async function routeExcel(file, opts) {
    opts = opts || {};
    let buf, table;
    try {
      buf = await file.arrayBuffer();
      table = E.readTable(buf, file.name);
    } catch (e) { toast(e.message, 'err'); return; }
    if (!opts.forceNew && state.profiles.length) {
      const keys = table.headers.map((x) => x.key);
      let best = null, score = 0;
      state.profiles.forEach((p) => {
        const s = E.headerSimilarity(p.headers.map((x) => x.key), keys);
        if (s > score) { score = s; best = p; }
      });
      if (best && score >= 0.6) {
        const exact = score >= 0.98;
        const choice = await modal({
          title: exact ? 'هذا الملف لاستبيان محفوظ عندك' : 'هذا الملف يشبه استبيانًا محفوظًا عندك',
          body: h('div', null,
            h('p', null, 'الاستبيان: «', h('b', null, best.name), '»'),
            h('p', { class: 'muted' }, 'تطابق الأسئلة: ' + Math.round(score * 100) + '٪' + (exact ? '' : ' — ربما عدّلت الوزارة صياغة بعض الأسئلة.')),
            h('p', null, 'هل تريد فتحه لاستخراج الطلبات الجديدة؟')),
          buttons: [
            { label: 'نعم، افتحه', value: 'open' },
            { label: 'لا، هذا استبيان جديد', value: 'new', kind: 'ghost' },
            { label: 'إلغاء', value: null, kind: 'ghost' },
          ],
        });
        if (choice === 'open') return openRun(best.id, { file, buf, table });
        if (choice !== 'new') return;
      }
    }
    startDraft(file, buf, table);
  }

  /* ================================================================== *
   *  إعداد الاستبيان
   * ================================================================== */

  function startDraft(file, buf, table) {
    const headers = table.headers.map(({ col, key }) => ({ col, key }));
    state.draft = {
      id: newId(),
      name: suggestName(file.name),
      createdAt: Date.now(),
      headers,
      fields: E.makeFields(table.headers),
      nameField: guessNameField(headers),
      mode: 'auto',
      template: null,
      guideMode: 'builtin',
      guideTemplate: null,
      rules: [],
      settings: Object.assign({}, DEFAULT_SETTINGS),
      processed: {},
      manualRows: [],
      source: 'excel',
      lastExcel: { name: file.name, data: buf, date: Date.now() },
      _isNew: true,
      _table: table,
    };
    state.view = 'editor';
    render();
    window.scrollTo({ top: 0 });
    toast('قُرئ الملف: ' + table.rows.length + ' طلبًا و' + headers.length + ' سؤالًا');
  }

  async function startDraftFromTemplate(file) {
    if (!/\.docx$/i.test(file.name)) return toast('اختر ملف وورد بصيغة ‎.docx‎', 'err');
    const data = await file.arrayBuffer();
    const ins = E.inspectTemplate(data);
    if (!ins.ok) {
      await modal({ title: 'في القالب أخطاء تمنع استعماله', body: h('ul', null, ins.errors.map((x) => h('li', null, x))), buttons: [{ label: 'حسنًا', value: true }] });
      return;
    }
    const { headers, fields } = E.fieldsFromTemplate(ins.tags);
    if (!headers.length) {
      await modal({
        title: 'لم أجد وسومًا في هذا الملف',
        body: h('div', null, h('p', null, 'ضع مكان كل معلومة وسمًا بين قوسين، مثل ', tag('اسم الشيخ'), ' أو ', tag('المسجد'), '، ثم احفظ الملف وارفعه من جديد.'),
          h('p', { class: 'muted' }, 'اكتب القوسين { } من لوحة المفاتيح الإنجليزية.')),
        buttons: [{ label: 'حسنًا', value: true }],
      });
      return;
    }
    const nameKey = (headers.find((x) => /(اسم|الشيخ|الأخ|العامل|السيد)/.test(x.col)) || headers[0]).key;
    state.draft = {
      id: newId(),
      name: suggestName(file.name),
      createdAt: Date.now(),
      headers,
      fields,
      nameField: nameKey,
      mode: 'custom',
      template: { name: file.name, data },
      guideMode: 'builtin',
      guideTemplate: null,
      rules: [],
      settings: Object.assign({}, DEFAULT_SETTINGS),
      processed: {},
      manualRows: [],
      source: 'template',
      lastExcel: null,
      _isNew: true,
      _table: null,
      _inspect: ins,
    };
    state.view = 'editor';
    render();
    window.scrollTo({ top: 0 });
    toast('وجدت في القالب ' + headers.length + ' حقلًا');
  }

  async function openEditor(id) {
    const p = await S.get(id);
    if (!p) return toast('لم أجد هذا الاستبيان.', 'err');
    const d = structuredClone(p);
    d.settings = Object.assign({}, DEFAULT_SETTINGS, d.settings || {});
    d._isNew = false;
    d._table = null;
    if (d.lastExcel && d.lastExcel.data) { try { d._table = E.readTable(d.lastExcel.data, d.lastExcel.name); } catch (e) { /* تجاهل */ } }
    d._inspect = d.template ? E.inspectTemplate(d.template.data) : null;
    d._guideInspect = d.guideTemplate ? E.inspectTemplate(d.guideTemplate.data) : null;
    state.draft = d;
    state.view = 'editor';
    render();
    window.scrollTo({ top: 0 });
  }

  function card(step, title, ...kids) {
    return h('section', { class: 'card' }, h('h2', null, step ? h('span', { class: 'step' }, step) : null, title), ...kids);
  }

  function radioCards(name, value, options, onChange) {
    return h('div', { class: 'radio-cards' }, options.map(([v, title, sub]) =>
      h('label', { class: 'radio-card' + (value === v ? ' on' : '') },
        h('input', { type: 'radio', name, value: v, checked: value === v, onchange: () => onChange(v) }),
        h('div', null, h('b', null, title), sub ? h('span', null, sub) : null))));
  }

  function viewEditor() {
    const d = state.draft;
    const wrap = h('div');
    wrap.appendChild(h('div', { class: 'row', style: { marginBottom: '12px' } },
      h('h2', { style: { fontSize: '24px' } }, d._isNew ? 'إعداد استبيان جديد' : 'إعدادات الاستبيان'),
      h('span', { class: 'spacer' }),
      h('button', { class: 'btn ghost small', onclick: cancelEditor }, 'إلغاء والرجوع')));

    wrap.appendChild(edName(d));
    wrap.appendChild(edFields(d));
    wrap.appendChild(edTemplate(d));
    wrap.appendChild(edRules(d));
    wrap.appendChild(edOutput(d));

    wrap.appendChild(h('div', { class: 'genbar row' },
      h('button', { class: 'btn big', onclick: saveDraft }, d._isNew ? 'حفظ الاستبيان والانتقال إلى التوليد' : 'حفظ التعديلات'),
      h('button', { class: 'btn ghost', onclick: cancelEditor }, 'إلغاء'),
      h('span', { class: 'spacer' }),
      !d._isNew ? h('button', { class: 'btn danger small', onclick: resetProcessed }, 'تصفير سجل الطلبات المنجزة (' + Object.keys(d.processed || {}).length + ')') : null));
    return wrap;
  }

  async function cancelEditor() {
    const d = state.draft;
    if (d && d._isNew) {
      const ok = await modal({ title: 'إلغاء الاستبيان الجديد؟', body: 'لن يُحفظ شيء من هذا الإعداد.', buttons: [{ label: 'نعم، ألغِ', value: true, kind: 'danger' }, { label: 'متابعة الإعداد', value: false, kind: 'ghost' }] });
      if (!ok) return;
    }
    if (d && !d._isNew) openRun(d.id); else goHome('saved');
  }

  /* ---------- ① الاسم وملف الإكسل ---------- */
  function edName(d) {
    const t = d._table;
    return card('١', 'اسم الاستبيان وملف الإكسل',
      h('label', { class: 'field', style: { marginTop: '8px' } }, 'اسم الاستبيان (يظهر في القائمة وفي أسماء الملفات المجمّعة)',
        h('input', { type: 'text', value: d.name, oninput: (e) => { d.name = e.target.value; } })),
      h('div', { class: 'row', style: { marginTop: '12px' } },
        d.lastExcel ? h('span', { class: 'chip' }, 'الملف: ' + d.lastExcel.name) : h('span', { class: 'chip manual' }, 'يُعبّأ يدويًا — لا يوجد ملف إكسل'),
        t ? h('span', { class: 'chip green' }, t.rows.length + ' طلبًا') : null,
        h('span', { class: 'chip gray' }, d.headers.length + (d.lastExcel ? ' سؤالًا' : ' حقلًا')),
        h('span', { class: 'spacer' }),
        h('button', { class: 'btn ghost small', onclick: async () => { const f = await pickFile('.xlsx,.xlsm,.csv'); if (f) replaceDraftExcel(f); } }, d.lastExcel ? 'تغيير ملف الإكسل' : 'ربط ملف إكسل (اختياري)')),
      !d.lastExcel ? h('p', { class: 'muted' }, 'يمكنك لاحقًا ربط ملف إكسل بهذا الاستبيان، فيأخذ البرنامج البيانات منه أيضًا، ويربط الأعمدة بالوسوم المتشابهة تلقائيًا.') : null);
  }

  async function replaceDraftExcel(file) {
    const d = state.draft;
    let buf, table;
    try { buf = await file.arrayBuffer(); table = E.readTable(buf, file.name); } catch (e) { return toast(e.message, 'err'); }
    const oldKeys = d.headers.map((x) => x.key);
    const added = table.headers.filter((x) => !oldKeys.includes(x.key));
    const removed = d.headers.filter((x) => !table.headers.some((y) => y.key === x.key));
    if (!d._isNew && d.lastExcel && (added.length || removed.length)) {
      const ok = await modal({
        title: 'أسئلة الملف الجديد تختلف',
        body: h('div', null,
          added.length ? h('p', null, 'أسئلة جديدة: ' + added.map((x) => x.col).join('، ')) : null,
          removed.length ? h('p', null, 'أسئلة غير موجودة الآن: ' + removed.map((x) => x.col).join('، ')) : null,
          h('p', { class: 'muted' }, 'تحديث الأسئلة يغيّر طريقة التعرّف على الطلبات، فقد تظهر الطلبات القديمة كأنها جديدة.')),
        buttons: [{ label: 'حدّث الأسئلة', value: true }, { label: 'إلغاء', value: false, kind: 'ghost' }],
      });
      if (!ok) return;
    }
    const oldHeaders = d.headers;
    const oldFields = d.fields;
    const oldNameTag = d.nameField && oldFields[d.nameField] ? oldFields[d.nameField].tag : '';
    d.headers = table.headers.map(({ col, key }) => ({ col, key }));
    d.fields = E.makeFields(table.headers, d.fields);
    let mapped = [];
    if (d.mode === 'custom' && d._inspect && d._inspect.ok) mapped = E.autoMapTags(d, d._inspect.tags);
    if (d.source === 'template') {
      // نُبقي حقول القالب التي لا مقابل لها في الإكسل، لتُكتب يدويًا
      const covered = new Set(d.headers.map((x) => E.norm(d.fields[x.key].tag)));
      oldHeaders.forEach((x) => {
        const f = oldFields[x.key];
        if (!f || d.headers.some((y) => y.key === x.key) || covered.has(E.norm(f.tag))) return;
        d.headers.push(x); d.fields[x.key] = f;
      });
    }
    if (oldNameTag) {
      const k = d.headers.find((x) => E.norm(d.fields[x.key].tag) === E.norm(oldNameTag));
      if (k) d.nameField = k.key;
    }
    if (!d.headers.some((x) => x.key === d.nameField)) d.nameField = guessNameField(d.headers);
    d.lastExcel = { name: file.name, data: buf, date: Date.now() };
    d._table = table;
    render();
    toast(mapped.length ? ('حُدّث ملف الإكسل، ورُبط تلقائيًا: ' + mapped.map((m) => '{' + m.tag + '} ← ' + m.col).join('، ')) : 'حُدّث ملف الإكسل');
  }

  /* ---------- ② الحقول ---------- */
  let reportTimer = null;
  function scheduleReport() {
    clearTimeout(reportTimer);
    reportTimer = setTimeout(() => {
      const box = document.getElementById('tpl-report');
      if (box && state.draft) box.replaceWith(templateReport(state.draft, state.draft._inspect, 'tpl-report'));
      const g = document.getElementById('guide-report');
      if (g && state.draft) g.replaceWith(templateReport(state.draft, state.draft._guideInspect, 'guide-report'));
      const pv = document.getElementById('name-preview');
      if (pv && state.draft) pv.textContent = namePreview(state.draft);
    }, 250);
  }

  function edFields(d) {
    const t = d._table;
    const rows = d.headers.map((hd) => {
      const f = d.fields[hd.key];
      const tr = h('tr', { class: f.include === false ? 'off' : '' });
      const listHint = f.type === 'list'
        ? h('div', { class: 'muted', style: { marginTop: '4px' } }, 'في الوورد: ', tag((f.tag || '') + ' 1'), ' ', tag((f.tag || '') + ' 2'), ' … أو ', tag(f.tag || ''), ' للكل')
        : null;
      tr.append(
        h('td', null, h('input', { type: 'checkbox', checked: f.include !== false, title: 'إظهار هذه المعلومة',
          onchange: (e) => { f.include = e.target.checked; tr.className = f.include ? '' : 'off'; scheduleReport(); } })),
        h('td', { class: 'q' }, hd.col),
        h('td', null, h('input', { type: 'text', value: f.tag || '', 'aria-label': 'الوسم',
          oninput: (e) => { f.tag = e.target.value.replace(/[{}]/g, '').trim(); scheduleReport(); } }), listHint),
        h('td', null, h('select', { onchange: (e) => { setType(f, e.target.value); render(); } },
          TYPE_OPTIONS.map(([v, l]) => h('option', { value: v, selected: typeValue(f) === v }, l))),
          f.type === 'choice' ? h('input', { type: 'text', style: { marginTop: '6px' }, placeholder: 'الخيارات: افصل بينها بفاصلة', title: 'الخيارات التي تظهر في التعبئة اليدوية',
            value: (f.options || []).join('، '), oninput: (e) => { f.options = e.target.value.split(/[،,]/).map((x) => x.trim()).filter(Boolean); } }) : null),
        h('td', null, h('div', { class: 'sample', title: sampleValue(t, hd.key) }, sampleValue(t, hd.key) || '—')),
        h('td', null, h('button', { class: 'btn ghost small', title: 'نسخ الوسم', onclick: () => copyText('{' + (f.type === 'list' ? (f.tag + ' 1') : f.tag) + '}') }, 'نسخ')));
      return tr;
    });
    const setAll = (v) => { d.headers.forEach((x) => { d.fields[x.key].include = v; }); render(); };
    return card('٢', 'المعلومات التي تظهر في الوورد',
      h('p', { class: 'lead' }, 'أشّر على المعلومات التي تريدها، وأعطِ كل واحدة اسمًا قصيرًا (الوسم). في قالبك الخاص تضع الوسم بين قوسين مكان المعلومة، مثل ', tag('المسجد'), '. وفي القالب التلقائي تظهر المعلومات المؤشَّرة كلها في جدول.'),
      h('div', { class: 'row', style: { marginBottom: '10px' } },
        h('label', { class: 'field', style: { flex: '1', minWidth: '260px' } }, 'حقل اسم صاحب الطلب (لأسماء الملفات وقائمة الطلبات وخطاب التوجيه)',
          h('select', { onchange: (e) => { d.nameField = e.target.value; scheduleReport(); } },
            d.headers.map((x) => h('option', { value: x.key, selected: d.nameField === x.key }, x.col)))),
        h('span', { class: 'spacer' }),
        h('button', { class: 'btn ghost small', onclick: () => setAll(true) }, 'تأشير الكل'),
        h('button', { class: 'btn ghost small', onclick: () => setAll(false) }, 'إلغاء تأشير الكل')),
      h('div', { class: 'table-wrap' }, h('table', { class: 't' },
        h('tr', null, h('th', null, 'إظهار'), h('th', null, 'السؤال كما في الإكسل'), h('th', null, 'الوسم (الاسم القصير)'), h('th', null, 'النوع'), h('th', null, 'مثال من الملف'), h('th', null, '')),
        rows)));
  }

  /* ---------- ③ القالب ---------- */
  function templateReport(d, ins, id) {
    const box = h('div', { id });
    if (!ins) return box;
    if (!ins.ok) {
      box.appendChild(h('div', { class: 'note err' }, h('b', null, 'في القالب أخطاء تمنع استعماله:'), h('ul', null, ins.errors.map((x) => h('li', null, x)))));
      return box;
    }
    if (!ins.tags.length) {
      box.appendChild(h('div', { class: 'note warn' }, 'لم أجد أي وسم {…} في هذا الملف. ضع الوسوم مكان المعلومات ثم ارفعه من جديد.'));
      return box;
    }
    const good = [], bad = [];
    ins.tags.forEach((x) => {
      const r = E.describeTag(d, x.tag);
      if (r.ok) good.push(x); else bad.push({ x, r });
    });
    box.appendChild(h('div', { class: 'note info' }, 'وجدت في القالب ' + ins.tags.length + ' وسمًا، منها ' + good.length + ' معروفًا.',
      h('div', { class: 'tagchips' }, good.map((x) => h('span', { class: 'chip green' }, '✔ ', '{' + (x.kind === 'section' ? '#' : '') + x.tag + '}')))));
    if (bad.length) {
      box.appendChild(h('div', { class: 'note warn' }, h('b', null, 'وسوم غير معروفة (ستظهر فارغة):'),
        h('ul', null, bad.map(({ x, r }) => h('li', null, '{' + x.tag + '}',
          r.suggestion ? h('span', null, ' — هل تقصد ', tag(r.suggestion), '؟') : (r.warn ? '' : ' — غيّر الوسم في الوورد، أو اكتبه في جدول الحقول بالاسم نفسه.'),
          r.warn ? ' — ' + r.warn + '.' : '')))));
    }
    return box;
  }

  async function loadTemplateFile(file, which) {
    const d = state.draft;
    if (!/\.docx$/i.test(file.name)) return toast('اختر ملف وورد بصيغة ‎.docx‎', 'err');
    const data = await file.arrayBuffer();
    const ins = E.inspectTemplate(data);
    let note = '';
    if (which === 'guide') { d.guideTemplate = { name: file.name, data }; d._guideInspect = ins; d.guideMode = 'custom'; }
    else {
      d.template = { name: file.name, data }; d._inspect = ins; d.mode = 'custom';
      if (ins.ok && d.source === 'template' && !d.lastExcel) {
        const fromTpl = E.fieldsFromTemplate(ins.tags);
        let added = 0;
        fromTpl.headers.forEach((x) => {
          if (d.headers.some((y) => y.key === x.key) || E.describeTag(d, x.col).ok) return;
          d.headers.push(x); d.fields[x.key] = fromTpl.fields[x.key]; added++;
        });
        if (added) note = ' — أُضيف ' + added + ' حقلًا جديدًا من القالب';
      } else if (ins.ok) {
        const mapped = E.autoMapTags(d, ins.tags);
        if (mapped.length) note = ' — رُبط تلقائيًا: ' + mapped.map((m) => '{' + m.tag + '} ← ' + m.col).join('، ');
      }
    }
    render();
    toast(ins.ok ? 'رُفع القالب: ' + file.name + note : 'في القالب أخطاء، انظر التفاصيل', ins.ok ? '' : 'err');
  }

  function edTemplate(d) {
    const body = [];
    body.push(radioCards('mode', d.mode, [
      ['custom', 'قالب وورد خاص بي', 'تضع الوسوم بنفسك داخل ملف الوورد الرسمي للاستبيان'],
      ['auto', 'قالب تلقائي', 'يصنع البرنامج مستندًا بترويسة المديرية وجدولٍ بالمعلومات المؤشَّرة'],
    ], (v) => { d.mode = v; render(); }));
    if (d.mode === 'custom') {
      if (d.template) {
        body.push(h('div', { class: 'row', style: { marginTop: '12px' } },
          h('span', { class: 'chip' }, 'القالب: ' + d.template.name),
          h('span', { class: 'spacer' }),
          h('button', { class: 'btn ghost small', onclick: async () => { const f = await pickFile('.docx'); if (f) loadTemplateFile(f, 'main'); } }, 'استبدال القالب'),
          h('button', { class: 'btn ghost small', onclick: () => download(d.template.data, d.template.name, DOCX_MIME) }, 'تنزيل القالب')));
        body.push(templateReport(d, d._inspect, 'tpl-report'));
      } else {
        body.push(h('div', { style: { marginTop: '12px' } }, dropZone('ارفع ملف الوورد (القالب) الذي وضعت فيه الوسوم', 'صيغة ‎.docx‎ — يمكنك سحبه وإفلاته هنا', '.docx', (f) => loadTemplateFile(f, 'main'), true)));
      }
    } else {
      body.push(h('div', { class: 'row', style: { marginTop: '12px' } },
        h('span', { class: 'muted' }, 'عنوان المستند هو اسم الاستبيان، وتظهر فيه المعلومات المؤشَّرة بترتيبها في الإكسل.'),
        h('span', { class: 'spacer' }),
        h('button', { class: 'btn ghost small', onclick: () => download(builtin().auto, 'القالب التلقائي.docx', DOCX_MIME) }, 'تنزيل القالب التلقائي لتعديله')));
      body.push(h('p', { class: 'muted' }, 'إن عدّلت القالب التلقائي في الوورد، فارفعه بعد ذلك باختيار «قالب وورد خاص بي».'));
    }
    return card('٣', 'قالب الوورد', ...body);
  }

  /* ---------- ④ الطلبات الناقصة ---------- */
  function edRules(d) {
    const t = d._table;
    const list = h('div');
    d.rules.forEach((rule, i) => {
      const dl = 'dl-' + i;
      const vals = distinctValues(t, rule.key);
      const noValue = rule.op === 'empty' || rule.op === 'notempty';
      list.appendChild(h('div', { class: 'rule' },
        h('div', { class: 'lbl' }, 'الشرط ' + (i + 1) + ': يكون الطلب مكتملًا إذا كان جواب'),
        h('select', { onchange: (e) => { rule.key = e.target.value; render(); } },
          d.headers.map((x) => h('option', { value: x.key, selected: rule.key === x.key }, x.col))),
        h('select', { onchange: (e) => { rule.op = e.target.value; render(); } },
          Object.keys(E.OPS).map((op) => h('option', { value: op, selected: rule.op === op }, E.OPS[op].label))),
        noValue ? h('span', { class: 'muted' }, '—') : h('input', { type: 'text', value: rule.value || '', list: dl, placeholder: 'القيمة، مثل: نعم', oninput: (e) => { rule.value = e.target.value; } }),
        h('button', { class: 'btn danger small', onclick: () => { d.rules.splice(i, 1); render(); } }, 'حذف'),
        h('datalist', { id: dl }, vals.map((v) => h('option', { value: v }))),
        h('div', { class: 'msg' }, h('span', null, 'وإلا يُكتب في خطاب التوجيه:'),
          h('input', { type: 'text', value: rule.message || '', placeholder: 'مثل: رفع السيرة الذاتية عبر الرابط المعتمد.', oninput: (e) => { rule.message = e.target.value; } }))));
    });
    const body = [
      h('p', { class: 'lead' }, 'حدّد ما يجب أن يتحقق ليكون الطلب مكتملًا. الطلب الذي لا يحقق شرطًا يُعدّ ناقصًا، ويُكتب له في خطاب التوجيه نصّ ذلك الشرط. إن لم تضع شروطًا فكل الطلبات مكتملة.'),
      list,
      h('button', { class: 'btn ghost', onclick: () => {
        d.rules.push({ key: d.headers[0].key, op: 'eq', value: '', message: '' });
        render();
      } }, '＋ إضافة شرط'),
    ];
    if (d.rules.length) {
      body.push(h('h3', { style: { margin: '18px 0 8px', fontSize: '17px' } }, 'ماذا يُفعل بالطلب الناقص؟'));
      body.push(radioCards('inc', d.settings.incompleteAction, [
        ['guide', 'توليد خطاب توجيه لصاحبه', 'يُذكر فيه ما ينقص طلبه'],
        ['exclude', 'استبعاده', 'لا يُولَّد له شيء'],
      ], (v) => { d.settings.incompleteAction = v; render(); }));
      body.push(h('p', { class: 'muted' }, 'هذا هو الاختيار الافتراضي، ويمكنك تغييره لكل طلب عند التوليد.'));
      body.push(h('h3', { style: { margin: '18px 0 8px', fontSize: '17px' } }, 'خطاب التوجيه'));
      body.push(radioCards('guide', d.guideMode, [
        ['builtin', 'الخطاب الافتراضي', 'خطاب رسمي جاهز بترويسة المديرية'],
        ['custom', 'خطاب من تصميمي', 'ملف وورد تكتبه أنت وتضع فيه الوسوم'],
      ], (v) => { d.guideMode = v; render(); }));
      if (d.guideMode === 'custom') {
        if (d.guideTemplate) {
          body.push(h('div', { class: 'row', style: { marginTop: '12px' } },
            h('span', { class: 'chip' }, 'الخطاب: ' + d.guideTemplate.name),
            h('span', { class: 'spacer' }),
            h('button', { class: 'btn ghost small', onclick: async () => { const f = await pickFile('.docx'); if (f) loadTemplateFile(f, 'guide'); } }, 'استبدال الخطاب')));
          body.push(templateReport(d, d._guideInspect, 'guide-report'));
        } else {
          body.push(h('div', { style: { marginTop: '12px' } }, dropZone('ارفع ملف الوورد لخطاب التوجيه', 'استعمل فيه ' + '{أسباب التوجيه}' + ' و{الاسم} وغيرها', '.docx', (f) => loadTemplateFile(f, 'guide'), true)));
        }
      }
      body.push(h('div', { class: 'row', style: { marginTop: '10px' } },
        h('span', { class: 'spacer' }),
        h('button', { class: 'btn ghost small', onclick: () => download(builtin().guide, 'خطاب التوجيه الافتراضي.docx', DOCX_MIME) }, 'تنزيل الخطاب الافتراضي لتعديله')));
    }
    return card('٤', 'الطلبات الناقصة وخطاب التوجيه', ...body);
  }

  /* ---------- ⑤ الإخراج ---------- */
  function namePreview(d) {
    const t = d._table;
    if (!t || !t.rows.length) return '';
    return 'مثال: ' + E.fileNameFor(d.settings.filePattern, d, t.rows[0]) + '.docx';
  }

  function edOutput(d) {
    const s = d.settings;
    const patternInput = h('input', { type: 'text', value: s.filePattern, oninput: (e) => { s.filePattern = e.target.value; scheduleReport(); } });
    const addTag = (tg) => {
      s.filePattern = (s.filePattern ? s.filePattern.trim() + ' - ' : '') + '{' + tg + '}';
      patternInput.value = s.filePattern; scheduleReport();
    };
    const fieldTags = d.headers.filter((x) => d.fields[x.key].include !== false && d.fields[x.key].type !== 'list').slice(0, 8).map((x) => d.fields[x.key].tag).filter((t) => t && t.length <= 24);
    return card('٥', 'الإخراج',
      h('label', { class: 'field', style: { marginTop: '8px' } }, 'اسم كل ملف ناتج', patternInput),
      h('div', { class: 'tagchips' }, h('span', { class: 'muted' }, 'إضافة:'),
        ['التسلسل', 'الاسم'].concat(fieldTags).map((tg) => h('button', { class: 'btn ghost small', onclick: () => addTag(tg) }, '{' + tg + '}'))),
      h('div', { id: 'name-preview', class: 'muted', style: { marginTop: '6px' } }, namePreview(d)),
      h('div', { class: 'row', style: { marginTop: '14px', alignItems: 'flex-end' } },
        h('label', { class: 'field', style: { flex: '2', minWidth: '240px' } }, 'اسم الجهة (يظهر مكان {الجهة} في الترويسة والتوقيع)',
          h('input', { type: 'text', value: s.entity, oninput: (e) => { s.entity = e.target.value; } })),
        h('label', { class: 'field', style: { flex: '1', minWidth: '180px' } }, 'ما يُكتب إن كان الجواب فارغًا',
          h('input', { type: 'text', value: s.emptyValue, placeholder: 'يُترك فارغًا', oninput: (e) => { s.emptyValue = e.target.value; } })),
        h('label', { class: 'field', style: { width: '150px' } }, 'علامة الاختيار',
          h('input', { type: 'text', value: s.markOn, oninput: (e) => { s.markOn = e.target.value || '●'; } }))),
      h('h3', { style: { margin: '18px 0 8px', fontSize: '17px' } }, 'طريقة الإخراج المفضّلة'),
      outputModeCards(s, () => render()));
  }

  function outputModeCards(s, after) {
    return radioCards('out', s.outputMode, [
      ['zip', 'ملفات منفصلة', 'ملف وورد لكل طلب باسمه، كلها داخل ملف مضغوط واحد'],
      ['merged', 'ملف واحد مجمّع', 'كل الطلبات في ملف وورد واحد، كل طلب يبدأ بصفحة جديدة — للطباعة دفعة واحدة'],
      ['both', 'الاثنان معًا', 'الملف المضغوط والملف المجمّع'],
    ], (v) => { s.outputMode = v; after(); });
  }

  async function saveDraft() {
    const d = state.draft;
    d.name = String(d.name || '').trim();
    if (!d.name) return toast('اكتب اسم الاستبيان أولًا', 'err');
    if (d.mode === 'custom' && !d.template) return toast('ارفع قالب الوورد، أو اختر «قالب تلقائي»', 'err');
    if (d.mode === 'custom' && d._inspect && !d._inspect.ok) return toast('أصلح أخطاء القالب أولًا ثم ارفعه من جديد', 'err');
    if (d.rules.length && d.guideMode === 'custom' && !d.guideTemplate) return toast('ارفع ملف خطاب التوجيه، أو اختر «الخطاب الافتراضي»', 'err');
    const tags = d.mode === 'custom' && d._inspect ? d._inspect.tags : [];
    const unknown = tags.filter((x) => !E.describeTag(d, x.tag).ok);
    if (unknown.length) {
      const ok = await modal({
        title: 'في القالب وسوم غير معروفة',
        body: h('div', null, h('p', null, unknown.map((x) => '{' + x.tag + '}').join('  ')), h('p', { class: 'muted' }, 'ستظهر هذه الوسوم فارغة في الملفات. هل تحفظ رغم ذلك؟')),
        buttons: [{ label: 'احفظ رغم ذلك', value: true }, { label: 'رجوع للتصحيح', value: false, kind: 'ghost' }],
      });
      if (!ok) return;
    }
    const out = {};
    Object.keys(d).forEach((k) => { if (!k.startsWith('_')) out[k] = d[k]; });
    try {
      await S.put(out);
    } catch (e) {
      return toast('تعذّر الحفظ في المتصفح: ' + (e.message || e), 'err');
    }
    await refreshProfiles();
    toast('حُفظ الاستبيان «' + out.name + '»');
    state.draft = null;
    openRun(out.id);
  }

  async function resetProcessed() {
    const d = state.draft;
    const n = Object.keys(d.processed || {}).length;
    if (!n) return toast('السجل فارغ أصلًا');
    const ok = await modal({ title: 'تصفير سجل الطلبات المنجزة؟', body: 'ستظهر كل الطلبات (' + n + ') كأنها جديدة في المرة القادمة.', buttons: [{ label: 'نعم، صفّر السجل', value: true, kind: 'danger' }, { label: 'إلغاء', value: false, kind: 'ghost' }] });
    if (!ok) return;
    d.processed = {};
    const p = await S.get(d.id);
    if (p) { p.processed = {}; await S.put(p); }
    render();
    toast('صُفّر السجل');
  }

  /* ================================================================== *
   *  التوليد
   * ================================================================== */

  async function openRun(id, incoming) {
    const p = await S.get(id);
    if (!p) return toast('لم أجد هذا الاستبيان.', 'err');
    p.settings = Object.assign({}, DEFAULT_SETTINGS, p.settings || {});
    state.current = p;
    state.view = 'run';
    p.manualRows = p.manualRows || [];
    state.run = { filter: 'new', table: null, rows: [], results: null, progress: 0, busy: false, errors: [],
      manual: { open: p.source === 'template' && !p.lastExcel && !incoming, values: {}, editing: null } };
    if (incoming) {
      p.lastExcel = { name: incoming.file.name, data: incoming.buf, date: Date.now() };
      await S.put(p);
      await refreshProfiles();
      setRunTable(incoming.table);
      toast('حُفظ ملف الإكسل الجديد مع الاستبيان');
    } else if (p.lastExcel && p.lastExcel.data) {
      try { setRunTable(E.readTable(p.lastExcel.data, p.lastExcel.name)); } catch (e) { toast(e.message, 'err'); }
    }
    if (!state.run.table) rebuildRows();
    render();
    window.scrollTo({ top: 0 });
  }

  function setRunTable(table) {
    const p = state.current;
    const run = state.run;
    run.table = table;
    run.results = null;
    run.errors = [];
    run.rows = [];
    rebuildRows();
    const keys = table.headers.map((x) => x.key);
    run.missing = p.headers.filter((x) => !keys.includes(x.key) && (p.fields[x.key] || {}).include !== false);
    run.extra = table.headers.filter((x) => !p.headers.some((y) => y.key === x.key));
    run.similarity = E.headerSimilarity(p.headers.map((x) => x.key), keys);
    if (!run.rows.some((x) => !x.done) && run.rows.length) run.filter = 'all';
  }

  function makeRow(p, r, fp, manual) {
    const ev = E.evaluateRow(p, r);
    const done = (p.processed || {})[fp] || null;
    return { r, fp, ev, done, manual, name: rowName(p, r), action: ev.complete ? 'main' : (p.settings.incompleteAction || 'guide'), sel: !done };
  }

  /** يجمع طلبات ملف الإكسل والطلبات اليدوية في قائمة واحدة، محافظًا على التحديد السابق */
  function rebuildRows() {
    const p = state.current;
    const run = state.run;
    const prev = new Map((run.rows || []).map((x) => [x.fp, x]));
    const rows = [];
    if (run.table) run.table.rows.forEach((r) => rows.push(makeRow(p, r, E.rowFingerprint(p, r), false)));
    (p.manualRows || []).forEach((m, i) => {
      rows.push(makeRow(p, { seq: i + 1, values: m.values, manual: true, id: m.id }, 'm:' + m.id, true));
    });
    rows.forEach((x) => {
      const old = prev.get(x.fp);
      if (old && !!old.done === !!x.done) { x.sel = old.sel; if (!x.ev.complete) x.action = old.action; }
    });
    run.rows = rows;
  }

  async function loadRunExcel(file) {
    const p = state.current;
    let buf, table;
    try { buf = await file.arrayBuffer(); table = E.readTable(buf, file.name); } catch (e) { return toast(e.message, 'err'); }
    const sim = E.headerSimilarity(p.headers.map((x) => x.key), table.headers.map((x) => x.key));
    if (sim < 0.6) {
      const ok = await modal({
        title: 'هذا الملف لا يشبه هذا الاستبيان',
        body: 'تطابق الأسئلة ' + Math.round(sim * 100) + '٪ فقط. ربما رفعت ملف استبيان آخر.',
        buttons: [{ label: 'اربطه بهذا الاستبيان', value: 'link' }, { label: 'افتحه كاستبيان جديد', value: 'new', kind: 'ghost' }, { label: 'إلغاء', value: null, kind: 'ghost' }],
      });
      if (ok === 'new') return startDraft(file, buf, table);
      if (ok === 'link') { await openEditor(p.id); return replaceDraftExcel(file); }
      return;
    }
    p.lastExcel = { name: file.name, data: buf, date: Date.now() };
    await S.put(p);
    await refreshProfiles();
    state.run.filter = 'new';
    setRunTable(table);
    render();
    const n = state.run.rows.filter((x) => !x.done).length;
    toast(n ? ('وُجد ' + n + ' طلبًا جديدًا') : 'لا توجد طلبات جديدة في هذا الملف');
  }

  function templateFor(p, action) {
    if (action === 'guide') return p.guideMode === 'custom' && p.guideTemplate ? p.guideTemplate.data : builtin().guide;
    return p.mode === 'custom' && p.template ? p.template.data : builtin().auto;
  }

  function viewRun() {
    const p = state.current;
    const run = state.run;
    const wrap = h('div');

    wrap.appendChild(h('div', { class: 'card' },
      h('div', { class: 'row' },
        h('div', { style: { flex: '1', minWidth: '240px' } },
          h('div', { class: 'kick' }, p.mode === 'auto' ? 'القالب: تلقائي' : 'القالب: ' + ((p.template && p.template.name) || '—')),
          h('h2', { style: { fontSize: '24px' } }, p.name)),
        h('button', { class: 'btn ghost small', onclick: () => openEditor(p.id) }, 'الإعدادات'),
        h('button', { class: 'btn ghost small', onclick: () => exportFlow(p) }, 'تصدير'),
        h('button', { class: 'btn ghost small', onclick: () => goHome('saved') }, 'كل الاستبيانات')),
      h('div', { style: { marginTop: '12px' } },
        dropZone(p.lastExcel ? 'ارفع النسخة الأحدث من ملف الإكسل' : (p.source === 'template' ? 'ارفع ملف إكسل (اختياري)' : 'ارفع ملف الإكسل الخاص بهذا الاستبيان'),
          p.lastExcel ? ('الملف المحفوظ الآن: ' + p.lastExcel.name + ' — رُفع في ' + fmtDate(p.lastExcel.date)) : 'أو اكتب الطلبات بيدك في «التعبئة اليدوية»',
          '.xlsx,.xlsm,.csv', loadRunExcel, true))));

    wrap.appendChild(viewManual());

    if (!run.rows.length) {
      if (!run.table && !run.manual.open) wrap.appendChild(h('div', { class: 'card empty' }, h('div', { class: 'big-ico' }, '📄'), h('p', null, 'ارفع ملف الإكسل لتظهر الطلبات هنا، أو أضف طلبًا يدويًا.')));
      return wrap;
    }

    // تنبيهات
    if (run.table && run.missing && run.missing.length) {
      wrap.appendChild(h('div', { class: 'note warn' }, h('b', null, 'أسئلة مؤشَّرة غير موجودة في هذا الملف (ستظهر فارغة): '), run.missing.map((x) => x.col).join('، ')));
    }
    if (run.table && run.extra && run.extra.length) {
      wrap.appendChild(h('div', { class: 'note info' }, 'في الملف أسئلة جديدة لم تكن في الاستبيان: ', run.extra.map((x) => x.col).join('، '), ' — لإضافتها افتح «الإعدادات» ثم «تغيير ملف الإكسل».'));
    }
    if (p.mode === 'custom' && p.template) {
      const ins = E.inspectTemplate(p.template.data);
      const bad = ins.ok ? ins.tags.filter((x) => !E.describeTag(p, x.tag).ok) : [];
      if (!ins.ok) wrap.appendChild(h('div', { class: 'note err' }, 'في القالب أخطاء: ', ins.errors.join(' · ')));
      else if (bad.length) wrap.appendChild(h('div', { class: 'note warn' }, 'وسوم غير معروفة في القالب ستظهر فارغة: ', bad.map((x) => '{' + x.tag + '}').join('  '), ' — صحّحها من «الإعدادات».'));
    }

    const all = run.rows;
    const nNew = all.filter((x) => !x.done).length;
    const nInc = all.filter((x) => !x.ev.complete).length;
    const shown = run.filter === 'new' ? all.filter((x) => !x.done) : all;

    const head = h('div', { class: 'card' },
      h('div', { class: 'row' },
        h('h2', { style: { fontSize: '20px' } }, 'الطلبات'),
        h('span', { class: 'chip' }, 'الكل: ' + all.length),
        all.some((x) => x.manual) ? h('span', { class: 'chip manual' }, 'يدوية: ' + all.filter((x) => x.manual).length) : null,
        h('span', { class: 'chip gold' }, 'جديدة: ' + nNew),
        h('span', { class: 'chip green' }, 'مكتملة: ' + (all.length - nInc)),
        nInc ? h('span', { class: 'chip red' }, 'ناقصة: ' + nInc) : null,
        h('span', { class: 'spacer' }),
        h('div', { class: 'tabs', style: { margin: '0', padding: '4px' } },
          h('button', { class: run.filter === 'new' ? 'on' : '', style: { minWidth: '120px' }, onclick: () => { run.filter = 'new'; render(); } }, 'الجديدة فقط'),
          h('button', { class: run.filter === 'all' ? 'on' : '', style: { minWidth: '120px' }, onclick: () => { run.filter = 'all'; render(); } }, 'الكل'))));

    if (!shown.length) {
      head.appendChild(h('div', { class: 'empty' },
        h('div', { class: 'big-ico' }, '✔'),
        h('h3', null, 'لا توجد طلبات جديدة'),
        h('p', { class: 'muted' }, 'كل الطلبات (' + all.length + ') سبق توليدها. اختر «الكل» لإعادة توليد أي طلب.')));
      head.appendChild(resultsBox());
      wrap.appendChild(head);
      return wrap;
    }

    const allSel = shown.every((x) => x.sel);
    const someSel = shown.some((x) => x.sel);
    const tbl = h('table', { class: 't' },
      h('tr', null,
        h('th', null, h('input', { type: 'checkbox', checked: allSel, indeterminate: someSel && !allSel, title: 'تحديد الكل',
          onchange: (e) => { shown.forEach((x) => { x.sel = e.target.checked; }); render(); } })),
        h('th', null, '#'), h('th', null, 'صاحب الطلب'), h('th', null, 'الحالة'), h('th', null, 'ما يُولَّد'), h('th', null, '')));
    shown.forEach((x) => {
      const actionSel = x.ev.complete
        ? h('span', { class: 'muted' }, 'الطلب')
        : h('select', { onchange: (e) => { x.action = e.target.value; render(); } },
          h('option', { value: 'guide', selected: x.action === 'guide' }, 'خطاب توجيه'),
          h('option', { value: 'exclude', selected: x.action === 'exclude' }, 'استبعاد'),
          h('option', { value: 'main', selected: x.action === 'main' }, 'الطلب رغم النقص'));
      tbl.appendChild(h('tr', { class: x.sel ? '' : 'off' },
        h('td', null, h('input', { type: 'checkbox', checked: x.sel, onchange: (e) => { x.sel = e.target.checked; render(); } })),
        h('td', null, x.manual ? h('span', { class: 'chip manual' }, 'يدوي ' + x.r.seq) : x.r.seq),
        h('td', null, x.name),
        h('td', null,
          x.ev.complete ? h('span', { class: 'chip green' }, 'مكتمل') : h('span', { class: 'chip red' }, 'ناقص'),
          x.done ? h('span', { class: 'chip gray', style: { marginInlineStart: '6px' } }, 'سبق توليده ' + fmtDate(x.done.t)) : h('span', { class: 'chip gold', style: { marginInlineStart: '6px' } }, 'جديد'),
          x.ev.reasons.length ? h('ul', { class: 'reasons' }, x.ev.reasons.map((m) => h('li', null, m))) : null),
        h('td', null, actionSel),
        h('td', null, h('div', { class: 'row', style: { gap: '6px', flexWrap: 'nowrap' } },
          h('button', { class: 'btn ghost small', disabled: x.action === 'exclude', title: 'تنزيل ملف هذا الطلب وحده', onclick: () => downloadOne(x) }, 'تنزيل'),
          x.manual ? h('button', { class: 'btn ghost small', onclick: () => editManual(x.r.id) }, 'تعديل') : null,
          x.manual ? h('button', { class: 'btn danger small', onclick: () => deleteManual(x.r.id) }, 'حذف') : null))));
    });
    head.appendChild(h('div', { class: 'table-wrap', style: { marginTop: '12px' } }, tbl));
    wrap.appendChild(head);
    wrap.appendChild(genBar());
    return wrap;
  }

  /* ---------- التعبئة اليدوية ---------- */
  function manualFields(p) {
    return p.headers.filter((x) => (p.fields[x.key] || {}).include !== false);
  }

  function updateManualWarn() {
    const box = document.getElementById('manual-warn');
    if (!box) return;
    const p = state.current;
    box.innerHTML = '';
    if (!(p.rules || []).length) return;
    const ev = E.evaluateRow(p, { seq: 0, values: state.run.manual.values });
    if (!ev.complete) box.appendChild(h('div', { class: 'note warn' }, 'هذا الطلب ناقص حسب شروط الاستبيان:', h('ul', null, ev.reasons.map((m) => h('li', null, m)))));
  }

  function manualInput(hd, f, m, idx) {
    const run = state.run;
    const set = (v) => { m.values[hd.key] = v; updateManualWarn(); };
    const val = m.values[hd.key] || '';
    let input;
    if (f.type === 'date') {
      input = h('input', { type: 'date', value: /^\d{4}-\d{2}-\d{2}$/.test(val) ? val : '', oninput: (e) => set(e.target.value) });
    } else if (f.type === 'list') {
      input = h('textarea', { rows: '3', placeholder: 'كل سطر عنصر', oninput: (e) => set(e.target.value) }, val);
    } else if (f.type === 'choice') {
      const opts = (f.options && f.options.length) ? f.options : distinctValues(run.table, hd.key);
      input = h('select', { onchange: (e) => set(e.target.value) },
        h('option', { value: '' }, '— اختر —'),
        opts.map((o) => h('option', { value: o, selected: val === o }, o)));
    } else {
      const sample = sampleValue(run.table, hd.key);
      const long = sample.length > 60 || /(اذكر|أهداف|اهداف|ملاحظات|أسباب|اسباب|تفاصيل|شرح|البرامج)/.test(hd.col);
      if (long) input = h('textarea', { rows: '2', oninput: (e) => set(e.target.value) }, val);
      else {
        const vals = distinctValues(run.table, hd.key);
        const dl = 'mdl-' + idx;
        input = h('span', null,
          h('input', { type: 'text', value: val, list: vals.length ? dl : null, oninput: (e) => set(e.target.value) }),
          vals.length ? h('datalist', { id: dl }, vals.map((v) => h('option', { value: v }))) : null);
      }
    }
    return h('label', { class: 'field' }, hd.col, input);
  }

  function viewManual() {
    const p = state.current;
    const run = state.run;
    const m = run.manual;
    if (!m.open) {
      return h('div', { class: 'card manual-head' },
        h('b', { style: { color: 'var(--forest-2)', fontSize: '18px' } }, 'التعبئة اليدوية'),
        h('span', { class: 'muted' }, 'اكتب طلبًا بيدك دون ملف إكسل، ثم نزّل ملف الوورد أو أضفه إلى قائمة الطلبات'),
        h('span', { class: 'spacer' }),
        h('button', { class: 'btn', onclick: () => { m.open = true; render(); } }, '＋ إضافة طلب يدويًا'));
    }
    const grid = h('div', { class: 'mgrid' });
    manualFields(p).forEach((hd, i) => grid.appendChild(manualInput(hd, p.fields[hd.key] || {}, m, i)));
    const box = h('section', { class: 'card', id: 'manual-card' },
      h('h2', null, m.editing ? 'تعديل طلب يدوي' : 'التعبئة اليدوية'),
      h('p', { class: 'lead' }, 'اكتب المعلومات ثم نزّل ملف الوورد مباشرة، أو أضف الطلب إلى القائمة لتولّده مع غيره بالطريقة التي تختارها. الحقول التي تتركها فارغة تبقى فارغة في الملف.'),
      grid,
      h('div', { id: 'manual-warn' }),
      h('div', { class: 'row', style: { marginTop: '14px' } },
        m.editing
          ? h('button', { class: 'btn', onclick: saveManualEdit }, 'حفظ التعديل')
          : h('button', { class: 'btn', onclick: addManual }, 'إضافة إلى قائمة الطلبات'),
        h('button', { class: 'btn gold', onclick: downloadManual }, '⬇ تنزيل ملف الوورد الآن'),
        h('button', { class: 'btn ghost', onclick: () => { m.values = {}; m.editing = null; render(); } }, m.editing ? 'إلغاء التعديل' : 'مسح الحقول'),
        h('span', { class: 'spacer' }),
        (run.table || run.rows.length) ? h('button', { class: 'btn ghost small', onclick: () => { m.open = false; m.editing = null; m.values = {}; render(); } }, 'إغلاق النموذج') : null));
    setTimeout(updateManualWarn, 0);
    return box;
  }

  function manualValues() {
    const out = {};
    Object.keys(state.run.manual.values).forEach((k) => {
      const v = String(state.run.manual.values[k] || '').replace(/\r\n?/g, '\n').trim();
      if (v) out[k] = v;
    });
    return out;
  }

  async function addManual() {
    const p = state.current;
    const values = manualValues();
    if (!Object.keys(values).length) return toast('اكتب معلومة واحدة على الأقل', 'err');
    p.manualRows = p.manualRows || [];
    p.manualRows.push({ id: newId(), values, createdAt: Date.now() });
    await S.put(p);
    await refreshProfiles();
    state.run.manual.values = {};
    state.run.results = null;
    rebuildRows();
    state.run.filter = 'new';
    render();
    toast('أُضيف الطلب إلى القائمة — يمكنك إضافة غيره');
    const c = document.getElementById('manual-card'); if (c) c.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function editManual(id) {
    const p = state.current;
    const row = (p.manualRows || []).find((x) => x.id === id);
    if (!row) return;
    state.run.manual = { open: true, values: Object.assign({}, row.values), editing: id };
    render();
    const c = document.getElementById('manual-card'); if (c) c.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  async function saveManualEdit() {
    const p = state.current;
    const m = state.run.manual;
    const row = (p.manualRows || []).find((x) => x.id === m.editing);
    if (!row) return;
    const values = manualValues();
    if (!Object.keys(values).length) return toast('اكتب معلومة واحدة على الأقل', 'err');
    row.values = values;
    row.updatedAt = Date.now();
    if (p.processed) delete p.processed['m:' + row.id];
    await S.put(p);
    state.run.manual = { open: true, values: {}, editing: null };
    rebuildRows();
    render();
    toast('حُفظ التعديل — صار الطلب جديدًا ليُولَّد من جديد');
  }

  async function deleteManual(id) {
    const p = state.current;
    const row = (p.manualRows || []).find((x) => x.id === id);
    if (!row) return;
    const ok = await modal({ title: 'حذف هذا الطلب اليدوي؟', body: 'سيُحذف «' + rowName(p, { seq: 0, values: row.values }) + '» من القائمة.', buttons: [{ label: 'نعم، احذفه', value: true, kind: 'danger' }, { label: 'إلغاء', value: false, kind: 'ghost' }] });
    if (!ok) return;
    p.manualRows = p.manualRows.filter((x) => x.id !== id);
    if (p.processed) delete p.processed['m:' + id];
    if (state.run.manual.editing === id) state.run.manual = { open: true, values: {}, editing: null };
    await S.put(p);
    await refreshProfiles();
    rebuildRows();
    render();
    toast('حُذف الطلب');
  }

  function downloadManual() {
    const p = state.current;
    const values = manualValues();
    if (!Object.keys(values).length) return toast('اكتب معلومة واحدة على الأقل', 'err');
    const r = { seq: (p.manualRows || []).length + 1, values, manual: true };
    const x = { r, ev: E.evaluateRow(p, r), action: 'main', name: rowName(p, r) };
    try {
      const data = E.renderOne(templateFor(p, 'main'), p, r, extraFor(p, x, {}));
      download(data, outName(p, x), DOCX_MIME);
      toast('نُزّل الملف. إن أردت الاحتفاظ به في القائمة فاضغط «إضافة إلى قائمة الطلبات»');
    } catch (e) {
      toast('تعذّر توليد الملف: ' + explain(e), 'err');
    }
  }

  function genBar() {
    const p = state.current;
    const run = state.run;
    const s = p.settings;
    const sel = run.rows.filter((x) => x.sel);
    const nMain = sel.filter((x) => x.action === 'main').length;
    const nGuide = sel.filter((x) => x.action === 'guide').length;
    const nEx = sel.filter((x) => x.action === 'exclude').length;
    const bar = h('div', { class: 'genbar' },
      h('div', { class: 'row' },
        h('b', { style: { color: 'var(--forest-2)' } }, 'طريقة الإخراج:'),
        ...[['zip', 'ملفات منفصلة (مضغوطة)'], ['merged', 'ملف واحد مجمّع'], ['both', 'الاثنان معًا']].map(([v, l]) =>
          h('label', { class: 'row', style: { gap: '6px', cursor: 'pointer' } },
            h('input', { type: 'radio', name: 'outmode', checked: s.outputMode === v, onchange: () => { s.outputMode = v; S.put(p); render(); } }), l)),
        h('span', { class: 'spacer' }),
        h('label', { class: 'row', style: { gap: '6px', cursor: 'pointer' } },
          h('input', { type: 'checkbox', checked: s.markDone !== false, onchange: (e) => { s.markDone = e.target.checked; S.put(p); } }),
          'تعليم المولَّد كمنجز (لا يظهر في «الجديدة» بعد ذلك)')),
      h('div', { class: 'row', style: { marginTop: '10px' } },
        h('button', { class: 'btn big', disabled: run.busy || !(nMain + nGuide + nEx), onclick: generate },
          run.busy ? 'جارٍ التوليد…' : ('توليد الملفات (' + (nMain + nGuide) + ')')),
        h('span', { class: 'muted' },
          'الطلبات: ' + nMain + (nGuide ? ' · خطابات التوجيه: ' + nGuide : '') + (nEx ? ' · المستبعَدة: ' + nEx : ''))),
      run.busy ? h('div', { class: 'progress' }, h('div', { id: 'prog', style: { width: Math.round(run.progress * 100) + '%' } })) : null);
    bar.appendChild(resultsBox());
    return bar;
  }

  /** نتائج آخر توليد: أزرار التنزيل والأخطاء */
  function resultsBox() {
    const run = state.run;
    const box = h('div', { id: 'results-box' });
    if (run.errors && run.errors.length) {
      box.appendChild(h('div', { class: 'note err' }, h('b', null, 'تعذّر توليد بعض الطلبات:'),
        h('ul', null, run.errors.map((x) => h('li', null, x)))));
    }
    if (run.results) {
      box.appendChild(h('div', { class: 'note info' }, run.results.summary,
        h('div', { class: 'results' }, run.results.files.map((f) =>
          h('button', { class: 'btn gold', onclick: () => download(f.data, f.name, f.mime) }, '⬇ ' + f.label)))));
    }
    return box;
  }

  function extraFor(p, x, slotsCache) {
    const isGuide = x.action === 'guide';
    const key = isGuide ? 'guide' : 'main';
    if (!slotsCache[key]) {
      const ins = E.inspectTemplate(templateFor(p, x.action));
      slotsCache[key] = ins.ok ? E.computeMaxSlots(p, ins.tags) : {};
    }
    return { reasons: x.ev.reasons, maxSlot: slotsCache[key], tagIndex: slotsCache.idx || (slotsCache.idx = E.buildTagIndex(p)) };
  }

  function outName(p, x) {
    const pattern = p.settings.filePattern || '{التسلسل} - {الاسم}';
    const base = E.fileNameFor(pattern, p, x.r);
    return (x.action === 'guide' ? 'توجيه - ' : '') + (x.r && x.r.manual && state.run && state.run.table ? 'يدوي ' : '') + base + '.docx';
  }

  async function downloadOne(x) {
    const p = state.current;
    try {
      const data = E.renderOne(templateFor(p, x.action), p, x.r, extraFor(p, x, {}));
      download(data, outName(p, x), DOCX_MIME);
    } catch (e) {
      toast('تعذّر توليد الملف: ' + explain(e), 'err');
    }
  }

  function explain(e) {
    if (e && e.properties && e.properties.errors) return e.properties.errors.map((x) => (x.properties && x.properties.explanation) || x.message).join(' · ');
    return (e && e.message) || String(e);
  }

  async function generate() {
    const p = state.current;
    const run = state.run;
    const sel = run.rows.filter((x) => x.sel);
    const work = sel.filter((x) => x.action !== 'exclude');
    run.busy = true; run.progress = 0; run.results = null; run.errors = [];
    render();
    const mains = [], guides = [], cache = {};
    for (let i = 0; i < work.length; i++) {
      const x = work[i];
      try {
        const data = E.renderOne(templateFor(p, x.action), p, x.r, extraFor(p, x, cache));
        (x.action === 'guide' ? guides : mains).push({ x, data, name: outName(p, x) });
      } catch (e) {
        run.errors.push('الطلب ' + x.r.seq + ' (' + x.name + '): ' + explain(e));
      }
      run.progress = (i + 1) / Math.max(work.length, 1);
      const bar = document.getElementById('prog');
      if (bar) bar.style.width = Math.round(run.progress * 100) + '%';
      if (i % 3 === 2) await tick();
    }
    const mode = p.settings.outputMode || 'zip';
    const files = [];
    const base = safeName(p.name) + ' — ' + stamp();
    try {
      if ((mode === 'zip' || mode === 'both') && (mains.length || guides.length)) {
        const both = mains.length && guides.length;
        const list = mains.map((m) => ({ name: (both ? 'الطلبات/' : '') + m.name, data: m.data }))
          .concat(guides.map((g) => ({ name: (both ? 'خطابات التوجيه/' : '') + g.name, data: g.data })));
        await tick();
        files.push({ label: 'الملفات منفصلة (' + list.length + ') — ملف مضغوط', name: base + '.zip', data: E.makeZip(list), mime: 'application/zip' });
      }
      if (mode === 'merged' || mode === 'both') {
        await tick();
        if (mains.length) files.push({ label: 'الطلبات مجمّعة (' + mains.length + ') — وورد', name: base + ' — الطلبات المجمّعة.docx', data: E.mergeDocs(mains.map((m) => m.data)), mime: DOCX_MIME });
        if (guides.length) files.push({ label: 'خطابات التوجيه مجمّعة (' + guides.length + ') — وورد', name: base + ' — خطابات التوجيه المجمّعة.docx', data: E.mergeDocs(guides.map((g) => g.data)), mime: DOCX_MIME });
      }
    } catch (e) {
      run.errors.push('تعذّر تجميع الملفات: ' + explain(e));
    }
    // تعليم المنجز
    if (p.settings.markDone !== false) {
      const now = Date.now();
      p.processed = p.processed || {};
      mains.forEach((m) => { p.processed[m.x.fp] = { t: now, n: m.x.name, k: 'main' }; m.x.done = p.processed[m.x.fp]; m.x.sel = false; });
      guides.forEach((g) => { p.processed[g.x.fp] = { t: now, n: g.x.name, k: 'guide' }; g.x.done = p.processed[g.x.fp]; g.x.sel = false; });
      sel.filter((x) => x.action === 'exclude').forEach((x) => { p.processed[x.fp] = { t: now, n: x.name, k: 'x' }; x.done = p.processed[x.fp]; x.sel = false; });
      try { await S.put(p); } catch (e) { run.errors.push('تعذّر حفظ سجل المنجز: ' + e.message); }
      await refreshProfiles();
    }
    run.busy = false;
    run.results = files.length ? {
      summary: 'تمّ التوليد — الطلبات: ' + mains.length + (guides.length ? ' · خطابات التوجيه: ' + guides.length : '') + '. اضغط لتنزيل الملفات:',
      files,
    } : null;
    if (!files.length && !run.errors.length) toast('لم يُولَّد شيء (كل الطلبات المحددة مستبعَدة)');
    render();
    if (files.length) {
      download(files[0].data, files[0].name, files[0].mime);
      const rb = document.getElementById('results-box');
      if (rb) rb.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }

  /* ================================================================== *
   *  التصدير والاستيراد والحذف
   * ================================================================== */

  async function exportFlow(p) {
    const full = await S.get(p.id);
    download(E.exportProfile(full), safeName(full.name) + '.استبيان.json', 'application/json');
    toast('صُدّر الاستبيان — أرسل الملف لمن يريد استعماله');
  }

  async function importFlow(fileArg) {
    const file = fileArg instanceof File ? fileArg : await pickFile('.json');
    if (!file) return;
    let p;
    try { p = E.importProfile(await file.text()); } catch (e) { return toast(e.message, 'err'); }
    const existing = await S.get(p.id);
    if (existing) {
      const choice = await modal({
        title: 'هذا الاستبيان موجود عندك',
        body: 'يوجد استبيان محفوظ بالمعرّف نفسه: «' + existing.name + '».',
        buttons: [{ label: 'استبدله بالمستورَد', value: 'replace', kind: 'danger' }, { label: 'احفظه نسخة مستقلة', value: 'copy' }, { label: 'إلغاء', value: null, kind: 'ghost' }],
      });
      if (!choice) return;
      if (choice === 'copy') { p.id = newId(); p.name += ' (نسخة)'; }
      else { p.processed = existing.processed || {}; p.lastExcel = existing.lastExcel || null; }
    }
    p.settings = Object.assign({}, DEFAULT_SETTINGS, p.settings || {});
    await S.put(p);
    await refreshProfiles();
    toast('استُورد الاستبيان «' + p.name + '» — ارفع ملف الإكسل الخاص به');
    openRun(p.id);
  }

  async function deleteFlow(p) {
    const ok = await modal({
      title: 'حذف الاستبيان؟',
      body: 'سيُحذف «' + p.name + '» من هذا المتصفح مع قالبه وملف الإكسل المحفوظ وسجل المنجز. لا يمكن التراجع إلا إن كان عندك نسخة مُصدَّرة.',
      buttons: [{ label: 'تصدير نسخة أولًا', value: 'export' }, { label: 'نعم، احذفه', value: 'delete', kind: 'danger' }, { label: 'إلغاء', value: null, kind: 'ghost' }],
    });
    if (ok === 'export') return exportFlow(p);
    if (ok !== 'delete') return;
    await S.remove(p.id);
    await refreshProfiles();
    toast('حُذف الاستبيان');
    goHome('saved');
  }

  /* ================================================================== *
   *  السحب والإفلات على كامل الصفحة
   * ================================================================== */

  let dragDepth = 0;
  window.addEventListener('dragenter', (e) => {
    if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return;
    dragDepth++; document.body.classList.add('dragging');
  });
  window.addEventListener('dragleave', () => { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) document.body.classList.remove('dragging'); });
  window.addEventListener('dragover', (e) => { e.preventDefault(); });
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    dragDepth = 0; document.body.classList.remove('dragging');
    const f = e.dataTransfer && e.dataTransfer.files[0];
    if (!f) return;
    const n = f.name.toLowerCase();
    if (/\.(xlsx|xlsm|csv)$/.test(n)) {
      if (state.view === 'run') loadRunExcel(f);
      else if (state.view === 'editor') replaceDraftExcel(f);
      else routeExcel(f);
    } else if (/\.docx$/.test(n)) {
      if (state.view === 'editor') loadTemplateFile(f, 'main');
      else toast('لرفع قالب الوورد افتح «الإعدادات» أولًا', 'err');
    } else if (/\.json$/.test(n)) importFlow(f);
    else if (/\.xls$/.test(n)) toast('هذا ملف إكسل قديم (‎.xls‎). احفظه من إكسل بصيغة ‎.xlsx‎ ثم أعد رفعه.', 'err');
    else toast('نوع الملف غير مدعوم: ' + f.name, 'err');
  });

  /* ================================================================== *
   *  البدء
   * ================================================================== */

  async function start() {
    document.getElementById('ver').textContent = E.VERSION;
    try {
      await refreshProfiles();
    } catch (e) {
      $app.innerHTML = '';
      $app.appendChild(h('div', { class: 'card note err' }, 'تعذّر فتح الحفظ الداخلي في هذا المتصفح: ', e.message || String(e)));
      return;
    }
    state.tab = state.profiles.length ? 'saved' : 'new';
    render();
    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
      navigator.serviceWorker.register('sw.js').catch(() => {});
    }
  }

  window.addEventListener('error', (e) => { toast('حدث خطأ غير متوقع: ' + (e.message || ''), 'err'); });
  start();
})();
