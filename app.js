// Phòng kiểm toán KMF – 3D realtime trên SharePoint List (Microsoft Graph)
{
const C = window.KMF_CONFIG || {};
const DEMO = /demo/.test(location.search + location.hash) || !C.clientId || C.clientId.startsWith('DAN-');   // chưa cấu hình -> tự chạy bản thử
const STEPS = ['Đang lập', 'Chờ Senior', 'Chờ Manager', 'Chờ Partner', 'Hoàn thành'];
const STEP_ROLE = {'Đang lập': 'prep', 'Chờ Senior': 'senior', 'Chờ Manager': 'manager', 'Chờ Partner': 'partner', 'Hoàn thành': null};
const LV = {Partner: '#0b1f4b', Manager: '#2b5aa8', Senior: '#3f8f7a', Staff: '#b9832f', 'Trợ lý': '#b9832f'};
const LV_ORDER = ['Partner', 'Manager', 'Senior', 'Staff', 'Trợ lý'];
const PCOL = ['#1f6fd6', '#f08a24', '#2a9d8f', '#8e5bd6', '#c8486b', '#5b7b2a', '#0f8fb8', '#a0522d'];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
const $ = id => document.getElementById(id);
const todayISO = () => new Date(Date.now() + 7 * 3600e3).toISOString().slice(0, 10);
const isLate = w => w.due && w.step !== 'Hoàn thành' && w.due < todayISO();
const vnDate = s => s ? s.split('-').reverse().join('/') : '';
const toISO = v => {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);   // số ngày Excel
  const m = String(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return /^\d{4}-\d{2}-\d{2}/.test(String(v)) ? String(v).slice(0, 10) : '';
};
function holderOf(w, step) { const r = STEP_ROLE[step]; return r ? (w[r] || '') : ''; }
function nextStep(w) {   // bỏ qua bước không có người
  let i = STEPS.indexOf(w.step);
  while (++i < STEPS.length) { if (STEPS[i] === 'Hoàn thành' || holderOf(w, STEPS[i])) return STEPS[i]; }
  return null;
}
function toast(t) { const el = $('toast'); el.textContent = t; el.style.display = 'block'; clearTimeout(toast.h); toast.h = setTimeout(() => el.style.display = 'none', 5000); }

// ================= KHO DỮ LIỆU: SharePoint (Graph) =================
const SCOPES = ['User.Read', 'Sites.Manage.All', 'Files.ReadWrite.All'];
const LISTS = {
  KMF_HoSo: ['Proj', 'Client', 'WP', 'WPName', 'Part', 'Step', 'Holder', 'Prep', 'Senior', 'Manager', 'Partner', 'Due', 'Note'],
  KMF_NhanSu: ['Email', 'Level', 'Status'],
  KMF_LichSu: ['FromStep', 'ToStep', 'FromHolder', 'ToHolder', 'By', 'Note']
};
const F2K = {Proj: 'proj', Client: 'client', WP: 'wp', WPName: 'name', Part: 'part', Step: 'step', Holder: 'holder', Prep: 'prep', Senior: 'senior', Manager: 'manager', Partner: 'partner', Due: 'due', Note: 'note'};
const K2F = Object.fromEntries(Object.entries(F2K).map(([a, b]) => [b, a]));
const Graph = {
  pca: null, acct: null, site: null, lid: {},
  async init() {
    this.pca = new msal.PublicClientApplication({auth: {clientId: C.clientId, authority: `https://login.microsoftonline.com/${C.tenantId}`, redirectUri: location.origin + location.pathname},
      cache: {cacheLocation: 'localStorage'}});
    const r = await this.pca.handleRedirectPromise(); this.acct = (r && r.account) || this.pca.getAllAccounts()[0] || null;
    return !!this.acct;
  },
  inFrame: () => window.self !== window.top,
  async login() {
    if (this.inFrame()) { const r = await this.pca.loginPopup({scopes: SCOPES}); this.acct = r.account; return true; }
    await this.pca.loginRedirect({scopes: SCOPES});
  },
  async token() {
    try { return (await this.pca.acquireTokenSilent({scopes: SCOPES, account: this.acct})).accessToken; }
    catch (e) { if (this.inFrame()) return (await this.pca.acquireTokenPopup({scopes: SCOPES})).accessToken; await this.pca.acquireTokenRedirect({scopes: SCOPES}); }
  },
  async g(path, opt = {}) {
    const url = path.startsWith('http') ? path : 'https://graph.microsoft.com/v1.0' + path;
    const res = await fetch(url, {...opt, headers: {Authorization: 'Bearer ' + await this.token(), 'Content-Type': 'application/json', Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly', ...(opt.headers || {})}});
    if (res.status === 204) return null;
    const j = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error((j.error && j.error.message) || res.statusText);
    return j;
  },
  async all(path) { let out = [], next = path; while (next) { const j = await this.g(next); out = out.concat(j.value || []); next = j['@odata.nextLink']; } return out; },
  async open() {
    this.site = (await this.g(`/sites/${C.siteHost}:${C.sitePath}`)).id;
    const ls = await this.all(`/sites/${this.site}/lists?$select=id,displayName`);
    for (const n of Object.keys(LISTS)) { const l = ls.find(x => x.displayName === n); if (l) this.lid[n] = l.id; }
    return Object.keys(LISTS).every(n => this.lid[n]);
  },
  async setup() {
    for (const [n, cols] of Object.entries(LISTS)) {
      if (this.lid[n]) continue;
      const l = await this.g(`/sites/${this.site}/lists`, {method: 'POST', body: JSON.stringify({displayName: n, list: {template: 'genericList'},
        columns: cols.map(c => c === 'Note' || c === 'WPName' ? {name: c, text: {allowMultipleLines: true}} : {name: c, text: {}})})});
      this.lid[n] = l.id;
    }
  },
  me() { return {email: (this.acct.username || '').toLowerCase(), name: this.acct.name || ''}; },
  async load() {
    const items = (await this.all(`/sites/${this.site}/lists/${this.lid.KMF_HoSo}/items?$expand=fields&$top=999`)).map(x => {
      const w = {id: x.id, modified: x.lastModifiedDateTime, by: (x.lastModifiedBy && x.lastModifiedBy.user && x.lastModifiedBy.user.displayName) || ''};
      for (const [f, k] of Object.entries(F2K)) w[k] = x.fields[f] || '';
      return w;
    });
    const staff = (await this.all(`/sites/${this.site}/lists/${this.lid.KMF_NhanSu}/items?$expand=fields&$top=999`))
      .map(x => ({id: x.id, name: x.fields.Title || '', email: (x.fields.Email || '').toLowerCase(), level: x.fields.Level || '', status: x.fields.Status || ''}));
    return {items, staff};
  },
  async update(w, ch) {
    const f = {}; for (const [k, v] of Object.entries(ch)) f[K2F[k]] = v;
    await this.g(`/sites/${this.site}/lists/${this.lid.KMF_HoSo}/items/${w.id}/fields`, {method: 'PATCH', body: JSON.stringify(f)});
  },
  async create(w) {
    const f = {Title: `${w.proj}|${w.wp}`}; for (const [k, F] of Object.entries(K2F)) if (w[k] != null) f[F] = String(w[k]);
    await this.g(`/sites/${this.site}/lists/${this.lid.KMF_HoSo}/items`, {method: 'POST', body: JSON.stringify({fields: f})});
  },
  async hist(rec) {
    await this.g(`/sites/${this.site}/lists/${this.lid.KMF_LichSu}/items`, {method: 'POST', body: JSON.stringify({fields: {Title: rec.key, FromStep: rec.fromStep, ToStep: rec.toStep, FromHolder: rec.fromHolder, ToHolder: rec.toHolder, By: rec.by, Note: rec.note || ''}})});
  },
  async history(key) {
    const rows = await this.all(`/sites/${this.site}/lists/${this.lid.KMF_LichSu}/items?$expand=fields&$filter=fields/Title eq '${key.replace(/'/g, "''")}'&$top=50`);
    return rows.map(x => ({at: x.createdDateTime, ...x.fields})).sort((a, b) => b.at.localeCompare(a.at));
  },
  async upsertStaff(s) {
    const f = {Title: s.name, Email: s.email, Level: s.level, Status: s.status || 'Đang làm việc'};
    if (s.id) await this.g(`/sites/${this.site}/lists/${this.lid.KMF_NhanSu}/items/${s.id}/fields`, {method: 'PATCH', body: JSON.stringify(f)});
    else await this.g(`/sites/${this.site}/lists/${this.lid.KMF_NhanSu}/items`, {method: 'POST', body: JSON.stringify({fields: f})});
  },
  async sheet(url, rel) {   // link SharePoint -> giá trị sheet đầu tiên
    const enc = 'u!' + btoa(unescape(encodeURIComponent(url.trim()))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
    let it = await this.g(`/shares/${enc}/driveItem`);
    const d = it.parentReference.driveId;
    if (rel) it = await this.g(`/drives/${d}/items/${it.id}:/${rel.split('/').map(encodeURIComponent).join('/')}`);
    const ws = (await this.g(`/drives/${d}/items/${it.id}/workbook/worksheets`)).value[0];
    return (await this.g(`/drives/${d}/items/${it.id}/workbook/worksheets/${encodeURIComponent(ws.id)}/usedRange(valuesOnly=true)`)).values;
  }
};

// ================= KHO DỮ LIỆU: bản thử (?demo) =================
const Demo = {
  data: null, hs: [],
  async init() { return true; }, async open() { return true; },
  me() { return {email: 'duy.vk@kmf-audit.com', name: 'Vũ Khánh Duy'}; },
  async load() {
    if (!this.data) {
      const D = window.KMF_DEMO || {projects: [], staff: []}; let n = 0;
      const items = D.projects.flatMap(p => p.wps.map(w => ({id: String(++n), proj: p.code, client: p.client, wp: w.wp, name: w.name, part: w.part, step: w.step, holder: w.holder,
        prep: w.holder, senior: (p.team || {}).Senior || '', manager: (p.team || {}).Manager || '', partner: (p.team || {}).Partner || '', due: w.due || '', note: '', modified: '', by: ''})));
      this.data = {items, staff: D.staff.map((s, i) => ({id: 's' + i, ...s, status: 'Đang làm việc'}))};
      setInterval(() => this.simulate(), 20000);
    }
    return JSON.parse(JSON.stringify(this.data));
  },
  simulate() {   // giả lập đồng nghiệp chuyển hồ sơ để xem hiệu ứng realtime
    const c = this.data.items.filter(w => w.holder && nextStep(w) && w.step !== 'Chờ Partner');
    if (!c.length) return; const w = c[Math.floor(Math.random() * c.length)], s = nextStep(w), from = w.holder;
    w.step = s; w.holder = holderOf(w, s); w.by = from || 'Đồng nghiệp'; w.modified = new Date().toISOString();
    this.hs.push({key: `${w.proj}|${w.wp}`, at: w.modified, FromStep: '', ToStep: s, FromHolder: from, ToHolder: w.holder, By: w.by});
  },
  async update(w, ch) { Object.assign(this.data.items.find(x => x.id === w.id), ch, {by: this.me().name, modified: new Date().toISOString()}); },
  async hist(r) { this.hs.push({key: r.key, at: new Date().toISOString(), FromStep: r.fromStep, ToStep: r.toStep, FromHolder: r.fromHolder, ToHolder: r.toHolder, By: r.by, Note: r.note}); },
  async history(key) { return this.hs.filter(h => h.key === key).reverse(); },
  async setup() {}, async create() {}, async upsertStaff() {}, async sheet() { throw new Error('Bản thử không nạp được dự án.'); }
};
const S = DEMO ? Demo : Graph;

// ================= TRẠNG THÁI =================
let ITEMS = [], STAFF = [], ME = {}, MYNAME = '', IS_ADMIN = false, filter = 'ALL', sel = null, selWp = null, prevHolder = new Map(), first = true;
const projects = () => [...new Set(ITEMS.map(w => w.proj))].sort();
const pcolor = p => PCOL[projects().indexOf(p) % PCOL.length];
const visible = () => ITEMS.filter(w => filter === 'ALL' || w.proj === filter);
function canAct(w) { return IS_ADMIN || (MYNAME && [w.holder, w.senior, w.manager, w.partner].includes(MYNAME)); }

// ================= 3D – văn phòng sa bàn (isometric) =================
const stage = $('stage'), tip = $('tip'), tagLayer = $('tags');
const renderer = new THREE.WebGLRenderer({antialias: true, alpha: true});
renderer.setPixelRatio(Math.min(devicePixelRatio, 2)); renderer.shadowMap.enabled = true; renderer.shadowMap.type = THREE.PCFSoftShadowMap;
renderer.outputEncoding = THREE.sRGBEncoding; renderer.toneMapping = THREE.ACESFilmicToneMapping; renderer.toneMappingExposure = .95;
stage.prepend(renderer.domElement);
const scene = new THREE.Scene();
const cam = new THREE.PerspectiveCamera(30, 1, .5, 500);
const HOME = {pos: new THREE.Vector3(44, 46, 52), tgt: new THREE.Vector3(1, 0, -2)};
cam.position.copy(HOME.pos);
const ctl = new THREE.OrbitControls(cam, renderer.domElement); ctl.target.copy(HOME.tgt); ctl.enableDamping = true; ctl.maxPolarAngle = 1.25; ctl.minDistance = 18; ctl.maxDistance = 140; ctl.autoRotateSpeed = .6;
scene.add(new THREE.HemisphereLight(0xffffff, 0xb8bec8, .55));
const sun = new THREE.DirectionalLight(0xfff1de, 1.0); sun.position.set(30, 60, 25); sun.castShadow = true; sun.shadow.mapSize.set(2048, 2048); sun.shadow.bias = -.0005;
Object.assign(sun.shadow.camera, {left: -50, right: 50, top: 50, bottom: -50, far: 160}); scene.add(sun);
const fillL = new THREE.DirectionalLight(0xdfe8ff, .35); fillL.position.set(-30, 25, -20); scene.add(fillL);
const M = {};
const lin = c => new THREE.Color(c).convertSRGBToLinear();
const mat = (c, o = {}) => { const k = c + JSON.stringify(o); if (o.emissive != null) o = {...o, emissive: lin(o.emissive)}; return M[k] ||= new THREE.MeshStandardMaterial({color: lin(c), roughness: .75, metalness: 0, ...o}); };
function box(w, h, d, color, x = 0, y = 0, z = 0, o) { const m = new THREE.Mesh(new THREE.BoxGeometry(w, h, d), typeof color === 'object' ? color : mat(color, o)); m.position.set(x, y, z); m.castShadow = m.receiveShadow = true; return m; }
function cyl(rt, rb, h, color, x = 0, y = 0, z = 0, seg = 16) { const m = new THREE.Mesh(new THREE.CylinderGeometry(rt, rb, h, seg), mat(color)); m.position.set(x, y, z); m.castShadow = m.receiveShadow = true; return m; }
function floorText(text, w, h, x, z, rot = 0, bg = '#23262d', fg = '#ffffff') {
  const c = document.createElement('canvas'); c.width = 1024; c.height = Math.round(1024 * h / w); const g = c.getContext('2d');
  g.fillStyle = bg; g.fillRect(0, 0, c.width, c.height);
  g.fillStyle = '#f5c542'; g.fillRect(0, 0, 18, c.height);
  g.fillStyle = fg; g.font = `700 ${Math.round(c.height * .42)}px "Be Vietnam Pro",sans-serif`; g.textBaseline = 'middle'; g.fillText(text, 60, c.height / 2, c.width - 90);
  const t = new THREE.CanvasTexture(c); t.encoding = THREE.sRGBEncoding; t.anisotropy = 8;
  const m = new THREE.Mesh(new THREE.PlaneGeometry(w, h), new THREE.MeshStandardMaterial({map: t, roughness: .9})); m.rotation.x = -Math.PI / 2; m.rotation.z = rot; m.position.set(x, .17, z); m.receiveShadow = true; return m;
}
function tree(x, z, s = 1) {
  const g = new THREE.Group(); g.position.set(x, 0, z); g.scale.setScalar(s);
  g.add(cyl(.18, .24, 1.6, 0x7a5a3a, 0, .8, 0, 8));
  const leaf = new THREE.Mesh(new THREE.IcosahedronGeometry(1.25, 1), mat(0x4f9a5a, {flatShading: true})); leaf.position.y = 2.4; leaf.castShadow = true; g.add(leaf);
  const leaf2 = new THREE.Mesh(new THREE.IcosahedronGeometry(.85, 1), mat(0x5fae68, {flatShading: true})); leaf2.position.set(.45, 3.1, .2); leaf2.castShadow = true; g.add(leaf2);
  return g;
}
function plant(x, z) { const g = new THREE.Group(); g.position.set(x, 0, z); g.add(cyl(.45, .35, .8, 0xf2f2f2, 0, .4, 0)); const l = new THREE.Mesh(new THREE.IcosahedronGeometry(.7, 0), mat(0x3f8f4e, {flatShading: true})); l.position.y = 1.3; l.castShadow = true; g.add(l); return g; }
const glassMat = new THREE.MeshStandardMaterial({color: lin(0x9fd3ea), transparent: true, opacity: .3, roughness: .1, metalness: .1, depthWrite: false});
function glass(w, h, x, z, rotY = 0) {
  const g = new THREE.Group(); g.position.set(x, 0, z); g.rotation.y = rotY;
  const p = new THREE.Mesh(new THREE.BoxGeometry(w, h, .08), glassMat); p.position.y = h / 2; g.add(p);
  g.add(box(w, .08, .14, 0x2b2f36, 0, h, 0), box(.1, h, .14, 0x2b2f36, -w / 2, h / 2, 0), box(.1, h, .14, 0x2b2f36, w / 2, h / 2, 0));
  return g;
}
const SKIN = [0xf1c9a5, 0xe9b98f, 0xdba47a, 0xf5d3b3], HAIR = [0x1f1f22, 0x3a2a20, 0x2b2b2b, 0x5a3b28];
const SHIRT = {Partner: 0x1f2f5c, Manager: 0x2f6fd0, Senior: 0x2f9a86, Staff: 0xe08a3c, 'Trợ lý': 0xe08a3c};
function person(level, i) {   // nhân vật khối kiểu sa bàn, ngồi quay mặt về -z
  const g = new THREE.Group(), shirt = SHIRT[level] || 0x888888, skin = SKIN[i % 4], hair = HAIR[i % 4];
  g.add(box(.55, .35, 1.0, 0x30343c, -.28, 1.25, -.35), box(.55, .35, 1.0, 0x30343c, .28, 1.25, -.35));   // đùi
  g.add(box(.5, 1.1, .45, 0x30343c, -.28, .55, -.85), box(.5, 1.1, .45, 0x30343c, .28, .55, -.85));   // cẳng chân
  g.add(box(1.25, 1.35, .75, shirt, 0, 2.05, .05));   // thân
  g.add(box(.36, 1.0, .36, shirt, -.82, 2.1, -.2), box(.36, 1.0, .36, shirt, .82, 2.1, -.2));   // tay
  g.add(box(.34, .34, .8, skin, -.82, 1.62, -.65), box(.34, .34, .8, skin, .82, 1.62, -.65));   // cẳng tay lên bàn
  const head = box(.9, .9, .9, skin, 0, 3.2, 0); g.add(head);
  g.add(box(.96, .32, .96, hair, 0, 3.6, .02), box(.96, .6, .2, hair, 0, 3.35, .42));
  g.add(box(.12, .12, .04, 0x1b1b1b, -.2, 3.25, -.46), box(.12, .12, .04, 0x1b1b1b, .2, 3.25, -.46));   // mắt
  g.userData.head = head; return g;
}
function chair() { const g = new THREE.Group(); g.add(box(1.3, .2, 1.3, 0x2c3038, 0, 1.05, 0), box(1.3, 1.5, .2, 0x2c3038, 0, 1.85, .62), cyl(.08, .08, .9, 0x8a8f98, 0, .5, 0, 8), cyl(.6, .6, .08, 0x5a5f68, 0, .06, 0, 10)); return g; }
function laptop() { const g = new THREE.Group(); g.add(box(1.3, .06, .9, 0xc9ccd2, 0, 0, 0)); const s = box(1.3, .85, .05, 0xc9ccd2, 0, .43, -.45); s.rotation.x = -.18; g.add(s);
  const scr = new THREE.Mesh(new THREE.PlaneGeometry(1.15, .7), new THREE.MeshBasicMaterial({color: lin(0x6fb4f0)})); scr.position.set(0, .44, -.41); scr.rotation.x = -.18; g.add(scr); return g; }
function desk(color = 0xf4f1ea) {
  const g = new THREE.Group(); const top = box(4.2, .16, 2.3, color, 0, 1.55, 0); g.add(top);
  g.add(box(.12, 1.5, 2.1, 0x9aa0a8, -2, .75, 0), box(.12, 1.5, 2.1, 0x9aa0a8, 2, .75, 0));
  const lp = laptop(); lp.position.set(-.3, 1.66, -.15); g.add(lp);
  g.add(box(.5, .02, .4, 0xffffff, 1.2, 1.64, .3), cyl(.12, .1, .3, 0xffffff, 1.4, 1.78, -.6, 10));   // giấy, ly
  g.userData.top = top; return g;
}

const room = new THREE.Group(), stacks = new THREE.Group(), flights = new THREE.Group(); scene.add(room, stacks, flights);
let desks = [], deskById = {}, pick = [], TAGS = {}, showTags = true;
function people() {   // danh bạ; thiếu thì suy ra từ vai trò trong hồ sơ
  const act = STAFF.filter(s => s.name && s.status !== 'Đã nghỉ việc');
  const have = new Set(act.map(s => s.name)), extra = [];
  const add = (n, l) => { if (n && !have.has(n)) { have.add(n); extra.push({name: n, level: l}); } };
  ITEMS.forEach(w => { add(w.partner, 'Partner'); add(w.manager, 'Manager'); add(w.senior, 'Senior'); add(w.prep, 'Staff'); add(w.holder, 'Staff'); });
  return act.concat(extra).sort((a, b) => LV_ORDER.indexOf(a.level) - LV_ORDER.indexOf(b.level));
}
const ZONES = [{lv: ['Partner'], name: 'KHU PARTNER'}, {lv: ['Manager', 'Senior'], name: 'KHU SOÁT XÉT'}, {lv: ['Staff', 'Trợ lý'], name: 'KHU THỰC HIỆN'}];
function buildRoom() {
  room.clear(); pick = []; desks = []; tagLayer.innerHTML = ''; TAGS = {};
  const ps = people(); let z = -15, rows = [];
  ZONES.forEach(zn => {
    const list = ps.filter(p => zn.lv.includes(p.level)); if (!list.length) return;
    rows.push({name: zn.name, z});
    for (let i = 0; i < list.length; i++) { if (i && i % 6 === 0) z += 8; desks.push({...list[i], id: 'p:' + list[i].name, x: -24 + (i % 6) * 6.6, z, idx: desks.length}); }
    rows[rows.length - 1].n = Math.min(list.length, 6);
    z += 10;
  });
  const zEnd = Math.max(z - 4, 12), W = 64, D = zEnd + 22, cz = -21 + D / 2;
  // sàn sa bàn
  const base = box(W, 1.2, D, 0xc9ccd1, 4, -.6, cz); base.castShadow = false; room.add(base);
  const fl = box(W - 1, .1, D - 1, 0xe9e9e6, 4, .05, cz); fl.castShadow = false; room.add(fl);
  const wood = box(18, .12, D - 6, 0xd9bd94, 26, .07, cz + 1); wood.castShadow = false; room.add(wood);   // sàn gỗ khu chung
  // tường kính phía sau + tường trái
  for (let x = -26; x <= 34; x += 6) room.add(glass(6, 4.2, x, -20.3));
  const wl = box(.6, 4.4, D - 1, 0xffffff, -27.8, 2.2, cz); room.add(wl);
  for (let zz = -17; zz < zEnd; zz += 8) room.add(box(.1, 2.2, 3.2, 0x9cc9e6, -27.45, 2.6, zz + 3, {emissive: 0x5a8fb0, emissiveIntensity: .15}));
  // vạch tên khu trên sàn + vách kính giữa các hàng
  rows.forEach(r => { room.add(floorText(r.name, 8, 1.5, -21.5, r.z + 4.4)); if (r.name === 'KHU PARTNER') return; const len = r.n * 6.6; room.add(glass(len, 1.1, -24 - 3.3 + len / 2, r.z - 2.4)); });
  // khu chung bên phải: lễ tân + logo, phòng họp, tủ hồ sơ, kho Final
  const rec = new THREE.Group(); rec.position.set(24, 0, -15); rec.add(box(7, 1.6, 1.6, 0xffffff, 0, .8, 0), box(7.2, .12, 1.8, 0x2b2f36, 0, 1.66, 0)); room.add(rec);
  const board = box(6, 3.2, .3, 0x23262d, 24, 2.2, -19.4); room.add(board);
  const lg = floorText('KMF AUDIT', 5.4, 1.3, 24, -19.2, 0, '#23262d'); lg.rotation.set(0, 0, 0); lg.position.set(24, 2.4, -19.2); room.add(lg);
  const mt = new THREE.Group(); mt.position.set(25, 0, -5); room.add(mt);
  mt.add(glass(10, 3.4, 0, -4), glass(10, 3.4, 0, 4), glass(8, 3.4, -5, 0, Math.PI / 2));
  mt.add(box(6, .16, 2.6, 0xf4f1ea, 0, 1.55, 0), box(.2, 1.5, .2, 0x9aa0a8, 0, .75, 0));
  [-2, 0, 2].forEach(x => { const c1 = chair(); c1.position.set(x, 0, 2); mt.add(c1); const c2 = chair(); c2.position.set(x, 0, -2); c2.rotation.y = Math.PI; mt.add(c2); });
  room.add(floorText('PHÒNG HỌP', 6, 1.2, 25, 0.2));
  // tủ chưa phân công & kho Final
  const zc = Math.min(8, zEnd - 4);
  desks.push({id: 'none', name: 'Chưa phân công', level: '', x: 20, z: zc, special: true}, {id: 'final', name: 'Lưu trữ Final', level: '', x: 29, z: zc, special: true});
  desks.forEach(d => {
    const g = new THREE.Group(); g.position.set(d.x, 0, d.z); room.add(g); d.g = g;
    d.face = d.level === 'Partner' ? -1 : 1; if (d.face < 0) g.rotation.y = Math.PI;   // Partner ngồi quay lưng vào tường, mặt hướng ra phòng
    if (d.id === 'none') {
      g.add(box(3.4, 3.2, 1.8, 0x8d939c, 0, 1.6, 0)); [0.7, 1.6, 2.5].forEach(y => g.add(box(3.0, .7, .05, 0xa7adb6, 0, y, .93)));
      d.top = 3.2; d.stackX = -.8; g.children.forEach(o => { o.userData.desk = d; pick.push(o); });
      room.add(floorText('CHƯA PHÂN CÔNG', 7, 1.2, d.x, d.z + 2.6));
    } else if (d.id === 'final') {
      for (let k = 0; k < 3; k++) { g.add(box(4.4, .12, 1.6, 0x6b4f37, 0, .6 + k * 1.1, 0)); for (let j = 0; j < 5; j++) g.add(box(.7, .8, 1.1, [0x3f8f5a, 0x5aa36f, 0x2f7a4a][(j + k) % 3], -1.7 + j * .85, 1.06 + k * 1.1, 0)); }
      g.add(box(.12, 3.6, 1.6, 0x6b4f37, -2.2, 1.8, 0), box(.12, 3.6, 1.6, 0x6b4f37, 2.2, 1.8, 0));
      d.top = 3.9; d.stackX = -1; g.children.forEach(o => { o.userData.desk = d; pick.push(o); });
      room.add(floorText('LƯU TRỮ FINAL', 7, 1.2, d.x, d.z + 2.6));
    } else {
      const dk = desk(); g.add(dk); dk.userData.top.userData.desk = d; pick.push(dk.userData.top);
      const ch = chair(); ch.position.set(0, 0, 1.9); g.add(ch);
      const p = person(d.level, d.idx); p.position.set(0, .15, 2.05); g.add(p); d.person = p;
      p.traverse(o => { if (o.isMesh) { o.userData.desk = d; pick.push(o); } });
      g.add(glass(2.4, 1.2, 2.2, 0, Math.PI / 2));   // vách ngăn nhỏ
      d.top = 1.63; d.stackX = 1.1;
    }
    const el = document.createElement('div'); el.className = 'tag'; el.dataset.id = d.id; tagLayer.appendChild(el); TAGS[d.id] = el;
    el.onclick = () => { sel = d.id; selWp = null; openSide(); drawStacks(); renderSide(true); };
  });
  deskById = Object.fromEntries(desks.map(d => [d.id, d]));
  // cây xanh quanh sa bàn
  for (let x = -24; x <= 30; x += 9) room.add(tree(x, zEnd + 5, .9));
  [[-25.5, -18], [33, -18], [33, zEnd + 2], [14, -18]].forEach(([x, zz]) => room.add(tree(x, zz, .8)));
  [[16, -12], [16, 3], [-26.5, zEnd]].forEach(([x, zz]) => room.add(plant(x, zz)));
  const sofa = new THREE.Group(); sofa.position.set(28, 0, zEnd + 1); sofa.add(box(4.5, .9, 1.6, 0x4a5568, 0, .45, 0), box(4.5, 1.1, .4, 0x4a5568, 0, 1.2, .7)); room.add(sofa);
  room.add(box(2, .5, 1, 0xd8b98f, 28, .25, zEnd - 1.2));
  // góc pantry ở giữa: bàn tròn, ghế đôn, thảm
  const rug = new THREE.Mesh(new THREE.CircleGeometry(5, 40), mat(0xdfe6ee)); rug.rotation.x = -Math.PI / 2; rug.position.set(-2, .12, zEnd - 1); rug.receiveShadow = true; room.add(rug);
  [[-4.5, zEnd - 2], [0.5, zEnd - .5]].forEach(([x, zz]) => { room.add(cyl(1.1, 1.1, .12, 0xffffff, x, 1.3, zz, 24), cyl(.1, .1, 1.25, 0x8a8f98, x, .65, zz, 8));
    [[1.6, 0], [-1.6, 0], [0, 1.6], [0, -1.6]].forEach(([dx, dz], k) => room.add(cyl(.45, .45, .8, [0xf5c542, 0xe86f5a, 0x5aa3e8, 0x6cc07a][k], x + dx, .4, zz + dz, 16))); });
  room.add(floorText('PANTRY', 4.5, 1.1, -2, zEnd + 2.6));
  HOME.tgt.set(3, 0, cz); const sc = Math.max(W, D * 1.4) / 64; HOME.pos.set(3 + 40 * sc, 46 * sc, cz + 50 * sc);
  if (first) { ctl.target.copy(HOME.tgt); cam.position.copy(HOME.pos); }
}
const deskOf = w => w.step === 'Hoàn thành' ? deskById.final : (deskById['p:' + w.holder] || deskById.none);
function byDesk() { const m = new Map(desks.map(d => [d.id, []])); visible().forEach(w => m.get(deskOf(w).id).push(w)); return m; }
function drawStacks() {
  stacks.clear(); pick = pick.filter(o => !o.userData.stack);
  const m = byDesk(), projs = projects();
  desks.forEach(d => {
    const items = m.get(d.id), byP = {}; items.forEach(w => (byP[w.proj] ||= []).push(w));
    Object.keys(byP).sort((a, b) => projs.indexOf(a) - projs.indexOf(b)).forEach((p, k) => {
      const ws = byP[p], n = Math.min(ws.length, 14), late = ws.some(isLate), col = lin(pcolor(p));
      for (let j = 0; j < n; j++) {   // từng tập bìa xếp chồng, lệch nhẹ
        const f = box(.95, .1, 1.25, new THREE.MeshStandardMaterial({color: j % 2 ? col.clone().multiplyScalar(.85) : col, roughness: .6, emissive: late ? 0xd93a2b : 0, emissiveIntensity: late ? .45 : 0}),
          d.special ? d.x + d.stackX + k * 1.05 + Math.sin(j * 2.1) * .05 : d.x + d.face * (d.stackX - k * 1.05) + Math.sin(j * 2.1) * .05, d.top + .06 + j * .11, d.z - .2 * (d.face || 1) + Math.cos(j * 1.7) * .05);
        f.rotation.y = Math.sin(j * 3.3) * .08; f.userData = {desk: d, late, stack: true}; stacks.add(f); pick.push(f);
      }
    });
    const late = items.filter(isLate).length, me = d.name === MYNAME, el = TAGS[d.id];
    const groups = Object.keys(byP).map(p => `<i class="dot" style="background:${pcolor(p)}"></i>${esc(p)} ${byP[p].length}`).join(' · ');
    el.className = 'tag' + (sel === d.id ? ' on' : '') + (me ? ' me' : '') + (d.special ? ' sp' : '');
    el.innerHTML = (items.length ? `<div class="bub${late ? ' late' : ''}">${groups}${late ? ` · <b>${late} quá hạn</b>` : ''}</div>` : '') +
      `<div class="nm"><i class="lv" style="background:${d.special ? '#9aa0a8' : '#' + (SHIRT[d.level] || 0x888888).toString(16).padStart(6, '0')}"></i>${esc(d.name)}${me ? ' (bạn)' : ''}${d.level ? `<small>${esc(d.level)}</small>` : ''}</div>`;
  });
}
let FL = [];
function fly(w, fromD, toD) {
  const b = box(.95, .14, 1.25, new THREE.MeshStandardMaterial({color: lin(pcolor(w.proj)), emissive: 0xffffff, emissiveIntensity: .25}));
  const a = new THREE.Vector3(fromD.x, fromD.top + .4, fromD.z), z = new THREE.Vector3(toD.x, toD.top + .4, toD.z);
  b.position.copy(a); flights.add(b); FL.push({b, a, z, t0: performance.now()});
}

// ================= BẢNG BÊN =================
const side = $('side');
async function renderSide(force) {
  if (!force && side.contains(document.activeElement) && /INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) return;
  const m = byDesk(), vis = visible();
  const late = vis.filter(isLate).length, done = vis.filter(w => w.step === 'Hoàn thành').length;
  if (selWp) { const w = ITEMS.find(x => x.id === selWp); if (w) { if (!force && cardSig === w.modified + w.step) return; return wpCard(w); } selWp = null; }
  if (sel && deskById[sel]) {
    const d = deskById[sel], items = m.get(sel).sort((a, b) => (a.due || '').localeCompare(b.due || '') || a.wp.localeCompare(b.wp));
    side.innerHTML = `<button class="link" id="back">← Tất cả nhân sự</button><h2>${esc(d.name)}</h2><div class="sub">${esc(d.level || '')} · ${items.length} hồ sơ</div>
      <table><thead><tr><th>Dự án</th><th>WP</th><th>Bước</th><th class="n">Hạn</th></tr></thead><tbody>${items.map(w => `<tr class="row" data-w="${esc(w.id)}" tabindex="0"><td><span class="dot" style="background:${pcolor(w.proj)}"></span> ${esc(w.proj)}</td><td title="${esc(w.name)}">${esc(w.wp)}</td><td><span class="pill">${esc(w.step)}</span></td><td class="n ${isLate(w) ? 'late' : ''}">${esc(vnDate(w.due))}</td></tr>`).join('') || '<tr><td colspan="4">Không có hồ sơ.</td></tr>'}</tbody></table>`;
    $('back').onclick = () => { sel = null; drawStacks(); renderSide(true); };
    side.querySelectorAll('tr.row').forEach(tr => { const go = () => { selWp = tr.dataset.w; renderSide(true); }; tr.onclick = go; tr.onkeydown = e => e.key === 'Enter' && go(); });
    return;
  }
  const rows = desks.map(d => ({d, n: m.get(d.id).length, l: m.get(d.id).filter(isLate).length})).filter(r => r.n).sort((a, b) => b.l - a.l || b.n - a.n);
  const mine = deskById['p:' + MYNAME];
  const pj = filter === 'ALL' ? projects() : [filter];
  side.innerHTML = `<button class="x" id="sx" aria-label="Đóng">×</button><h2>${filter === 'ALL' ? 'Tất cả dự án' : esc(filter)}</h2><div class="sub">${pj.map(p => { const w = ITEMS.find(x => x.proj === p); return `<span class="dot" style="background:${pcolor(p)}"></span> ${esc(p)} – ${esc(w ? w.client : '')}`; }).join('<br>')}</div>
    <div class="kpis"><div class="kpi"><b>${vis.length}</b><span>hồ sơ</span></div><div class="kpi"><b class="${late ? 'late' : ''}">${late}</b><span>quá hạn</span></div><div class="kpi"><b>${done}</b><span>hoàn thành</span></div></div>
    ${mine ? `<button id="myDesk">Bàn của tôi (${m.get(mine.id).length} hồ sơ)</button>` : ''}
    <table style="margin-top:10px"><thead><tr><th>Người đang giữ</th><th>Cấp</th><th class="n">Hồ sơ</th><th class="n">Quá hạn</th></tr></thead><tbody>${rows.map(r => `<tr class="row" data-id="${esc(r.d.id)}" tabindex="0"><td>${esc(r.d.name)}</td><td>${esc(r.d.level || '')}</td><td class="n">${r.n}</td><td class="n ${r.l ? 'late' : ''}">${r.l}</td></tr>`).join('') || '<tr><td colspan="4">Chưa có hồ sơ. Admin bấm "Nạp dự án".</td></tr>'}</tbody></table>
    <p class="note">${DEMO ? 'Bản thử: dữ liệu CME, BAP lúc 25/11/2026; cứ 20 giây có một hồ sơ được giả lập chuyển bước để xem hiệu ứng. Thao tác ở bản thử không lưu.' : 'Dữ liệu tự làm mới mỗi ' + (C.refreshSeconds || 8) + ' giây từ SharePoint.'}</p>`;
  $('sx').onclick = closeSide;
  if (mine) $('myDesk').onclick = () => { sel = mine.id; drawStacks(); renderSide(true); };
  side.querySelectorAll('tr.row').forEach(tr => { const go = () => { sel = tr.dataset.id; drawStacks(); renderSide(true); }; tr.onclick = go; tr.onkeydown = e => e.key === 'Enter' && go(); });
}
let cardSig = '';
async function wpCard(w) {
  cardSig = w.modified + w.step;
  const nx = nextStep(w), ok = canAct(w), key = `${w.proj}|${w.wp}`;
  side.innerHTML = `<button class="link" id="back">← ${esc(deskOf(w).name)}</button>
    <div class="wpcard"><div class="sub"><span class="dot" style="background:${pcolor(w.proj)}"></span> ${esc(w.proj)} – ${esc(w.client)}</div>
    <h2>${esc(w.wp)}</h2><div>${esc(w.name)}</div>
    <table style="margin-top:8px"><tbody>
      <tr><th>Bước hiện tại</th><td><span class="pill">${esc(w.step)}</span></td></tr>
      <tr><th>Đang giữ</th><td>${esc(w.holder || (w.step === 'Hoàn thành' ? 'Final' : 'Chưa phân công'))}</td></tr>
      <tr><th>Người lập · Senior</th><td>${esc(w.prep || '–')} · ${esc(w.senior || '–')}</td></tr>
      <tr><th>Manager · Partner</th><td>${esc(w.manager || '–')} · ${esc(w.partner || '–')}</td></tr>
      <tr><th>Cập nhật</th><td>${esc(w.by || '')} ${w.modified ? esc(new Date(w.modified).toLocaleString('vi-VN')) : ''}</td></tr>
    </tbody></table>
    <label for="due">Hạn bước này</label><input type="date" id="due" value="${esc(w.due)}" ${ok ? '' : 'disabled'}>
    <label for="note">Ghi chú / lý do trả lại</label><textarea id="note" rows="3" ${ok ? '' : 'disabled'}>${esc(w.note)}</textarea>
    <div class="acts">
      ${nx ? `<button class="primary" id="go" ${ok ? '' : 'disabled'}>Chuyển → ${esc(nx)}${holderOf(w, nx) ? ' (' + esc(holderOf(w, nx)) + ')' : ''}</button>` : ''}
      ${w.step !== 'Đang lập' ? `<button class="warn" id="ret" ${ok ? '' : 'disabled'}>Trả lại người lập</button>` : ''}
      <button id="save" ${ok ? '' : 'disabled'}>Lưu hạn, ghi chú</button>
    </div>${ok ? '' : '<p class="note">Chỉ người đang giữ, Senior, Manager, Partner của dự án hoặc admin mới thao tác được.</p>'}</div>
    <h2 style="font-size:15px">Lịch sử chuyển</h2><ul class="hist" id="hist"><li>Đang tải…</li></ul>`;
  $('back').onclick = () => { selWp = null; sel = deskOf(w).id; renderSide(true); };
  const move = async (to, label) => {
    const from = w.holder, fromStep = w.step, note = $('note').value.trim();
    const ch = {step: to, holder: holderOf(w, to), note, due: $('due').value};
    try { await S.update(w, ch); await S.hist({key, fromStep, toStep: to, fromHolder: from, toHolder: ch.holder, by: MYNAME || ME.name, note});
      toast(`${label} ${w.wp} → ${to}${ch.holder ? ' (' + ch.holder + ')' : ''}`); selWp = null; sel = deskOf({...w, ...ch}).id; await refresh(true); }
    catch (e) { toast('Không lưu được: ' + e.message); }
  };
  if ($('go')) $('go').onclick = () => move(nx, 'Đã chuyển');
  if ($('ret')) $('ret').onclick = () => { if (!$('note').value.trim()) { toast('Ghi lý do trả lại vào ô Ghi chú trước.'); $('note').focus(); return; } move('Đang lập', 'Đã trả lại'); };
  $('save').onclick = async () => { try { await S.update(w, {note: $('note').value.trim(), due: $('due').value}); toast('Đã lưu'); await refresh(true); } catch (e) { toast('Không lưu được: ' + e.message); } };
  try { const h = await S.history(key); $('hist').innerHTML = h.map(x => `<li>${esc(new Date(x.at).toLocaleString('vi-VN'))}: ${esc(x.By)} chuyển ${esc(x.FromHolder || '–')} → ${esc(x.ToHolder || x.ToStep)} (${esc(x.ToStep)})${x.Note ? ' – ' + esc(x.Note) : ''}</li>`).join('') || '<li>Chưa có lần chuyển nào.</li>'; }
  catch (e) { $('hist').innerHTML = '<li>Không tải được lịch sử.</li>'; }
}

// ================= LÀM MỚI REALTIME =================
function chips() {
  const ps = projects(), el = $('chips');
  if (filter !== 'ALL' && !ps.includes(filter)) filter = 'ALL';
  el.innerHTML = `<button data-f="ALL" aria-pressed="${filter === 'ALL'}">Tất cả</button>` + ps.map(p => `<button data-f="${esc(p)}" aria-pressed="${filter === p}"><span class="dot" style="background:${pcolor(p)}"></span>${esc(p)}</button>`).join('');
}
$('chips').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; filter = b.dataset.f; sel = selWp = null; chips(); drawStacks(); renderSide(true); });
let busy = false, sig = '';
async function refresh(force) {
  if (busy) return; busy = true;
  try {
    const d = await S.load(); ITEMS = d.items; STAFF = d.staff;
    const me = STAFF.find(s => s.email && s.email === ME.email); MYNAME = me ? me.name : ME.name;
    const nsig = people().map(p => p.name + p.level).join('|');
    if (nsig !== sig) { sig = nsig; buildRoom(); }
    const moved = [];
    ITEMS.forEach(w => { const k = w.proj + '|' + w.wp, now = deskOf(w).id, was = prevHolder.get(k); if (!first && was && was !== now) moved.push([w, was, now]); prevHolder.set(k, now); });
    moved.slice(0, 12).forEach(([w, a, b]) => deskById[a] && fly(w, deskById[a], deskById[b]));
    if (moved.length && !force) { const [w] = moved[0]; toast(`${w.by || 'Có người'} vừa chuyển ${w.proj} ${w.wp} → ${w.step}${moved.length > 1 ? ` (+${moved.length - 1} hồ sơ)` : ''}`); }
    first = false; chips(); drawStacks(); renderSide(force);
    $('kStaff').textContent = desks.filter(d => !d.special).length; $('kProj').textContent = projects().length; $('kLate').textContent = visible().filter(isLate).length;
    $('upd').textContent = `${projects().length} dự án · ${ITEMS.length} hồ sơ · cập nhật ${new Date().toLocaleTimeString('vi-VN')}`;
    $('live').hidden = false; $('liveTxt').textContent = DEMO ? 'Bản thử' : 'Trực tuyến';
  } catch (e) { $('liveTxt').textContent = 'Mất kết nối – thử lại…'; console.error(e); }
  busy = false;
}

// ================= TƯƠNG TÁC 3D =================
const ray = new THREE.Raycaster(), ptr = new THREE.Vector2(); let down = null;
function hit(e) { const r = renderer.domElement.getBoundingClientRect(); ptr.set((e.clientX - r.left) / r.width * 2 - 1, -(e.clientY - r.top) / r.height * 2 + 1); ray.setFromCamera(ptr, cam); const h = ray.intersectObjects(pick)[0]; return h && h.object.userData.desk; }
renderer.domElement.addEventListener('pointerdown', e => { down = [e.clientX, e.clientY]; ctl.autoRotate = false; $('bRot').setAttribute('aria-pressed', 'false'); });
renderer.domElement.addEventListener('pointerup', e => { if (!down || Math.hypot(e.clientX - down[0], e.clientY - down[1]) > 6) return; const d = hit(e); sel = d ? d.id : null; selWp = null; if (d) openSide(); drawStacks(); renderSide(true); });
renderer.domElement.addEventListener('pointermove', e => { const d = hit(e); if (!d) { tip.style.display = 'none'; renderer.domElement.style.cursor = ''; return; }
  const n = (byDesk().get(d.id) || []).length; tip.textContent = `${d.name} – ${n} hồ sơ`; const r = stage.getBoundingClientRect();
  tip.style.left = (e.clientX - r.left + 14) + 'px'; tip.style.top = (e.clientY - r.top + 10) + 'px'; tip.style.display = 'block'; renderer.domElement.style.cursor = 'pointer'; });
renderer.domElement.addEventListener('pointerleave', () => tip.style.display = 'none');
function size() { const w = stage.clientWidth, h = stage.clientHeight; if (!w || !h) return; renderer.setSize(w, h, false); cam.aspect = w / h; cam.fov = w < h ? Math.min(62, 30 / (w / h) * .7) : 30; cam.updateProjectionMatrix(); }
const ro = new ResizeObserver(size); ro.observe(stage);
// thanh công cụ dưới
$('bHome').onclick = () => { cam.position.copy(HOME.pos); if (stage.clientWidth < 600) cam.position.y += 30; ctl.target.copy(HOME.tgt); };
$('bTags').onclick = () => { showTags = !showTags; $('bTags').setAttribute('aria-pressed', showTags); tagLayer.hidden = !showTags; };
$('bRot').onclick = () => { ctl.autoRotate = !ctl.autoRotate; $('bRot').setAttribute('aria-pressed', ctl.autoRotate); };
$('bList').onclick = () => { $('side').classList.contains('open') ? closeSide() : openSide(); };
function openSide() { $('side').classList.add('open'); $('bList').setAttribute('aria-pressed', 'true'); }
function closeSide() { $('side').classList.remove('open'); $('bList').setAttribute('aria-pressed', 'false'); }
const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
const V = new THREE.Vector3();
let t = 0;
(function loop(now) {
  requestAnimationFrame(loop); t += .016;
  if (!reduce) {
    stacks.children.forEach(b => { if (b.userData.late) b.material.emissiveIntensity = .3 + .3 * Math.sin(t * 4); });
    desks.forEach((d, i) => { if (d.person) { d.person.userData.head.rotation.y = Math.sin(t * .7 + i) * .25; d.person.position.y = .15 + Math.abs(Math.sin(t * 1.3 + i * 1.7)) * .03; } });
  }
  FL = FL.filter(f => { const p = Math.min(1, ((now || 0) - f.t0) / (reduce ? 1 : 1500)); f.b.position.lerpVectors(f.a, f.z, p); f.b.position.y += Math.sin(Math.PI * p) * 6; f.b.rotation.y = p * Math.PI * 2;
    if (p >= 1) { flights.remove(f.b); return false; } return true; });
  ctl.update(); renderer.render(scene, cam);
  if (showTags) {   // bảng tên HTML bám theo vị trí 3D
    const W = stage.clientWidth, H = stage.clientHeight;
    desks.forEach(d => { const el = TAGS[d.id]; if (!el) return; V.set(d.x, d.special ? 4.6 : 4.4, d.z + (d.special ? 0 : 1.9 * d.face)).project(cam);
      if (V.z > 1) { el.style.display = 'none'; return; } el.style.display = '';
      el.style.transform = `translate(-50%,-100%) translate(${(V.x * .5 + .5) * W}px,${(-V.y * .5 + .5) * H}px)`; el.style.zIndex = String(Math.round((1 - V.z) * 10000)); });
  }
})();

// ================= NẠP DỰ ÁN / THIẾT LẬP =================
$('btnSetup').onclick = async () => { try { toast('Đang tạo 3 danh sách trên SharePoint…'); await S.setup(); toast('Đã tạo xong KMF_HoSo, KMF_NhanSu, KMF_LichSu.'); await refresh(true); } catch (e) { toast('Lỗi thiết lập: ' + e.message); } };
$('btnImport').onclick = () => { $('iMsg').textContent = ''; $('dlgImport').showModal(); };
$('fImport').addEventListener('submit', async e => {
  if (e.submitter && e.submitter.value === 'cancel') return;
  e.preventDefault(); const code = $('iCode').value.trim().toUpperCase(), msg = $('iMsg'); $('iGo').disabled = true;
  try {
    msg.textContent = 'Đang đọc SoTheoDoi_WP.xlsx…';
    const v = await S.sheet($('iUrl').value, '04_Viec_NhanSu/SoTheoDoi_WP.xlsx');
    const title = String(v[0][0] || ''), mt = title.match(/HỒ SƠ – (.*?) – Partner: (.*?) – /), client = mt ? mt[1] : '', partner = mt && !/chưa phân công/i.test(mt[2]) ? mt[2] : '';
    const hi = v.findIndex(r => r[1] === 'WP'); if (hi < 0) throw new Error('Không thấy dòng tiêu đề "WP" trong sổ theo dõi.');
    const have = Object.fromEntries(ITEMS.filter(w => w.proj === code).map(w => [w.wp, w])); let nNew = 0, nUpd = 0;
    for (const r of v.slice(hi + 1)) {
      if (!r[1]) continue;
      const w = {proj: code, client, wp: String(r[1]), name: String(r[2] || ''), part: String(r[3] || ''), prep: String(r[4] || ''), senior: String(r[5] || ''), manager: String(r[6] || ''), partner};
      const old = have[w.wp];
      if (old) { await S.update(old, {client, name: w.name, part: w.part, prep: w.prep, senior: w.senior, manager: w.manager, partner, holder: holderOf({...old, ...w}, old.step)}); nUpd++; }
      else { const step = STEPS.includes(r[7]) ? r[7] : 'Đang lập'; await S.create({...w, step, holder: holderOf(w, step), due: toISO(r[8]), note: String(r[9] || '')}); nNew++; }
      msg.textContent = `Đã nạp ${nNew + nUpd} hồ sơ…`;
    }
    if ($('iBook').value.trim()) {
      msg.textContent = 'Đang đọc danh bạ…';
      const b = await S.sheet($('iBook').value, ''); const hb = b.findIndex(r => r[2] === 'Họ và tên'); let ns = 0;
      for (const r of b.slice(hb + 1)) { if (!r[2] || !r[4]) continue; const ex = STAFF.find(s => s.name === r[2] || (r[3] && s.email === String(r[3]).toLowerCase()));
        await S.upsertStaff({id: ex && ex.id, name: String(r[2]).trim(), email: String(r[3] || '').trim().toLowerCase(), level: r[4], status: r[9] || 'Đang làm việc'}); ns++; }
      msg.textContent += ` Danh bạ: ${ns} người.`;
    }
    msg.textContent = `Xong: ${nNew} hồ sơ mới, ${nUpd} hồ sơ cập nhật thông tin (giữ nguyên bước đang làm).` + (msg.textContent.includes('Danh bạ') ? msg.textContent.slice(msg.textContent.indexOf(' Danh bạ')) : '');
    await refresh(true);
  } catch (err) { msg.textContent = 'Lỗi: ' + err.message; }
  $('iGo').disabled = false;
});

// ================= KHỞI ĐỘNG =================
(async () => {
  try {
    const signed = await S.init();
    if (!signed) { $('gate').hidden = false; $('upd').textContent = 'Chưa đăng nhập'; $('btnLogin').onclick = async () => { if (await Graph.login()) location.reload(); }; return; }
    ME = S.me(); IS_ADMIN = DEMO || (C.admins || []).map(x => x.toLowerCase()).includes(ME.email);
    $('app').hidden = false; size();
    $('btnUser').hidden = false; $('btnUser').textContent = DEMO ? 'Bản thử' : ME.name; if (!DEMO) $('btnUser').onclick = () => Graph.pca.logoutRedirect({account: Graph.acct});
    const ready = await S.open();
    if (IS_ADMIN && !DEMO) { $('btnImport').hidden = !ready; $('btnSetup').hidden = ready; }
    if (!ready) { openSide(); side.innerHTML = `<h2>Chưa có dữ liệu</h2><p>${IS_ADMIN ? 'Bấm <b>Thiết lập</b> để tạo 3 danh sách trên SharePoint (chỉ làm một lần), sau đó bấm <b>Nạp dự án</b>.' : 'Admin chưa thiết lập phần mềm. Liên hệ anh Duy.'}</p>`; return; }
    if (document.fonts) await document.fonts.ready;
    await refresh(true);
    setInterval(() => refresh(false), (C.refreshSeconds || 8) * 1000);
  } catch (e) { $('upd').textContent = 'Lỗi khởi động: ' + e.message; console.error(e); }
})();
}
