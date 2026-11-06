/**
 * ============================================================
 *  ELS.ID — SISTEM PERMINTAAN KEY OFFICE PRO PLUS 2021
 *  Cabang: Purwokerto (PWT), Tegal (TGL), Pekalongan (PKL)
 * ============================================================
 *  Cara pakai singkat:
 *  1. Tempel file ini sebagai Code.gs di Spreadsheet (Extensions > Apps Script)
 *  2. Buat file HTML bernama "Index", tempel isi Index.html
 *  3. Jalankan fungsi setupAwal() sekali (atau menu 🔑 Key Office > Setup Awal)
 *  4. Deploy > New deployment > Web app
 *     - Execute as : Me
 *     - Who has access : Anyone
 * ============================================================
 */

const CFG = {
  // Tambah cabang baru cukup di sini, lalu jalankan setupAwal() lagi.
  // Kode harus sama dengan kode di nomor nota (JL-KODE-...)
  CABANG: { PWT: 'Purwokerto', TGL: 'Tegal', PKL: 'Pekalongan' },
  // Kode nota yang diterima tiap cabang (1 cabang boleh lebih dari 1 kode).
  // Tegal: TGL = gedung lama, TGN = gedung baru
  KODE_NOTA: { PWT: ['PWT'], TGL: ['TGL', 'TGN'], PKL: ['PKL'] },
  SH_STOK: 'STOK_KEY',
  SH_USER: 'USERS',
  SH_SET: 'SETTING',
  SH_SALES: 'SALES',
  SESSION_DETIK: 8 * 60 * 60, // login bertahan 8 jam
  MAX_QTY: 20,               // maksimal key per 1 kali minta
  SH_HARGA: 'HARGA'
};

const REKAP_HEADERS = [
  'NO', 'PERIODE', 'TANGGAL', 'NOMOR NOTA', 'NAMA USER', 'NAMA BARANG', 'PRODUCT KEY',
  'HARGA JUAL', 'HARGA MODAL', 'TGL TRF MODAL', 'SI', 'TGL TRF SI', 'TOTAL MODAL + SI',
  'BU INDAH TRF', 'SALES', 'KODE SALES', 'SI SALES', 'TGL TRF SALES', 'SI TEKSU', 'TGL TRF TEKSU',
  'SISA CLAIM', 'DIMINTA OLEH', 'EMAIL TUJUAN', 'WAKTU MINTA', 'KETERANGAN'
];
const STOK_HEADERS = [
  'PRODUCT KEY', 'STATUS', 'ALOKASI CABANG', 'HARGA MODAL', 'TGL TRF MODAL',
  'TGL MASUK', 'NOMOR NOTA', 'DIPAKAI OLEH', 'TGL DIPAKAI'
];
const SALES_HEADERS = ['CABANG', 'NAMA SALES', 'KODE SALES', 'AKTIF', 'NIK'];
const KEY_RE = /[A-Z0-9]{5}(?:-[A-Z0-9]{5}){4}/g;
const USER_HEADERS = ['USERNAME', 'PIN', 'NAMA AKUN', 'CABANG', 'EMAIL', 'ROLE', 'AKTIF'];
const BULAN = ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI',
  'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'];

/* ============================ WEB APP ============================ */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Permintaan Key Office — ELS.ID')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔑 Key Office')
    .addItem('Setup Awal / Perbaiki Sheet', 'setupAwal')
    .addItem('Lihat Ringkasan Stok', 'menuCekStok_')
    .addItem('Periksa & Sinkron Stok', 'menuSinkron_')
    .addSeparator()
    .addItem('Aktifkan rapikan otomatis tiap malam', 'aktifkanRapikanMalam')
    .addItem('Matikan rapikan otomatis tiap malam', 'matikanRapikanMalam')
    .addToUi();
}

/* ============================ SETUP ============================ */

function setupAwal() {
  const ss = ss_();

  const stok = mkSheet_(CFG.SH_STOK, STOK_HEADERS, '#EE7A12');
  stok.setColumnWidth(1, 270);
  stok.getRange('D2:D').setNumberFormat('#,##0');
  stok.getRange('E2:F').setNumberFormat('dd/MM/yyyy');
  stok.getRange('I2:I').setNumberFormat('dd/MM/yyyy HH:mm');
  stok.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['TERSEDIA', 'TERPAKAI', 'RUSAK/BLOKIR'], true).setAllowInvalid(false).build());
  stok.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['UMUM'].concat(Object.keys(CFG.CABANG)), true).setAllowInvalid(false).build());

  const users = mkSheet_(CFG.SH_USER, USER_HEADERS, '#EE7A12');
  if (users.getLastRow() < 2) {
    const myEmail = Session.getActiveUser().getEmail() || 'email.admin@gmail.com';
    users.getRange(2, 1, 4, 7).setValues([
      ['superadmin', '1234', 'Andi (Super Admin)', 'ALL', myEmail, 'ADMIN', true],
      ['pwt', '1111', 'Toko Purwokerto', 'PWT', 'tekspuwt@gmail.com', 'USER', true],
      ['tgl', '2222', 'Toko Tegal', 'TGL', 'pic.tegal@elsnotebookcenter.com', 'USER', true],
      ['pkl', '3333', 'Toko Pekalongan', 'PKL', 'isi.email.pekalongan@gmail.com', 'USER', true]
    ]);
  }
  users.getRange('B2:B').setNumberFormat('@');
  users.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(CFG.CABANG).concat(['ALL']), true).build());
  users.getRange('F2:F').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['USER', 'ADMIN'], true).build());
  users.getRange('G2:G').insertCheckboxes();

  const set = mkSheet_(CFG.SH_SET, ['PENGATURAN', 'NILAI', 'KETERANGAN'], '#EE7A12');
  ensureSettings_(set);
  set.setColumnWidth(1, 200).setColumnWidth(2, 240).setColumnWidth(3, 380);

  const sal = mkSheet_(CFG.SH_SALES, SALES_HEADERS, '#EE7A12');
  if (sal.getLastRow() < 2) {
    const pwt = [['ANIQ','SPW-O-ANIQ'],['NOVI','SPW-O-NOVI'],['DARMANTO','SPW-O-DARMANTO'],['ANGGUN','SPW-O-ANGGUN'],
      ['WULAN','SPW-O-WULAN'],['PITA','SPW-R-PITA'],['ANAN','SPW-R-ANAN'],['AYI','SPW-R-AYI'],['JULIAN','SPW-R-JULIAN'],
      ['ANTENG','SPW-R-ANTENG'],['TATIK','SPW-R-TATIK'],['PIPIH','SPW-R-PIPIH'],['RIRIH','SPW-R-RIRIH'],['KHOTI','SPW-R-KHOTI'],
      ['DIAN','SPW-R-DIAN'],['GIGIH','SPW-R-GIGIH'],['ANTO','SPW-R-P-AC-1'],['SUSI','SPW-R-P-AC-4'],['DINA','SPW-R-P-AS-1'],
      ['RAFLI','SPW-R-P-AS-2'],['AJI','SPW-R-P-LE-1'],['NAFIS','SPW-R-P-LE-2'],['HANA','SPW-R-P-LE-3'],['DITA','SPW-R-P-HP-5'],
      ['SHENY','SPW-R-P-HP-2'],['ABRI','SPW-R-P-AX-2'],['ANGGITA','SPW-R-P-AX-3'],['RIZAL','SPW-R-P-ADV-1']];
    sal.getRange(2, 1, pwt.length, 4).setValues(pwt.map(r => ['PWT', r[0], r[1], true]));
  }
  sal.getRange('A2:A').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(CFG.CABANG), true).build());
  sal.getRange('D2:D').insertCheckboxes();
  sal.setColumnWidth(2, 180).setColumnWidth(3, 180).setColumnWidth(5, 160);
  sal.getRange('E2:E').setNumberFormat('@');

  Object.keys(CFG.CABANG).forEach(k => mkRekap_('REKAP_' + k));
  hargaSheet_();

  try { ss.toast('Setup selesai. Cek sheet USERS untuk atur akun & PIN.', '🔑 Key Office', 6); } catch (e) {}
}

function ensureSettings_(set) {
  const defaults = [
    ['NAMA_BARANG', 'OFFICE PRO PLUS 2021', 'Nama barang di rekap'],
    ['MIN_STOK', 10, 'Kirim email peringatan jika stok cabang <= angka ini'],
    ['EMAIL_ADMIN', Session.getActiveUser().getEmail() || '', 'Penerima peringatan stok menipis'],
    ['PERIODE_DARI', 'MINTA', 'MINTA = periode ikut tanggal minta key | NOTA = ikut tanggal di nomor nota'],
    ['GABUNG_HARI', 3, 'Sisa minggu di awal/akhir bulan yang <= angka ini digabung ke minggu sebelahnya (0 = tidak digabung)']
  ];
  const have = {};
  if (set.getLastRow() > 1) set.getRange(2, 1, set.getLastRow() - 1, 1).getValues().forEach(r => have[String(r[0]).trim()] = true);
  const add = defaults.filter(d => !have[d[0]]);
  if (add.length) set.getRange(set.getLastRow() + 1, 1, add.length, 3).setValues(add);
}

function mkSheet_(name, headers, color) {
  const ss = ss_();
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(color).setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 30);
  return sh;
}

function mkRekap_(name) {
  const sh = mkSheet_(name, REKAP_HEADERS, '#F6A21A');
  const widths = [45, 200, 90, 190, 180, 170, 270, 90, 95, 105, 90, 95, 125, 105, 100, 140, 85, 105, 85, 105, 95, 150, 200, 135, 240];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  ['C', 'J', 'L', 'N', 'R', 'T'].forEach(c => sh.getRange(c + '2:' + c).setNumberFormat('dd/MM/yyyy'));
  ['H', 'I', 'K', 'M', 'Q', 'S', 'U'].forEach(c => sh.getRange(c + '2:' + c).setNumberFormat('#,##0'));
  sh.getRange('X2:X').setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange('A2:A').setHorizontalAlignment('center');
  if (!sh.getFilter()) sh.getRange(1, 1, sh.getMaxRows(), REKAP_HEADERS.length).createFilter();
  return sh;
}

function menuCekStok_() {
  const c = countStock_();
  const baris = Object.keys(CFG.CABANG).map(k => 'Khusus ' + CFG.CABANG[k] + ': ' + c[k]).join('\n');
  SpreadsheetApp.getUi().alert(
    'Ringkasan Stok Key\n\nUmum (semua cabang): ' + c.UMUM + '\n' + baris +
    '\n\nTotal tersedia: ' + c.total + '\nSudah terpakai: ' + c.TERPAKAI);
}

/* ============================ URUTAN NOTA ============================ */
/** Urutan rekap: tanggal nota -> reguler dulu baru KIOSK -> kode gedung -> nomor urut nota -> waktu minta */
function kunciUrut_(r, cab) {
  const n = parseNota_(String(r[3]).toUpperCase().trim());
  const kodeList = (CFG.KODE_NOTA && CFG.KODE_NOTA[cab]) || [cab];
  const tgl = r[2] instanceof Date ? r[2].getTime() : (n ? n.tanggal.getTime() : 9e15);
  return [
    tgl,
    n ? (n.kiosk ? 1 : 0) : 2,
    n ? Math.max(0, kodeList.indexOf(n.kode)) : 99,
    n ? Number(n.urut) : 9e9,
    r[23] instanceof Date ? r[23].getTime() : 9e15
  ];
}
function bandingUrut_(a, b) {
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return a[i] - b[i];
  return 0;
}

/* ============================ SINKRON ============================ */
/**
 * Menyamakan STOK_KEY dengan isi semua sheet REKAP (dipakai setelah edit manual di sheet):
 *  - key yang ada di rekap  -> TERPAKAI (+ nota, peminta, waktu)
 *  - key TERPAKAI yang tidak ada di rekap -> kembali TERSEDIA
 *  - rekap: baris kosong dihapus, NO diurutkan ulang, rumus TOTAL & SISA CLAIM diperbaiki
 * Status RUSAK/BLOKIR tidak disentuh.
 */
function sinkronData_(opt) {
  opt = opt || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const hasil = { dikembalikan: [], ditandai: 0, barisKosong: 0, hargaDiisi: 0, hurufBesar: 0, diurutkan: [], keyDobelRekap: [], tidakDiStok: [], keyDobelStok: [] };
    const hargaList = daftarHarga_();
    const pakai = {};

    Object.keys(CFG.CABANG).forEach(k => {
      const sh = ss_().getSheetByName('REKAP_' + k);
      if (!sh || sh.getLastRow() < 2) return;
      // hapus baris kosong (bekas cut/hapus isi) dari bawah ke atas
      const vals = sh.getRange(2, 1, sh.getLastRow() - 1, REKAP_HEADERS.length).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        const r = vals[i];
        if (!String(r[3]).trim() && !String(r[4]).trim() && !String(r[6]).trim()) { sh.deleteRow(i + 2); hasil.barisKosong++; }
      }
      const n = sh.getLastRow() - 1;
      if (n < 1) return;
      let rows = sh.getRange(2, 1, n, REKAP_HEADERS.length).getValues();
      // urutkan: tanggal nota -> reguler -> kiosk -> nomor urut
      const urut = rows.map((r, i) => ({ r: r, k: kunciUrut_(r, k), i: i }))
        .sort((a, b) => bandingUrut_(a.k, b.k) || (a.i - b.i));
      if (urut.some((x, i) => x.i !== i)) {
        rows = urut.map(x => x.r);
        sh.getRange(2, 1, n, REKAP_HEADERS.length).setValues(rows);
        hasil.diurutkan.push(CFG.CABANG[k]);
      }
      const no = [], fT = [], fS = [], hk = [], sq = [], st2 = [], ue = [], uy = [];
      rows.forEach((r, i) => {
        const rr = i + 2;
        no.push([i + 1]);
        const up = v => (typeof v === 'string' ? clean_(v).toUpperCase() : v);
        const nu = up(r[4]), nk = up(r[24]);
        if (nu !== r[4] || nk !== r[24]) hasil.hurufBesar++;
        r[4] = nu; r[24] = nk; ue.push([nu]); uy.push([nk]);
        // harga: H=jual, I=modal, J=tgl trf modal (tetap), K=SI
        const tMinta = r[23] instanceof Date ? r[23] : (r[2] instanceof Date ? r[2] : new Date());
        const hg = hargaBerlaku_(tMinta, hargaList);
        const baru = [r[7], r[8], r[9], r[10]], qs = [r[16], r[18]];
        if (String(r[6]).trim()) {
          const nh = nilaiHarga_(hg, !!String(r[24]).trim());
          let ubah = false;
          const set = (arr, idx, v) => {
            const kosong = arr[idx] === '' || arr[idx] === null;
            if (opt.timpaHarga ? arr[idx] !== v : (kosong && v !== '')) { arr[idx] = v; ubah = true; }
          };
          set(baru, 0, nh.jual); set(baru, 1, nh.modal); set(baru, 3, nh.si);
          set(qs, 0, nh.siSales); set(qs, 1, nh.siTeksu);
          if (ubah) hasil.hargaDiisi++;
        }
        hk.push(baru); sq.push([qs[0]]); st2.push([qs[1]]);
        fT.push(['=IF(AND(I' + rr + '="",K' + rr + '=""),"",N(I' + rr + ')+N(K' + rr + '))']);
        fS.push(['=IF(K' + rr + '="","",N(K' + rr + ')-N(Q' + rr + ')-N(S' + rr + '))']);
        const key = String(r[6]).trim().toUpperCase();
        if (!key) return;
        if (pakai[key]) hasil.keyDobelRekap.push(key + ' (' + pakai[key].cab + ' & ' + k + ')');
        else pakai[key] = { cab: k, nota: String(r[3]), oleh: String(r[21]) + ' (' + k + ')', waktu: r[23] };
      });
      sh.getRange(2, 1, n, 1).setValues(no);
      sh.getRange(2, 5, n, 1).setValues(ue);
      sh.getRange(2, 25, n, 1).setValues(uy);
      sh.getRange(2, 8, n, 4).setValues(hk);
      sh.getRange(2, 17, n, 1).setValues(sq);
      sh.getRange(2, 19, n, 1).setValues(st2);
      sh.getRange(2, 13, n, 1).setFormulas(fT);
      sh.getRange(2, 21, n, 1).setFormulas(fS);
    });

    const st = ss_().getSheetByName(CFG.SH_STOK);
    if (st && st.getLastRow() > 1) {
      const n = st.getLastRow() - 1;
      const data = st.getRange(2, 1, n, 9).getValues();
      const seen = {};
      data.forEach(r => {
        const key = String(r[0]).trim().toUpperCase();
        if (!key) return;
        if (seen[key]) { hasil.keyDobelStok.push(key); return; }
        seen[key] = true;
        const status = String(r[1]).trim().toUpperCase();
        if (status === 'RUSAK/BLOKIR') return;
        const p = pakai[key];
        if (p) {
          if (status !== 'TERPAKAI' || String(r[6]) !== p.nota) hasil.ditandai++;
          r[1] = 'TERPAKAI'; r[6] = p.nota; r[7] = p.oleh; r[8] = p.waktu || r[8] || new Date();
        } else if (status === 'TERPAKAI') {
          hasil.dikembalikan.push(key);
          r[1] = 'TERSEDIA'; r[6] = ''; r[7] = ''; r[8] = '';
        } else if (!status) {
          r[1] = 'TERSEDIA';
        }
      });
      st.getRange(2, 2, n, 1).setValues(data.map(r => [r[1]]));
      st.getRange(2, 7, n, 3).setValues(data.map(r => [r[6], r[7], r[8]]));
      Object.keys(pakai).forEach(k => { if (!seen[k]) hasil.tidakDiStok.push(k); });
    }
    SpreadsheetApp.flush();
    bustCache_();
    hasil.stok = countStock_();
    return hasil;
  } finally {
    lock.releaseLock();
  }
}

/** Dijalankan trigger tiap malam: urutkan & rapikan rekap tanpa menimpa harga */
function rapikanMalam() { sinkronData_({ timpaHarga: false }); }

function aktifkanRapikanMalam() {
  matikanRapikanMalam(true);
  ScriptApp.newTrigger('rapikanMalam').timeBased().everyDays(1).atHour(23).create();
  try { SpreadsheetApp.getUi().alert('Rapikan otomatis aktif: tiap malam sekitar jam 23.00 rekap diurutkan & disinkron.'); } catch (e) {}
}

function matikanRapikanMalam(diam) {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'rapikanMalam') ScriptApp.deleteTrigger(t); });
  if (diam !== true) { try { SpreadsheetApp.getUi().alert('Rapikan otomatis tiap malam dimatikan.'); } catch (e) {} }
}

function sinkronStok(token, opt) {
  const u = auth_(token);
  if (u.role !== 'ADMIN') throw new Error('Khusus admin.');
  const h = sinkronData_({ timpaHarga: !!(opt && opt.timpaHarga) });
  h.stokCabang = stokCabang_(h.stok);
  return h;
}

function menuSinkron_() {
  const ui = SpreadsheetApp.getUi();
  const jawab = ui.alert('Periksa & Sinkron',
    'Samakan juga HARGA MODAL, HARGA JUAL, SI, SI SALES dan SI TEKSU di semua rekap dengan sheet HARGA (sesuai tanggal minta key)?\n\n' +
    'YES = timpa semua harga · NO = hanya isi harga yang kosong', ui.ButtonSet.YES_NO_CANCEL);
  if (jawab === ui.Button.CANCEL || jawab === ui.Button.CLOSE) return;
  const h = sinkronData_({ timpaHarga: jawab === ui.Button.YES });
  ui.alert('Sinkron selesai\n\n' +
    'Key dikembalikan ke TERSEDIA: ' + h.dikembalikan.length + (h.dikembalikan.length ? '\n  ' + h.dikembalikan.join('\n  ') : '') + '\n' +
    'Status diperbaiki jadi TERPAKAI: ' + h.ditandai + '\n' +
    'Baris kosong di rekap dihapus: ' + h.barisKosong + '\n' +
    'Baris rekap yang harganya disesuaikan: ' + h.hargaDiisi + '\n' +
    'Nama user / keterangan dijadikan huruf besar: ' + h.hurufBesar + '\n' +
    'Rekap diurutkan ulang: ' + (h.diurutkan.length ? h.diurutkan.join(', ') : 'sudah urut') + '\n' +
    (h.keyDobelRekap.length ? '\n⚠️ Key dipakai di 2 baris rekap:\n  ' + h.keyDobelRekap.join('\n  ') + '\n' : '') +
    (h.keyDobelStok.length ? '\n⚠️ Key dobel di STOK_KEY:\n  ' + h.keyDobelStok.join('\n  ') + '\n' : '') +
    (h.tidakDiStok.length ? '\n⚠️ Key di rekap tapi tidak ada di STOK_KEY:\n  ' + h.tidakDiStok.join('\n  ') + '\n' : '') +
    '\nStok tersedia sekarang: ' + h.stok.total + ' · Terpakai: ' + h.stok.TERPAKAI);
}

/* ============================ AUTH ============================ */

function login(username, pin) {
  username = String(username || '').trim().toLowerCase();
  pin = String(pin || '').trim();
  if (!username || !pin) throw new Error('Isi username dan PIN.');
  const u = findUser_(username);
  if (!u || String(u.pin) !== pin) throw new Error('Username atau PIN salah.');
  if (!u.aktif) throw new Error('Akun ini sedang nonaktif. Hubungi admin.');
  const token = Utilities.getUuid();
  const sess = { username: u.username, nama: u.nama, cabang: u.cabang, email: u.email, role: u.role, aktif: true };
  CacheService.getScriptCache().put('S_' + token, JSON.stringify(sess), CFG.SESSION_DETIK);
  return { token: token, data: buildDashboard_(u) };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('S_' + token);
  return true;
}

function auth_(token) {
  const raw = token ? CacheService.getScriptCache().get('S_' + token) : null;
  if (!raw) throw new Error('SESSION_EXPIRED');
  try { const u = JSON.parse(raw); if (u && u.username) return u; } catch (e) {}
  throw new Error('SESSION_EXPIRED');
}

function findUser_(username) {
  const sh = ss_().getSheetByName(CFG.SH_USER);
  if (!sh) throw new Error('Sheet USERS belum ada. Jalankan setupAwal().');
  const rows = sh.getDataRange().getValues().slice(1);
  for (const r of rows) {
    if (String(r[0]).trim().toLowerCase() === username) {
      return {
        username: String(r[0]).trim().toLowerCase(),
        pin: String(r[1]).trim(),
        nama: String(r[2]).trim(),
        cabang: String(r[3]).trim().toUpperCase(),
        email: String(r[4]).trim(),
        role: String(r[5]).trim().toUpperCase(),
        aktif: r[6] === true || String(r[6]).toUpperCase() === 'TRUE'
      };
    }
  }
  return null;
}

/* ============================ DASHBOARD ============================ */

function getDashboard(token) {
  return buildDashboard_(auth_(token));
}

function buildDashboard_(u) {
  const cabangAkses = (u.cabang === 'ALL' || u.role === 'ADMIN') ? Object.keys(CFG.CABANG) : [u.cabang];
  const stok = countStock_();
  return {
    user: { nama: u.nama, cabang: u.cabang, email: u.email, role: u.role, username: u.username },
    cabangMap: CFG.CABANG,
    cabangAkses: cabangAkses,
    stok: stok,
    stokCabang: stokCabang_(stok),
    sales: getSalesMaster_(cabangAkses),
    minStok: Number(getSettings_().MIN_STOK || 10),
    maxQty: CFG.MAX_QTY,
    kodeNota: kodeNotaMap_(),
    ver: (function () { try { return CacheService.getScriptCache().get('LIVE_VER') || ''; } catch (e) { return ''; } })()
  };
}

function stokCabang_(stok) {
  const o = {};
  Object.keys(CFG.CABANG).forEach(k => o[k] = (stok[k] || 0) + stok.UMUM);
  return o;
}

/** Data tab Admin (dimuat saat tab dibuka) */
function getAdminData(token) {
  const u = auth_(token);
  if (u.role !== 'ADMIN') throw new Error('Khusus admin.');
  const sum = adminSummary_();
  const stok = countStock_();
  return { stats: sum.stats, riwayat: sum.riwayat, stok: stok, stokCabang: stokCabang_(stok) };
}

/** Ringkasan pemakaian (hari ini / minggu ini / bulan ini) + 10 permintaan terakhir.
 *  Disimpan di cache 2 menit, dihapus otomatis setiap ada permintaan/edit/sinkron. */
function adminSummary_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('ADM_SUM');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }

  const now = new Date();
  const awalHari = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const awalMinggu = awalHari - ((now.getDay() + 6) % 7) * 864e5;
  const awalBulan = new Date(now.getFullYear(), now.getMonth(), 1).getTime();
  const stats = {}, riwayat = [];
  Object.keys(CFG.CABANG).forEach(k => {
    const s = { hari: 0, minggu: 0, bulan: 0 };
    stats[k] = s;
    const sh = ss_().getSheetByName('REKAP_' + k);
    if (!sh || sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, REKAP_HEADERS.length).getValues();
    const top = [];
    rows.forEach(r => {
      if (!r[3] || !(r[23] instanceof Date)) return;
      const t = r[23].getTime();
      if (t >= awalHari) s.hari++;
      if (t >= awalMinggu) s.minggu++;
      if (t >= awalBulan) s.bulan++;
      top.push([t, r]);
    });
    top.sort((a, b) => b[0] - a[0]);
    top.slice(0, 10).forEach(x => {
      const r = x[1];
      riwayat.push({ cabang: k, nota: r[3], user: r[4], key: r[6], sales: r[14], ket: r[24], oleh: r[21], waktu: x[0] });
    });
  });
  riwayat.sort((a, b) => b.waktu - a.waktu);
  const out = { stats: stats, riwayat: riwayat.slice(0, 10) };
  try { cache.put('ADM_SUM', JSON.stringify(out), 120); } catch (e) {}
  return out;
}

/** Tandai data berubah: semua layar yang terbuka akan mengambil stok terbaru pada pengecekan berikutnya */
function bustCache_() {
  try {
    const c = CacheService.getScriptCache();
    c.removeAll(['ADM_SUM', 'STOK_CNT']);
    c.put('LIVE_VER', String(Date.now()), 21600);
  } catch (e) {}
}

/** Jumlah stok dari cache (60 detik), dihitung ulang setiap ada perubahan */
function stokCached_() {
  const c = CacheService.getScriptCache();
  const hit = c.get('STOK_CNT');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const s = countStock_();
  try { c.put('STOK_CNT', JSON.stringify(s), 60); } catch (e) {}
  return s;
}

/**
 * Dipanggil layar setiap ±10 detik. Sangat ringan:
 * kalau versi data sama -> langsung jawab "tidak ada perubahan" tanpa membaca sheet.
 */
function getLive(token, ver) {
  auth_(token);
  const c = CacheService.getScriptCache();
  let v = c.get('LIVE_VER');
  if (!v) { v = String(Date.now()); c.put('LIVE_VER', v, 21600); }
  if (ver && ver === v) return { ver: v, same: true, t: Date.now() };
  const stok = stokCached_();
  return { ver: v, same: false, stok: stok, stokCabang: stokCabang_(stok), t: Date.now() };
}

/** Total pemakaian per periode. p = { mode: 'HARI'|'MINGGU'|'BULAN', dari: 'yyyy-mm-dd', sampai: 'yyyy-mm-dd' } */
function getUsage(token, p) {
  const u = auth_(token);
  if (u.role !== 'ADMIN') throw new Error('Khusus admin.');
  const mode = String(p.mode || 'BULAN').toUpperCase();
  let a = new Date(p.dari + 'T00:00:00'), b = new Date(p.sampai + 'T00:00:00');
  if (isNaN(a) || isNaN(b)) throw new Error('Tanggal belum lengkap.');
  if (a > b) { const t = a; a = b; b = t; }
  if (mode === 'MINGGU') { a = new Date(a.getTime() - ((a.getDay() + 6) % 7) * 864e5); b = new Date(b.getTime() + (6 - (b.getDay() + 6) % 7) * 864e5); }
  if (mode === 'BULAN') { a = new Date(a.getFullYear(), a.getMonth(), 1); b = new Date(b.getFullYear(), b.getMonth() + 1, 0); }
  const end = new Date(b.getFullYear(), b.getMonth(), b.getDate(), 23, 59, 59).getTime();
  const BLN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const p2 = x => ('0' + x).slice(-2);

  // susun ember periode
  const buckets = [];
  let cur = new Date(a);
  while (cur.getTime() <= end && buckets.length < 400) {
    let next, label;
    if (mode === 'HARI') { next = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1); label = p2(cur.getDate()) + '/' + p2(cur.getMonth() + 1) + '/' + cur.getFullYear(); }
    else if (mode === 'MINGGU') { next = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 7); const e = new Date(next.getTime() - 864e5);
      label = p2(cur.getDate()) + ' ' + BLN[cur.getMonth()] + ' – ' + p2(e.getDate()) + ' ' + BLN[e.getMonth()] + ' ' + e.getFullYear(); }
    else { next = new Date(cur.getFullYear(), cur.getMonth() + 1, 1); label = BLN[cur.getMonth()] + ' ' + cur.getFullYear(); }
    const row = { label: label, mulai: cur.getTime(), akhir: next.getTime(), total: 0 };
    Object.keys(CFG.CABANG).forEach(k => row[k] = 0);
    buckets.push(row);
    cur = next;
  }
  if (buckets.length >= 400) throw new Error('Rentang terlalu panjang. Persempit periodenya.');

  const total = { total: 0 };
  Object.keys(CFG.CABANG).forEach(k => {
    total[k] = 0;
    const sh = ss_().getSheetByName('REKAP_' + k);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 24, sh.getLastRow() - 1, 1).getValues().forEach(r => {
      if (!(r[0] instanceof Date)) return;
      const t = r[0].getTime();
      if (t < a.getTime() || t > end) return;
      const bk = buckets.find(x => t >= x.mulai && t < x.akhir);
      if (bk) { bk[k]++; bk.total++; total[k]++; total.total++; }
    });
  });
  return { mode: mode, dari: a.getTime(), sampai: b.getTime(), buckets: buckets, total: total };
}

function getSalesMaster_(cabangList) {
  const out = {};
  cabangList.forEach(k => out[k] = []);
  const sh = ss_().getSheetByName(CFG.SH_SALES);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(r => {
    const cab = String(r[0]).trim().toUpperCase();
    const aktif = r[3] === true || String(r[3]).toUpperCase() === 'TRUE';
    if (!out[cab] || !aktif || !String(r[1]).trim()) return;
    out[cab].push({ nama: String(r[1]).trim().toUpperCase(), kode: String(r[2]).trim().toUpperCase() });
  });
  Object.keys(out).forEach(k => out[k].sort((a, b) => a.nama.localeCompare(b.nama)));
  return out;
}

/* ============================ HARGA ============================ */
// Sheet HARGA: tiap baris = harga yang berlaku MULAI tanggal tsb. Data lama tidak diubah.

const HARGA_HEADERS = ['BERLAKU MULAI', 'HARGA MODAL', 'HARGA JUAL', 'SI CLAIM', 'DIUBAH OLEH', 'WAKTU UBAH', 'SI SALES', 'SI TEKSU'];

function hargaSheet_() {
  let sh = ss_().getSheetByName(CFG.SH_HARGA);
  if (!sh) {
    sh = mkSheet_(CFG.SH_HARGA, HARGA_HEADERS, '#EE7A12');
    sh.getRange(2, 1, 1, 8).setValues([[new Date(2020, 0, 1), 21000, 300000, 30000, 'Default', new Date(), 10000, 10000]]);
    sh.getRange('A2:A').setNumberFormat('dd/MM/yyyy');
    sh.getRange('B2:D').setNumberFormat('#,##0');
    sh.getRange('G2:H').setNumberFormat('#,##0');
    sh.getRange('F2:F').setNumberFormat('dd/MM/yyyy HH:mm');
    [130, 120, 120, 110, 180, 140, 110, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  } else if (String(sh.getRange(1, 7).getValue()).trim() !== 'SI SALES') {
    // migrasi sheet HARGA lama: tambah kolom SI SALES & SI TEKSU, isi default 10.000
    mkSheet_(CFG.SH_HARGA, HARGA_HEADERS, '#EE7A12');
    if (sh.getLastRow() > 1) {
      const n = sh.getLastRow() - 1;
      const gh = sh.getRange(2, 7, n, 2).getValues().map(r => [r[0] === '' ? 10000 : r[0], r[1] === '' ? 10000 : r[1]]);
      sh.getRange(2, 7, n, 2).setValues(gh).setNumberFormat('#,##0');
    }
    sh.setColumnWidth(7, 110).setColumnWidth(8, 110);
  }
  return sh;
}

function daftarHarga_() {
  const sh = hargaSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues()
    .filter(r => r[0] instanceof Date)
    .map(r => ({ mulai: r[0].getTime(), modal: numOr0_(r[1]), jual: numOr0_(r[2]), si: numOr0_(r[3]),
      oleh: String(r[4] || ''), waktu: r[5] instanceof Date ? r[5].getTime() : 0,
      siSales: numOr0_(r[6]), siTeksu: numOr0_(r[7]) }))
    .sort((a, b) => (b.mulai - a.mulai) || (b.waktu - a.waktu));
}

function hargaBerlaku_(tgl, list) {
  const t = new Date(tgl.getFullYear(), tgl.getMonth(), tgl.getDate(), 23, 59, 59).getTime();
  const h = (list || daftarHarga_()).find(x => x.mulai <= t);
  return h || { modal: 21000, jual: 300000, si: 30000, siSales: 10000, siTeksu: 10000 };
}

function getHarga(token) {
  const u = auth_(token);
  if (u.role !== 'ADMIN') throw new Error('Khusus admin.');
  const list = daftarHarga_();
  return { aktif: hargaBerlaku_(new Date()), riwayat: list };
}

function setHarga(token, p) {
  const u = auth_(token);
  if (u.role !== 'ADMIN') throw new Error('Khusus admin.');
  const angka = v => Number(String(v == null ? '' : v).replace(/[^\d]/g, ''));
  const modal = angka(p.modal), jual = angka(p.jual), si = angka(p.si), siSales = angka(p.siSales), siTeksu = angka(p.siTeksu);
  if (!jual) throw new Error('Harga jual wajib diisi.');
  if (siSales + siTeksu > si) throw new Error('SI sales + SI teksu melebihi SI claim.');
  if (!p.mulai) throw new Error('Tanggal berlaku wajib diisi.');
  const mulai = new Date(p.mulai + 'T00:00:00');
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (mulai < today) throw new Error('Tanggal berlaku minimal hari ini, supaya data yang sudah ada tidak berubah.');
  const sh = hargaSheet_();
  sh.appendRow([mulai, modal, jual, si, u.nama, new Date(), siSales, siTeksu]);
  return getHarga(token);
}

/**
 * Nilai kolom harga untuk 1 baris rekap.
 * Normal: jual, modal, SI, SI sales, SI teksu.
 * Ada keterangan (replace / ganti unit): hanya modal, sisanya kosong.
 */
function nilaiHarga_(h, adaKet) {
  return adaKet
    ? { jual: '', modal: h.modal, si: '', siSales: '', siTeksu: '' }
    : { jual: h.jual, modal: h.modal, si: h.si, siSales: h.siSales, siTeksu: h.siTeksu };
}

/* ============================ REKAP & EXPORT ============================ */

/** Filter: { cabang: 'ALL' | kode, dari: 'yyyy-mm-dd', sampai: 'yyyy-mm-dd', basis: 'NOTA' | 'MINTA' } */
function queryRekap_(f) {
  const cabs = (!f.cabang || f.cabang === 'ALL') ? Object.keys(CFG.CABANG) : [String(f.cabang).toUpperCase()];
  const dari = f.dari ? new Date(f.dari + 'T00:00:00') : null;
  const sampai = f.sampai ? new Date(f.sampai + 'T23:59:59') : null;
  const byMinta = String(f.basis || 'NOTA').toUpperCase() === 'MINTA';
  const out = [];
  const nik = nikMap_();
  cabs.forEach(k => {
    if (!CFG.CABANG[k]) return;
    const sh = ss_().getSheetByName('REKAP_' + k);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, REKAP_HEADERS.length).getValues().forEach((r, i) => {
      if (!r[3]) return;
      const d = byMinta ? r[23] : r[2];
      if (!(d instanceof Date)) return;
      if (dari && d < dari) return;
      if (sampai && d > sampai) return;
      out.push(rowObj_(k, r, i + 2, nik));
    });
  });
  out.sort((a, b) => (a.tanggal - b.tanggal) || (a.waktu - b.waktu));
  return out;
}

/** Peta KODE SALES -> NIK dari sheet SALES (NIK boleh kosong) */
function nikMap_() {
  const m = {};
  const sh = ss_().getSheetByName(CFG.SH_SALES);
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(r => {
    const kode = String(r[2]).trim().toUpperCase(), nik = String(r[4] || '').trim();
    if (kode && nik) m[kode] = nik;
  });
  return m;
}

function rowObj_(k, r, rowNo, nik) {
  return {
    cabang: k, row: rowNo, tanggal: r[2] instanceof Date ? r[2].getTime() : 0, nota: String(r[3]), user: String(r[4]),
    barang: String(r[5]), key: String(r[6]), jual: numOr0_(r[7]), modal: numOr0_(r[8]), si: numOr0_(r[10]),
    sales: String(r[14]), kodeSales: String(r[15]), siSales: numOr0_(r[16]), siTeksu: numOr0_(r[18]), claim: numOr0_(r[20]), oleh: String(r[21]),
    waktu: r[23] instanceof Date ? r[23].getTime() : 0, ket: String(r[24] || ''),
    nik: (nik && nik[String(r[15]).toUpperCase()]) || ''
  };
}

/** Admin: perbaiki nomor nota / nama user / sales / keterangan. Tercatat di LOG_EDIT. */
function updateRekap(token, p) {
  const u = auth_(token);
  if (u.role !== 'ADMIN') throw new Error('Khusus admin.');
  const cab = String(p.cabang || '').toUpperCase();
  if (!CFG.CABANG[cab]) throw new Error('Cabang tidak dikenal.');
  const nota = normNota_(p.nota);
  const n = parseNota_(nota);
  if (!n) throw new Error('Format nomor nota tidak dikenali.');
  if (n.cabang !== cab) throw new Error('Kode ' + n.kode + ' di nota bukan kode cabang ' + CFG.CABANG[cab] + '.');
  const namaUser = clean_(p.namaUser).toUpperCase();
  if (!namaUser) throw new Error('Nama user wajib diisi.');
  const ket = clean_(p.ket).toUpperCase();
  const kodeSales = clean_(p.kodeSales).toUpperCase();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName('REKAP_' + cab);
    let rowNo = Number(p.row);
    let r = rowNo >= 2 && rowNo <= sh.getLastRow() ? sh.getRange(rowNo, 1, 1, REKAP_HEADERS.length).getValues()[0] : [];
    if (String(r[6]) !== p.key) {
      const col = sh.getRange(2, 7, sh.getLastRow() - 1, 1).getValues();
      const i = col.findIndex(x => String(x[0]) === p.key);
      if (i < 0) throw new Error('Data tidak ditemukan. Muat ulang rekap.');
      rowNo = i + 2;
      r = sh.getRange(rowNo, 1, 1, REKAP_HEADERS.length).getValues()[0];
    }

    let sales = String(r[14]), kode = String(r[15]);
    if (kodeSales && kodeSales !== kode.toUpperCase()) {
      const s = (getSalesMaster_([cab])[cab] || []).find(x => x.kode === kodeSales);
      if (!s) throw new Error('Sales tidak ada di daftar.');
      sales = s.nama; kode = s.kode;
    }

    const perubahan = [];
    const cek = (label, a, b) => { if (String(a) !== String(b)) perubahan.push(label + ': ' + a + ' → ' + b); };
    cek('Nota', r[3], nota); cek('User', r[4], namaUser); cek('Sales', r[15], kode); cek('Ket', r[24], ket);
    if (!perubahan.length) return { ok: true, row: rowObj_(cab, r, rowNo, nikMap_()), berubah: false };

    const ketBerubah = !!String(r[24]).trim() !== !!ket;
    r[2] = n.tanggal; r[3] = nota; r[4] = namaUser; r[14] = sales; r[15] = kode; r[24] = ket;
    if (ketBerubah) {
      const t = r[23] instanceof Date ? r[23] : new Date();
      const nh = nilaiHarga_(hargaBerlaku_(t), !!ket);
      r[7] = nh.jual; r[8] = nh.modal; r[10] = nh.si; r[16] = nh.siSales; r[18] = nh.siTeksu;
      sh.getRange(rowNo, 8, 1, 2).setValues([[r[7], r[8]]]);
      sh.getRange(rowNo, 11).setValue(r[10]);
      sh.getRange(rowNo, 17).setValue(r[16]);
      sh.getRange(rowNo, 19).setValue(r[18]);
      perubahan.push('Harga dihitung ulang (' + (ket ? 'tanpa SI' : 'normal') + ')');
    }
    const S = getSettings_();
    if (String(S.PERIODE_DARI || 'MINTA').toUpperCase() === 'NOTA') {
      const g = S.GABUNG_HARI === '' || S.GABUNG_HARI == null ? 3 : Number(S.GABUNG_HARI);
      r[1] = periode_(n.tanggal, g);
      sh.getRange(rowNo, 2).setValue(r[1]);
    }
    sh.getRange(rowNo, 3, 1, 3).setValues([[r[2], r[3], r[4]]]);
    sh.getRange(rowNo, 15, 1, 2).setValues([[r[14], r[15]]]);
    sh.getRange(rowNo, 25).setValue(r[24]);

    // sinkron nomor nota di STOK_KEY
    const st = ss_().getSheetByName(CFG.SH_STOK);
    if (st && st.getLastRow() > 1) {
      const keys = st.getRange(2, 1, st.getLastRow() - 1, 1).getValues();
      const j = keys.findIndex(x => String(x[0]).toUpperCase() === String(r[6]).toUpperCase());
      if (j >= 0) st.getRange(j + 2, 7).setValue(nota);
    }

    const log = ss_().getSheetByName('LOG_EDIT') || mkSheet_('LOG_EDIT', ['WAKTU', 'ADMIN', 'CABANG', 'PRODUCT KEY', 'PERUBAHAN'], '#EE7A12');
    log.appendRow([new Date(), u.nama, cab, r[6], perubahan.join(' | ')]);
    bustCache_();
    return { ok: true, row: rowObj_(cab, r, rowNo, nikMap_()), berubah: true };
  } finally {
    lock.releaseLock();
  }
}

function numOr0_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }

function getRekap(token, f) {
  const u = auth_(token);
  if (u.role !== 'ADMIN') throw new Error('Khusus admin.');
  return queryRekap_(f || {});
}

/* ============================ STOK ============================ */

function isAvail_(r) {
  const st = String(r[1]).trim().toUpperCase();
  return String(r[0]).trim() !== '' && (st === '' || st === 'TERSEDIA');
}

function alokasi_(r) {
  const a = String(r[2]).trim().toUpperCase();
  return CFG.CABANG[a] ? a : 'UMUM';
}

function countStock_() {
  const c = { UMUM: 0, TERPAKAI: 0, total: 0 };
  Object.keys(CFG.CABANG).forEach(k => c[k] = 0);
  const sh = ss_().getSheetByName(CFG.SH_STOK);
  if (!sh || sh.getLastRow() < 2) return c;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
    if (!String(r[0]).trim()) return;
    if (isAvail_(r)) { c[alokasi_(r)]++; c.total++; }
    else if (String(r[1]).toUpperCase() === 'TERPAKAI') c.TERPAKAI++;
  });
  return c;
}

/** Ambil semua key dari teks bebas, misal hasil copy chat WA:
 *  "[11.43, 24/9/2026] Office Eternal Solution: FTPTR-XJ3DJ-VH7R4-XTV38-MCBN9" -> "FTPTR-XJ3DJ-VH7R4-XTV38-MCBN9" */
function extractKeys_(text) {
  const keys = [];
  let tanpaKey = 0;
  String(text || '').split(/\r?\n/).forEach(line => {
    if (!line.trim()) return;
    const found = line.toUpperCase().replace(/[\u2010-\u2015]/g, '-').match(KEY_RE);
    if (found) found.forEach(k => keys.push(k)); else tanpaKey++;
  });
  return { keys: keys, tanpaKey: tanpaKey };
}

/** Paste langsung di sheet STOK_KEY kolom A juga otomatis dirapikan */
function onEdit(e) {
  try {
    const nm = e.range.getSheet().getName();
    if (nm === CFG.SH_STOK || nm.indexOf('REKAP_') === 0) bustCache_();
  } catch (err) {}
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SH_STOK || e.range.getColumn() !== 1 || e.range.getRow() < 2) return;
    const rg = sh.getRange(e.range.getRow(), 1, e.range.getNumRows(), 6);
    const vals = rg.getValues();
    let changed = false;
    vals.forEach(r => {
      const raw = String(r[0] || '');
      if (!raw) return;
      const m = raw.toUpperCase().replace(/[\u2010-\u2015]/g, '-').match(KEY_RE);
      if (m && m[0] !== raw) { r[0] = m[0]; changed = true; }
      if (m && !r[1]) { r[1] = 'TERSEDIA'; changed = true; }
      if (m && !r[2]) { r[2] = 'UMUM'; changed = true; }
      if (m && !r[5]) { r[5] = new Date(); changed = true; }
    });
    if (changed) rg.setValues(vals);
  } catch (err) {}
}

function addStock(token, p) {
  const u = auth_(token);
  if (u.role !== 'ADMIN') throw new Error('Hanya admin yang bisa menambah stok.');

  const ext = extractKeys_(p.keys);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName(CFG.SH_STOK);
    const existing = {};
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
        .forEach(r => existing[String(r[0]).toUpperCase().trim()] = true);
    }
    const valid = [], dup = [];
    const seen = {};
    ext.keys.forEach(k => {
      if (existing[k] || seen[k]) dup.push(k);
      else { seen[k] = true; valid.push(k); }
    });
    if (valid.length) {
      const now = new Date();
      const rows = valid.map(k => [k, 'TERSEDIA', 'UMUM', '', '', now, '', '', '']);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, STOK_HEADERS.length).setValues(rows);
    }
    const stok = countStock_();
    if (valid.length) bustCache_();
    return { masuk: valid.length, tanpaKey: ext.tanpaKey, duplikat: dup, stok: stok, stokCabang: stokCabang_(stok) };
  } finally {
    lock.releaseLock();
  }
}

/* ============================ PERMINTAAN KEY ============================ */

function submitRequest(token, f) {
  const u = auth_(token);
  const bebasCabang = u.cabang === 'ALL' || u.role === 'ADMIN';
  const cab = bebasCabang ? String(f.cabang || '').toUpperCase() : u.cabang;
  if (!CFG.CABANG[cab]) throw new Error('Pilih cabang dulu.');

  const namaUser = clean_(f.namaUser).toUpperCase();
  const adaKet = f.adaKet === true;
  const ket = adaKet ? clean_(f.ketText).toUpperCase() : '';
  if (!namaUser) throw new Error('Nama user wajib diisi.');
  if (adaKet && !ket) throw new Error('Keterangan dicentang, tulis keterangannya dulu.');
  const qty = Math.floor(Number(f.qty || 1));
  if (!(qty >= 1 && qty <= CFG.MAX_QTY)) throw new Error('Qty harus 1 sampai ' + CFG.MAX_QTY + '.');

  const kodeSales = clean_(f.kodeSales).toUpperCase();
  const salesObj = (getSalesMaster_([cab])[cab] || []).find(s => s.kode === kodeSales);
  if (!salesObj) throw new Error('Pilih nama sales dari daftar.');

  const nota = normNota_(f.nota);
  const n = parseNota_(nota);
  if (!n) throw new Error('Format nomor nota belum sesuai.');
  if (n.cabang !== cab) throw new Error('Kode ' + n.kode + ' di nota bukan kode cabang ' + CFG.CABANG[cab] + '.');

  const S = getSettings_();
  const now = new Date();
  const harga = hargaBerlaku_(now);
  const nh = nilaiHarga_(harga, !!ket);
  const tglPeriode = String(S.PERIODE_DARI || 'MINTA').toUpperCase() === 'NOTA' ? n.tanggal : now;
  const gabung = S.GABUNG_HARI === '' || S.GABUNG_HARI == null ? 3 : Number(S.GABUNG_HARI);
  const keys = [];
  let firstRow, stokBaru;

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const rekap = ss_().getSheetByName('REKAP_' + cab) || mkRekap_('REKAP_' + cab);

    // Nota sama: wajib dijawab di layar (YA = 1 nota beberapa unit, TIDAK = wajib keterangan)
    const dup = infoNota_(rekap, nota);
    if (dup.jumlah > 0) {
      if (f.dupJawab !== 'YA' && f.dupJawab !== 'TIDAK') throw new Error('NOTA_DUP');
      if (f.dupJawab === 'TIDAK' && !ket) throw new Error('Nota sudah pernah dipakai. Centang keterangan dan tulis alasannya (misal replace unit).');
    }

    const sh = ss_().getSheetByName(CFG.SH_STOK);
    const data = sh.getLastRow() > 1 ? sh.getRange(1, 1, sh.getLastRow(), 5).getValues() : [[]];
    const pick = [];
    for (let i = 1; i < data.length && pick.length < qty; i++) if (isAvail_(data[i]) && alokasi_(data[i]) === cab) pick.push(i);
    for (let i = 1; i < data.length && pick.length < qty; i++) if (isAvail_(data[i]) && alokasi_(data[i]) === 'UMUM') pick.push(i);
    if (pick.length < qty) {
      throw new Error(pick.length === 0 ? 'Stok key sedang habis. Hubungi admin untuk isi stok.'
        : 'Stok key hanya tersisa ' + pick.length + ', tidak cukup untuk ' + qty + ' key.');
    }

    firstRow = rekap.getLastRow() + 1;
    if (firstRow + qty > rekap.getMaxRows()) rekap.insertRowsAfter(rekap.getMaxRows(), Math.max(200, qty));
    const rows = [], fTotal = [], fSisa = [];
    pick.forEach((idx, j) => {
      const key = String(data[idx][0]).trim().toUpperCase();
      keys.push(key);
      data[idx][1] = 'TERPAKAI';
      sh.getRange(idx + 1, 2).setValue('TERPAKAI');
      sh.getRange(idx + 1, 7, 1, 3).setValues([[nota, u.nama + ' (' + cab + ')', now]]);
      const r = firstRow + j;
      rows.push([
        r - 1, periode_(tglPeriode, gabung), n.tanggal, nota, namaUser, S.NAMA_BARANG || 'OFFICE PRO PLUS 2021', key,
        nh.jual, nh.modal, data[idx][4] || '', nh.si, '', '', '', salesObj.nama, salesObj.kode, nh.siSales, '',
        nh.siTeksu, '', '', u.nama, u.email, now, ket
      ]);
      fTotal.push(['=IF(AND(I' + r + '="",K' + r + '=""),"",N(I' + r + ')+N(K' + r + '))']);
      fSisa.push(['=IF(K' + r + '="","",N(K' + r + ')-N(Q' + r + ')-N(S' + r + '))']);
    });
    rekap.getRange(firstRow, 1, qty, REKAP_HEADERS.length).setValues(rows);
    rekap.getRange(firstRow, 13, qty, 1).setFormulas(fTotal);
    rekap.getRange(firstRow, 21, qty, 1).setFormulas(fSisa);

    stokBaru = { UMUM: 0, TERPAKAI: 0, total: 0 };
    Object.keys(CFG.CABANG).forEach(k => stokBaru[k] = 0);
    for (let i = 1; i < data.length; i++) if (isAvail_(data[i])) { stokBaru[alokasi_(data[i])]++; stokBaru.total++; }
    SpreadsheetApp.flush();
    bustCache_();
  } finally {
    lock.releaseLock();
  }

  return {
    keys: keys, nota: nota, cabang: cab, row: firstRow, email: u.email,
    stok: stokBaru, stokCabang: stokCabang_(stokBaru)
  };
}

/** Cek nota di layar (dipanggil saat nota ditempel, sebelum tombol Minta key) */
function checkNota(token, cabang, nota) {
  const u = auth_(token);
  const cab = (u.cabang === 'ALL' || u.role === 'ADMIN') ? String(cabang || '').toUpperCase() : u.cabang;
  if (!CFG.CABANG[cab]) return { jumlah: 0 };
  const sh = ss_().getSheetByName('REKAP_' + cab);
  if (!sh) return { jumlah: 0 };
  const info = infoNota_(sh, normNota_(nota));
  info.nota = normNota_(nota);
  return info;
}

function infoNota_(sh, nota) {
  const last = sh.getLastRow();
  const out = { jumlah: 0, terakhir: '', user: '' };
  if (last < 2) return out;
  const col = sh.getRange(2, 4, last - 1, 1).getValues();
  let lastIdx = -1;
  col.forEach((x, i) => { if (String(x[0]).toUpperCase().trim() === nota) { out.jumlah++; lastIdx = i; } });
  if (lastIdx >= 0) {
    const r = sh.getRange(lastIdx + 2, 1, 1, REKAP_HEADERS.length).getValues()[0];
    out.user = String(r[4]);
    out.terakhir = r[23] instanceof Date ? Utilities.formatDate(r[23], tz_(), 'dd/MM/yyyy HH:mm') : '';
  }
  return out;
}

/** Dipanggil dari layar setelah key tampil: kirim 1 email berisi semua key + cek stok menipis */
function sendRequestEmail(token, ref) {
  const u = auth_(token);
  const cab = String(ref.cabang || '').toUpperCase();
  if (!CFG.CABANG[cab]) throw new Error('Cabang tidak dikenal.');
  const want = (ref.keys || []).map(String);
  if (!want.length) throw new Error('Tidak ada key.');
  const sh = ss_().getSheetByName('REKAP_' + cab);
  let found = [];
  const start = Number(ref.row);
  if (start >= 2 && start + want.length - 1 <= sh.getLastRow()) {
    found = sh.getRange(start, 1, want.length, REKAP_HEADERS.length).getValues().filter(r => want.indexOf(String(r[6])) > -1);
  }
  if (found.length !== want.length) { // baris bergeser: cari lewat kolom key saja
    const col = sh.getRange(2, 7, sh.getLastRow() - 1, 1).getValues();
    found = [];
    col.forEach((x, i) => { if (want.indexOf(String(x[0])) > -1) found.push(i + 2); });
    found = found.map(rn => sh.getRange(rn, 1, 1, REKAP_HEADERS.length).getValues()[0]);
  }
  if (!found.length) throw new Error('Data permintaan tidak ditemukan.');
  if (found.some(r => String(r[21]) !== u.nama)) throw new Error('Data permintaan bukan milik akun ini.');
  const r = found[0];
  let sent = false;
  if (u.email) {
    sendKeyEmail_(u, { keys: found.map(x => String(x[6])), nota: r[3], namaUser: r[4], sales: r[14] + ' (' + r[15] + ')',
      cabang: cab, ket: r[24], barang: r[5] });
    sent = true;
  }
  try { checkLowStock_(cab, getSettings_()); } catch (e) {}
  return { sent: sent, email: u.email };
}

function normNota_(s) {
  return String(s || '').toUpperCase().replace(/\s+/g, '').replace(/_/g, '-').trim();
}

function kodeNotaMap_() {
  const m = {};
  Object.keys(CFG.CABANG).forEach(cab => ((CFG.KODE_NOTA && CFG.KODE_NOTA[cab]) || [cab]).forEach(k => m[k] = cab));
  return m;
}

function parseNota_(nota) {
  const map = kodeNotaMap_();
  const m = nota.match(new RegExp('^JL-(' + Object.keys(map).join('|') + ')(-KIOSK)?-(\\d{2})(\\d{2})(\\d{2})-(\\d{4,5})$'));
  if (!m) return null;
  const y = 2000 + Number(m[3]), mo = Number(m[4]), d = Number(m[5]);
  const tgl = new Date(y, mo - 1, d);
  if (tgl.getMonth() !== mo - 1 || tgl.getDate() !== d) return null;
  return { kode: m[1], cabang: map[m[1]], kiosk: !!m[2], tanggal: tgl, urut: m[6] };
}

/**
 * Periode mingguan Senin–Minggu dalam 1 bulan.
 * Potongan minggu di awal/akhir bulan yang <= `gabung` hari digabung ke minggu sebelahnya.
 * Contoh (gabung 3): 24/09/2026 -> "2026_09 21-30 SEPTEMBER 2026"
 *                    02/08/2026 -> "2026_08 01-09 AGUSTUS 2026"
 */
function periode_(d, gabung) {
  gabung = isNaN(gabung) ? 3 : gabung;
  const y = d.getFullYear(), m = d.getMonth();
  const last = new Date(y, m + 1, 0).getDate();
  const seg = [];
  let s = 1;
  for (let day = 1; day <= last; day++) {
    const isSunday = new Date(y, m, day).getDay() === 0;
    if (isSunday || day === last) { seg.push([s, day]); s = day + 1; }
  }
  const len = x => x[1] - x[0] + 1;
  if (seg.length > 1 && len(seg[seg.length - 1]) <= gabung) {
    const tail = seg.pop(); seg[seg.length - 1][1] = tail[1];
  }
  if (seg.length > 1 && len(seg[0]) <= gabung) {
    const head = seg.shift(); seg[0][0] = head[0];
  }
  const tgl = d.getDate();
  const p = seg.find(x => tgl >= x[0] && tgl <= x[1]);
  const p2 = x => ('0' + x).slice(-2);
  return y + '_' + p2(m + 1) + ' ' + p2(p[0]) + '-' + p2(p[1]) + ' ' + BULAN[m] + ' ' + y;
}

/* ============================ EMAIL ============================ */

function sendKeyEmail_(u, x) {
  const barang = x.barang || 'Office Pro Plus 2021';
  const keys = x.keys || [x.key];
  const keyHtml = keys.map((k, i) =>
    '<div style="background:#FFF1E3;border:1px dashed #EE7A12;padding:12px;margin-bottom:8px;border-radius:8px;text-align:center;' +
    'font-family:Consolas,monospace;font-size:18px;letter-spacing:1px;font-weight:bold">' +
    (keys.length > 1 ? '<span style="font-size:12px;color:#A94B02;font-family:Arial">Key ' + (i + 1) + '</span><br>' : '') + k + '</div>').join('');
  const html =
    '<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;color:#3A2A1A">' +
    '<div style="background:#EE7A12;color:#fff;padding:18px 22px;border-radius:10px 10px 0 0">' +
    '<div style="font-size:13px;opacity:.85">ELS.ID ' + CFG.CABANG[x.cabang] + '</div>' +
    '<div style="font-size:19px;font-weight:bold">Product key ' + barang + (keys.length > 1 ? ' (' + keys.length + ' key)' : '') + '</div></div>' +
    '<div style="border:1px solid #F3DCC6;border-top:0;padding:22px;border-radius:0 0 10px 10px">' + keyHtml +
    '<table style="width:100%;margin-top:12px;font-size:14px;border-collapse:collapse">' +
    row_('Nomor nota', x.nota) + row_('Nama user', x.namaUser) + row_('Sales', x.sales) +
    row_('Jumlah key', keys.length) + row_('Diminta oleh', u.nama) + (x.ket ? row_('Keterangan', x.ket) : '') +
    '</table><p style="font-size:12px;color:#8C7866;margin-top:18px">Key sudah tercatat atas nota di atas dan tidak bisa dipakai ulang. Simpan email ini sebagai arsip.</p></div></div>';
  MailApp.sendEmail({
    to: u.email,
    subject: '[Key Office] ' + x.nota + ' — ' + x.namaUser + (keys.length > 1 ? ' (' + keys.length + ' key)' : ''),
    htmlBody: html,
    name: 'ELS.ID Key Office'
  });
}

function row_(a, b) {
  return '<tr><td style="padding:6px 0;color:#5B6B7C;width:130px">' + a + '</td>' +
    '<td style="padding:6px 0;font-weight:bold">' + b + '</td></tr>';
}

function checkLowStock_(cab, S) {
  const min = Number(S.MIN_STOK || 10);
  const to = String(S.EMAIL_ADMIN || '').trim();
  if (!to) return;
  const c = countStock_();
  const sisa = c[cab] + c.UMUM;
  if (sisa > min) return;
  const props = PropertiesService.getScriptProperties();
  const flag = 'ALERT_' + cab + '_' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd');
  if (props.getProperty(flag)) return;
  props.setProperty(flag, '1');
  MailApp.sendEmail(to, '⚠️ Stok key ' + CFG.CABANG[cab] + ' tinggal ' + sisa,
    'Stok key Office untuk cabang ' + CFG.CABANG[cab] + ' tersisa ' + sisa + ' (khusus cabang ' + c[cab] +
    ', umum ' + c.UMUM + ').\nSegera tambah stok lewat panel admin atau sheet STOK_KEY.');
}

/* ============================ UTIL ============================ */

function getSettings_() {
  const sh = ss_().getSheetByName(CFG.SH_SET);
  const o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) o[String(r[0]).trim()] = r[1]; });
  return o;
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return Session.getScriptTimeZone() || 'Asia/Jakarta'; }
function clean_(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }
function num_(v) { const n = Number(String(v).replace(/[^\d.-]/g, '')); return isNaN(n) || v === '' ? '' : n; }
