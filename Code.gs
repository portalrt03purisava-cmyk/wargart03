/** Portal Warga RT 03 Puri Sava — Code.gs (script TERIKAT ke Google Sheets) */
const TZ = 'Asia/Jakarta', TTL = 21600, MAX_GAGAL = 5;
const SKEMA = {
  Users: ['id','nama','blok','noRumah','hp','salt','hash','role','status','dibuat','loginTerakhir','catatan'],
  Pengumuman: ['id','tanggal','judul','isi','oleh'],
  Iuran: ['id','userId','nama','periode','jumlah','tanggalBayar','metode','dicatatOleh','catatan'],
  Kas: ['id','tanggal','tipe','kategori','jumlah','keterangan','dicatatOleh'],
  Log: ['waktu','userId','aksi','detail']
};
const K = 'KETUA', B = 'BENDAHARA', W = 'WARGA', KB = [K, B];
// Matriks hak akses — dicek di SERVER pada setiap panggilan
const ACL = {
  me:'*', logout:'*', ubahSandi:'*', pengumuman:'*', iuranSaya:'*',
  tambahPengumuman:[K], hapusPengumuman:[K], menunggu:[K], verifikasi:[K], warga:[K],
  ubahWarga:[K], resetSandi:[K], logAktivitas:[K],
  ringkasan:KB, listKas:KB, listIuran:KB, wargaAktif:KB,
  catatIuran:[B], tambahKas:[B]
};
const MUTASI = ['ubahSandi','tambahPengumuman','hapusPengumuman','verifikasi','ubahWarga','resetSandi','catatIuran','tambahKas'];

class Err extends Error { constructor(m, auth) { super(m); this.auth = !!auth; } }

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Portal Warga RT 03 Puri Sava')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- Helper ---------- */
const props = () => PropertiesService.getScriptProperties();
const db = n => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
const now = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
const today = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const uid = () => Utilities.getUuid().slice(0, 8);
const hex = b => b.map(x => ((x & 255) + 256).toString(16).slice(1)).join('');
const sha = s => hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8));
const bersih = (s, max, multi) => String(s == null ? '' : s)
  .replace(multi ? /[\u0000-\u0009\u000b-\u001f\u007f]/g : /[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max || 200);
const aman = v => (typeof v === 'string' && /^[=+\-@']/.test(v)) ? "'" + v : v;  // anti formula-injection
const baca = v => typeof v === 'string' ? v.replace(/^'/, '') : v;

function semua(n) {
  const v = db(n).getDataRange().getValues(), h = v.shift();
  return v.filter(r => r[0] !== '' || n === 'Log').map((r, i) => {
    const o = { _baris: i + 2 }; h.forEach((k, j) => o[k] = baca(r[j])); return o;
  });
}
function tambah(n, o) { db(n).appendRow(SKEMA[n].map(k => aman(o[k] == null ? '' : o[k]))); }
function ubah(n, baris, o) {
  const s = db(n);
  Object.keys(o).forEach(k => s.getRange(baris, SKEMA[n].indexOf(k) + 1).setValue(aman(o[k])));
}
function cari(n, id) {
  const r = semua(n).find(x => x.id === String(id || ''));
  if (!r) throw new Err('Data tidak ditemukan.');
  return r;
}
function catat(u, aksi, d) { tambah('Log', { waktu: now(), userId: u, aksi, detail: bersih(d, 300) }); }
function kunci(fn) {
  const l = LockService.getScriptLock(); l.waitLock(15000);
  try { return fn(); } finally { l.releaseLock(); }
}
function normHp(h) {
  let d = String(h || '').replace(/\D/g, '');
  if (d.indexOf('62') === 0) d = '0' + d.slice(2);
  if (!/^08\d{8,12}$/.test(d)) throw new Err('Nomor HP tidak valid (contoh 081234567890).');
  return d;
}
function hashPw(pw, salt) {
  let h = salt + '|' + pw + '|' + props().getProperty('PEPPER');
  for (let i = 0; i < 1500; i++) h = sha(h + i);
  return h;
}
function samaAman(a, b) {
  if (a.length !== b.length) return false;
  let r = 0; for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}
function cekSandi(p) {
  if (typeof p !== 'string' || p.length < 8 || p.length > 64 || !/[A-Za-z]/.test(p) || !/\d/.test(p))
    throw new Err('Kata sandi 8–64 karakter, wajib mengandung huruf dan angka.');
}
const profil = u => ({ id: u.id, nama: u.nama, blok: u.blok, noRumah: u.noRumah, role: u.role });
const angka = (v, min, max, label) => {
  const n = Number(v);
  if (!Number.isInteger(n) || n < min || n > max) throw new Err(label + ' tidak valid.');
  return n;
};

/* ---------- Pendaftaran, login, sesi ---------- */
function register(p) {
  const c = CacheService.getScriptCache(), k = 'REG' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHH');
  const n = Number(c.get(k) || 0);
  if (n >= 15) throw new Err('Terlalu banyak pendaftaran. Coba lagi nanti.');
  c.put(k, String(n + 1), 3600);
  const nama = bersih(p.nama, 60), blok = bersih(p.blok, 10).toUpperCase(), no = bersih(p.noRumah, 10), hp = normHp(p.hp);
  if (nama.length < 3 || !blok || !no) throw new Err('Nama, blok, dan nomor rumah wajib diisi.');
  cekSandi(p.sandi);
  const salt = Utilities.getUuid(), hash = hashPw(p.sandi, salt);
  return kunci(() => {
    if (semua('Users').some(u => u.hp === hp)) throw new Err('Nomor HP sudah terdaftar.');
    tambah('Users', { id: uid(), nama, blok, noRumah: no, hp, salt, hash, role: W, status: 'MENUNGGU', dibuat: now() });
    return 'Pendaftaran terkirim. Akun aktif setelah diverifikasi Ketua RT.';
  });
}
function login(p) {
  const hp = normHp(p.hp), c = CacheService.getScriptCache(), fk = 'F' + hp, g = Number(c.get(fk) || 0);
  if (g >= MAX_GAGAL) throw new Err('Terlalu banyak percobaan. Coba lagi 15 menit lagi.');
  const u = semua('Users').find(x => x.hp === hp);
  const ok = samaAman(hashPw(String(p.sandi || '').slice(0, 64), u ? u.salt : 'x'), u ? u.hash : sha('dummy'));
  if (!u || !ok) { c.put(fk, String(g + 1), 900); throw new Err('Nomor HP atau kata sandi salah.'); }
  if (u.status === 'MENUNGGU') throw new Err('Akun Anda masih menunggu verifikasi Ketua RT.');
  if (u.status !== 'AKTIF') throw new Err('Akun tidak aktif. Hubungi Ketua RT.');
  c.remove(fk);
  const t = Utilities.getUuid() + Utilities.getUuid();
  c.put('S' + sha(t), u.id, TTL);
  ubah('Users', u._baris, { loginTerakhir: now() });
  catat(u.id, 'LOGIN', '');
  return { token: t, profil: profil(u) };
}
function auth(t) {
  const gagal = () => new Err('Sesi berakhir. Silakan masuk kembali.', true);
  if (typeof t !== 'string' || t.length < 60) throw gagal();
  const id = CacheService.getScriptCache().get('S' + sha(t));
  if (!id) throw gagal();
  const u = semua('Users').find(x => x.id === id);   // peran & status dibaca ulang tiap request
  if (!u || u.status !== 'AKTIF') throw gagal();
  return u;
}

/* ---------- Dispatcher tunggal ---------- */
function api(token, action, p) {
  try {
    p = (p && typeof p === 'object') ? p : {};
    if (action === 'register') return { ok: true, data: register(p) };
    if (action === 'login') return { ok: true, data: login(p) };
    const acl = ACL[action];
    if (!acl || !Object.prototype.hasOwnProperty.call(A, action)) throw new Err('Aksi tidak dikenal.');
    const u = auth(token);
    if (acl !== '*' && acl.indexOf(u.role) < 0) { catat(u.id, 'DITOLAK', action); throw new Err('Anda tidak berwenang.'); }
    const jalan = () => A[action](u, p, token);
    return { ok: true, data: MUTASI.indexOf(action) >= 0 ? kunci(jalan) : jalan() };
  } catch (e) {
    if (e instanceof Err) return { ok: false, error: e.message, auth: e.auth };
    console.error(e);
    return { ok: false, error: 'Terjadi kesalahan server.' };
  }
}

/* ---------- Aksi ---------- */
const A = {
  me: u => profil(u),
  logout: (u, p, t) => { CacheService.getScriptCache().remove('S' + sha(t)); return 'ok'; },
  ubahSandi: (u, p) => {
    if (!samaAman(hashPw(String(p.lama || ''), u.salt), u.hash)) throw new Err('Kata sandi lama salah.');
    cekSandi(p.baru);
    const salt = Utilities.getUuid();
    ubah('Users', u._baris, { salt, hash: hashPw(p.baru, salt) });
    catat(u.id, 'UBAH_SANDI', ''); return 'Kata sandi diperbarui.';
  },
  pengumuman: () => semua('Pengumuman').reverse().slice(0, 30),
  tambahPengumuman: (u, p) => {
    const judul = bersih(p.judul, 100), isi = bersih(p.isi, 1500, true);
    if (!judul || !isi) throw new Err('Judul dan isi wajib diisi.');
    tambah('Pengumuman', { id: uid(), tanggal: today(), judul, isi, oleh: u.nama });
    catat(u.id, 'PENGUMUMAN', judul); return 'Pengumuman terbit.';
  },
  hapusPengumuman: (u, p) => {
    const r = cari('Pengumuman', p.id); db('Pengumuman').deleteRow(r._baris);
    catat(u.id, 'HAPUS_PENGUMUMAN', r.judul); return 'Dihapus.';
  },
  iuranSaya: u => semua('Iuran').filter(x => x.userId === u.id).reverse()
    .map(i => ({ periode: i.periode, jumlah: Number(i.jumlah), tanggalBayar: i.tanggalBayar, metode: i.metode })),
  menunggu: () => semua('Users').filter(x => x.status === 'MENUNGGU')
    .map(x => ({ id: x.id, nama: x.nama, blok: x.blok, noRumah: x.noRumah, hp: x.hp, dibuat: x.dibuat })),
  verifikasi: (u, p) => {
    const t = cari('Users', p.id);
    if (t.status !== 'MENUNGGU') throw new Err('Akun ini sudah diproses.');
    ubah('Users', t._baris, { status: p.setuju ? 'AKTIF' : 'DITOLAK', catatan: bersih(p.catatan, 100) });
    catat(u.id, p.setuju ? 'VERIF_SETUJU' : 'VERIF_TOLAK', t.nama + ' ' + t.hp);
    return p.setuju ? 'Warga diaktifkan.' : 'Pendaftaran ditolak.';
  },
  warga: () => semua('Users').map(x => ({ id: x.id, nama: x.nama, blok: x.blok, noRumah: x.noRumah, hp: x.hp, role: x.role, status: x.status })),
  ubahWarga: (u, p) => {
    const t = cari('Users', p.id);
    if (t.id === u.id || t.role === K) throw new Err('Akun Ketua tidak dapat diubah dari sini.');
    if (p.role !== undefined) {
      if ([W, B].indexOf(p.role) < 0 || t.status !== 'AKTIF') throw new Err('Peran tidak valid.');
      if (p.role === B) semua('Users').filter(x => x.role === B && x.id !== t.id).forEach(x => ubah('Users', x._baris, { role: W }));
      ubah('Users', t._baris, { role: p.role }); catat(u.id, 'UBAH_PERAN', t.nama + ' → ' + p.role);
    }
    if (p.status !== undefined) {
      if (['AKTIF', 'NONAKTIF'].indexOf(p.status) < 0 || ['AKTIF', 'NONAKTIF'].indexOf(t.status) < 0) throw new Err('Status tidak valid.');
      ubah('Users', t._baris, { status: p.status, role: p.status === 'NONAKTIF' ? W : t.role });
      catat(u.id, 'UBAH_STATUS', t.nama + ' → ' + p.status);
    }
    return 'Tersimpan.';
  },
  resetSandi: (u, p) => {
    const t = cari('Users', p.id);
    if (t.role === K) throw new Err('Tidak diizinkan.');
    const tmp = Utilities.getUuid().replace(/-/g, '').slice(0, 10) + 'a1', salt = Utilities.getUuid();
    ubah('Users', t._baris, { salt, hash: hashPw(tmp, salt) });
    catat(u.id, 'RESET_SANDI', t.nama); return tmp;
  },
  logAktivitas: () => semua('Log').reverse().slice(0, 80),
  wargaAktif: () => semua('Users').filter(x => x.status === 'AKTIF')
    .map(x => ({ id: x.id, nama: x.nama, blok: x.blok, noRumah: x.noRumah })),
  ringkasan: (u, p) => {
    const per = /^\d{4}-\d{2}$/.test(p.periode) ? p.periode : today().slice(0, 7);
    let masuk = 0, keluar = 0;
    semua('Kas').forEach(k => { if (k.tipe === 'MASUK') masuk += Number(k.jumlah); else keluar += Number(k.jumlah); });
    return { saldo: masuk - keluar, masuk, keluar, periode: per,
      warga: semua('Users').filter(x => x.status === 'AKTIF').length,
      sudah: semua('Iuran').filter(x => x.periode === per).length };
  },
  listKas: () => semua('Kas').reverse().slice(0, 100).map(k => ({ tanggal: k.tanggal, tipe: k.tipe, kategori: k.kategori, jumlah: Number(k.jumlah), keterangan: k.keterangan, oleh: k.dicatatOleh })),
  listIuran: (u, p) => {
    if (!/^\d{4}-\d{2}$/.test(p.periode)) throw new Err('Periode tidak valid.');
    const sudah = semua('Iuran').filter(x => x.periode === p.periode), ids = sudah.map(x => x.userId);
    return {
      sudah: sudah.map(x => ({ nama: x.nama, jumlah: Number(x.jumlah), tanggalBayar: x.tanggalBayar, metode: x.metode })),
      belum: semua('Users').filter(x => x.status === 'AKTIF' && ids.indexOf(x.id) < 0).map(x => ({ nama: x.nama, blok: x.blok, noRumah: x.noRumah }))
    };
  },
  catatIuran: (u, p) => {
    const w = cari('Users', p.userId), per = String(p.periode);
    if (w.status !== 'AKTIF') throw new Err('Warga tidak aktif.');
    if (!/^\d{4}-\d{2}$/.test(per)) throw new Err('Periode tidak valid.');
    if (['Tunai', 'Transfer'].indexOf(p.metode) < 0) throw new Err('Metode tidak valid.');
    const jml = angka(p.jumlah, 1, 10000000, 'Jumlah');
    if (semua('Iuran').some(x => x.userId === w.id && x.periode === per)) throw new Err('Iuran periode ini sudah tercatat.');
    tambah('Iuran', { id: uid(), userId: w.id, nama: w.nama, periode: per, jumlah: jml, tanggalBayar: today(), metode: p.metode, dicatatOleh: u.nama, catatan: bersih(p.catatan, 100) });
    tambah('Kas', { id: uid(), tanggal: today(), tipe: 'MASUK', kategori: 'Iuran', jumlah: jml, keterangan: 'Iuran ' + per + ' - ' + w.nama, dicatatOleh: u.nama });
    catat(u.id, 'IURAN', w.nama + ' ' + per + ' ' + jml); return 'Iuran tercatat.';
  },
  tambahKas: (u, p) => {
    if (['MASUK', 'KELUAR'].indexOf(p.tipe) < 0) throw new Err('Jenis tidak valid.');
    const jml = angka(p.jumlah, 1, 100000000, 'Jumlah'), kat = bersih(p.kategori, 40), ket = bersih(p.keterangan, 200);
    if (!kat || !ket) throw new Err('Kategori dan keterangan wajib diisi.');
    tambah('Kas', { id: uid(), tanggal: today(), tipe: p.tipe, kategori: kat, jumlah: jml, keterangan: ket, dicatatOleh: u.nama });
    catat(u.id, 'KAS_' + p.tipe, kat + ' ' + jml); return 'Transaksi tercatat. (Koreksi dilakukan dengan transaksi pembalik.)';
  }
};

/* ---------- Setup (dijalankan manual dari editor, SEKALI) ---------- */
function setupAwal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), me = Session.getEffectiveUser();
  Object.keys(SKEMA).forEach(n => {
    const h = SKEMA[n], s = ss.getSheetByName(n) || ss.insertSheet(n);
    s.getRange(1, 1, s.getMaxRows(), h.length).setNumberFormat('@');
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    s.setFrozenRows(1);
    if (!s.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) {
      const pr = s.protect().setDescription('Dikelola skrip portal');
      pr.addEditor(me); pr.removeEditors(pr.getEditors().filter(e => e.getEmail() !== me.getEmail()));
      if (pr.canDomainEdit()) pr.setDomainEdit(false);
    }
  });
  ss.getSheets().filter(s => !SKEMA[s.getName()]).forEach(s => { if (ss.getSheets().length > 1) ss.deleteSheet(s); });
  if (!props().getProperty('PEPPER')) props().setProperty('PEPPER', Utilities.getUuid() + Utilities.getUuid());
}
/** Isi INIT_NAMA, INIT_HP, INIT_SANDI (opsional INIT_BLOK, INIT_NO) di Script Properties, lalu jalankan. */
function buatKetua() {
  const P = props(), nama = P.getProperty('INIT_NAMA'), pw = P.getProperty('INIT_SANDI');
  if (!nama || !pw) throw new Error('Isi INIT_NAMA, INIT_HP, INIT_SANDI di Script Properties dahulu.');
  cekSandi(pw);
  if (semua('Users').some(u => u.role === K)) throw new Error('Akun Ketua sudah ada.');
  const salt = Utilities.getUuid();
  tambah('Users', { id: uid(), nama: bersih(nama, 60), blok: P.getProperty('INIT_BLOK') || '-', noRumah: P.getProperty('INIT_NO') || '-',
    hp: normHp(P.getProperty('INIT_HP')), salt, hash: hashPw(pw, salt), role: K, status: 'AKTIF', dibuat: now() });
  ['INIT_NAMA', 'INIT_HP', 'INIT_SANDI', 'INIT_BLOK', 'INIT_NO'].forEach(k => P.deleteProperty(k));
}
