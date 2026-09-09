// ============================================================
// Code.gs — Google Apps Script untuk Sobat Minisoccer Portal
// Deploy sebagai: Web App | Execute as: Me | Access: Anyone
// ============================================================

const SS_ID   = '1z6nbrxSHzDnW3WNVBWX4mP_f5rqJSh3Z27GJWftuq84'; // Sheet "Minisoccer_DB" ASLI (akun sbt.minisoccer) — udah di-share view ke intanktrav, strukturnya udah dicek cocok sama kode di bawah
const SS      = () => SpreadsheetApp.openById(SS_ID);

const SHEET = {
  PLAYER   : 'DB_Player',
  JERSEY   : 'Stok_Jersey',
  SHTM     : 'SHTM_Log',
  REKAP    : 'Rekap_Keuangan',
  DEPOSIT  : 'Deposit_Log',
  INVENTARIS: 'Inventaris',
  SCORER   : 'Rekap_Skorer', // tab baru — BELUM ada di sheet data riil kamu juga,
                              // dibikin otomatis sama setupSheets()
};

// ── CORS & Router ──────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';
  let result;
  try {
    switch(action) {
      case 'getPlayers':       result = getPlayers(); break;
      case 'getMatches':       result = getMatches(); break;
      case 'getMatch':         result = getMatch(e.parameter.id); break;
      case 'getStokJersey':    result = getStokJersey(); break;
      case 'getShtm':          result = getShtm(); break;
      case 'getDeposit':       result = getDeposit(); break;
      case 'getDashboard':     result = getDashboard(); break;
      case 'getConfig':        result = getConfig(); break;
      case 'getScorer':        result = getScorer(); break;
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const action  = payload.action || '';
  let result;
  try {
    switch(action) {
      case 'createMatch':       result = createMatch(payload); break;
      case 'addPlayer':         result = addPlayerToMatch(payload); break;
      case 'updatePlayer':      result = updatePlayerInMatch(payload); break;
      case 'deletePlayer':      result = deletePlayerFromMatch(payload); break;
      case 'addBiaya':          result = addBiaya(payload); break;
      case 'updateBiaya':       result = updateBiaya(payload); break;
      case 'deleteBiaya':       result = deleteBiaya(payload); break;
      case 'addGol':            result = addGol(payload); break;
      case 'deleteGol':         result = deleteGol(payload); break;
      case 'flagShtm':          result = flagShtm(payload); break;
      case 'useShtm':           result = useShtm(payload); break;
      case 'addDeposit':        result = addDeposit(payload); break;
      case 'useDeposit':        result = useDeposit(payload); break;
      case 'closeMatch':        result = closeMatch(payload); break;
      case 'updateStokJersey':  result = updateStokJersey(payload); break;
      case 'addPlayerDb':       result = addNewPlayer(payload); break;
      case 'updateConfig':      result = updateConfig(payload); break;
      case 'addScorer':         result = addScorer(payload); break;
      case 'updateScorer':      result = updateScorer(payload); break;
      case 'deleteScorer':      result = deleteScorer(payload); break;
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Helpers ────────────────────────────────────────────────
function getSheet(name) {
  return SS().getSheetByName(name);
}

function sheetToObjects(sheetName, headerRow) {
  const ws   = getSheet(sheetName);
  const data = ws.getDataRange().getValues();
  const hdrs = data[headerRow - 1].map(h => String(h).trim());
  const rows = [];
  for (let r = headerRow; r < data.length; r++) {
    const row = data[r];
    if (!row[0] && !row[1]) continue; // skip empty rows
    const obj = {};
    hdrs.forEach((h, i) => { obj[h] = row[i] === '' ? null : row[i]; });
    obj._row = r + 1; // 1-based row index for updates
    rows.push(obj);
  }
  return rows;
}

function toNum(v) {
  if (typeof v === 'number') return v;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function generateId(prefix) {
  const d  = new Date();
  const ts = Utilities.formatDate(d, 'Asia/Jakarta', 'yyyyMMdd');
  return prefix + '-' + ts + '-' + Math.floor(Math.random()*1000);
}

function today() {
  return Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMM yyyy');
}

// ── GET: Dashboard aggregates ───────────────────────────────
function getDashboard() {
  const players = getPlayers();
  const matches = getMatches();
  const shtm    = getShtm();
  const jersey  = getStokJersey();

  // Setiap baris di Rekap_Keuangan MEMANG match yang udah ditutup — cuma
  // closeMatch() yang nulis ke sheet ini, dan sheet ini nggak punya kolom
  // "Status Match" (itu cuma ada di TEMPLATE_Match/per-match sheet). Kode
  // lama nge-filter pakai m['Status Match'], yang selalu undefined utk
  // baris dari sini — jadi `closed` selalu kosong dan totalMatch selalu 0.
  const closed = matches;
  const last   = closed[closed.length - 1] || {};

  // Top scorer all-time
  const scorerMap = {};
  closed.forEach(m => {
    if (m._scorers) {
      m._scorers.forEach(s => {
        scorerMap[s.nama] = (scorerMap[s.nama] || 0) + (s.gol || 0);
      });
    }
  });
  const topScorer = Object.entries(scorerMap)
    .sort((a,b) => b[1]-a[1])
    .slice(0,10)
    .map(([nama,gol]) => ({nama,gol}));

  // Frequent players
  const freqMap = {};
  players.forEach(p => {
    freqMap[p['Nama Player']] = p['Total\nHadir'] || p['Total Hadir'] || 0;
  });
  const frequent = Object.entries(freqMap)
    .sort((a,b) => b[1]-a[1])
    .slice(0,9)
    .map(([nama,cnt]) => ({nama,cnt}));

  // Match terbaik (margin tertinggi, all-time — bukan cuma 12 terakhir)
  let bestMatch = null;
  closed.forEach(m => {
    const margin = toNum(m['Margin\n(Rp)'] || m['Margin']);
    if (!bestMatch || margin > bestMatch.margin) {
      bestMatch = { tgl: m['Match / Tanggal'] || '', tipe: m['Tipe Event'] || '', margin };
    }
  });

  // Delta kumulatif Malik/Filan dari N match terakhir (N=4, atau kurang
  // kalau match yang tercatat belum sampai 5) — buat subtext "Shared profit".
  const deltaN = Math.min(4, closed.length > 0 ? closed.length - 1 : 0);
  const refMatch = deltaN > 0 ? closed[closed.length - 1 - deltaN] : null;
  const kumulMalikNow  = toNum(last['Kumul Malik\n(Rp)'] || last['Kumul Malik']);
  const kumulFilanNow  = toNum(last['Kumul Filan\n(Rp)'] || last['Kumul Filan']);
  const deltaMalik = refMatch ? (kumulMalikNow - toNum(refMatch['Kumul Malik\n(Rp)'] || refMatch['Kumul Malik'])) : kumulMalikNow;
  const deltaFilan = refMatch ? (kumulFilanNow - toNum(refMatch['Kumul Filan\n(Rp)'] || refMatch['Kumul Filan'])) : kumulFilanNow;

  return {
    totalMatch  : closed.length,
    totalPlayer : players.length,
    lastMatchTgl: last['Match / Tanggal'] || '',
    kumulatifMalik: kumulMalikNow,
    kumulatifFilan: kumulFilanNow,
    deltaMalik, deltaFilan, deltaN,
    bestMatch,
    topScorer,
    frequent,
    // Riwayat SEMUA match (bukan cuma 12 terakhir) — biar chart "Margin per
    // match" bisa switch 12/24/Semua tanpa perlu fetch ulang.
    matches: closed.map(m => ({
      tgl   : m['Match / Tanggal'],
      jenis : m['Jenis\nGame'] || m['Jenis Game'],
      tipe  : m['Tipe Event'],
      margin: toNum(m['Margin\n(Rp)'] || m['Margin']),
      lap   : m['Lap.'],
      tahun : m['Tahun'] || '',
    }))
  };
}

// ── GET: Players ────────────────────────────────────────────
function getPlayers() {
  return sheetToObjects(SHEET.PLAYER, 2);
}

// ── GET: Matches (dari sheet Rekap_Keuangan + sheet per-match jika ada) ──
function getMatches() {
  return sheetToObjects(SHEET.REKAP, 2)
    .filter(m => m['No'] && String(m['No']).match(/^\d+$/));
}

function getMatch(matchId) {
  // Cari sheet dengan nama match (misal: "Match-20260920-001")
  const ss = SS();
  const ws = ss.getSheetByName(matchId);
  if (!ws) return { error: 'Match sheet not found: ' + matchId };
  const data = ws.getDataRange().getValues();
  // Parse struktur sheet match
  return parseMatchSheet(data, matchId);
}

function parseMatchSheet(data, matchId) {
  // Info event: row 3-13 col C
  const info = {};
  const INFO_LABELS = ['Match ID','Tanggal','Venue','Jenis Game','Tipe Event',
                       'Tipe Lapangan','HTM Player','HTM GK','Kuota Kiper','Kuota Field','Status Match'];
  INFO_LABELS.forEach((lbl, i) => { info[lbl] = data[2+i] ? data[2+i][2] : null; });

  // Estimasi margin: row 16-21 di TEMPLATE_Match asli (row 15 = judul section
  // "ESTIMASI MARGIN", row 16-21 = 6 field-nya) — sebelumnya kode ini baca
  // row 17-22 (kegeser 1 baris), dibenerin biar cocok sama sheet aslinya.
  const margin = {
    salesPasti    : data[15] ? data[15][2] : 0,
    salesEstimasi : data[16] ? data[16][2] : 0,
    totalCost     : data[17] ? data[17][2] : 0,
    marginProyeksi: data[18] ? data[18][2] : 0,
    splitMalik    : data[19] ? data[19][2] : 0,
    splitFilan    : data[20] ? data[20][2] : 0,
  };

  // Daftar player: cari header row "No | Nama Player | ..."
  let playerStartRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (data[r][0] === 'No' && data[r][1] === 'Nama Player') {
      playerStartRow = r + 1; break;
    }
  }
  const players = [];
  if (playerStartRow > 0) {
    const hdrs = data[playerStartRow-1];
    for (let r = playerStartRow; r < data.length; r++) {
      if (!data[r][1]) break;
      const obj = {};
      hdrs.forEach((h,i) => { obj[h] = data[r][i]; });
      obj._row = r + 1;
      players.push(obj);
    }
  }

  return { matchId, info, margin, players };
}

// ── GET: Stok Jersey ────────────────────────────────────────
function getStokJersey() {
  return sheetToObjects(SHEET.JERSEY, 2);
}

// ── GET: SHTM ───────────────────────────────────────────────
function getShtm() {
  return sheetToObjects(SHEET.SHTM, 2)
    .filter(s => s['Nama Player']);
}

// ── GET: Deposit ────────────────────────────────────────────
function getDeposit() {
  return sheetToObjects(SHEET.DEPOSIT, 2);
}

// ── GET: Config ─────────────────────────────────────────────
function getConfig() {
  // Ambil dari sheet Rekap baris BASELINE untuk kumul
  const rekap = getMatches();
  const last  = rekap[rekap.length - 1] || {};
  return {
    htmPlayer  : 95000,
    htmGk      : 35000,
    htmShtm    : 50000,
    splitMalik : 60,
    kumulMalik : last['Kumul Malik\n(Rp)'] || last['Kumul Malik'] || 8147730,
    kumulFilan : last['Kumul Filan\n(Rp)'] || last['Kumul Filan'] || 5372762,
  };
}

// ── POST: Create Match ──────────────────────────────────────
function createMatch(p) {
  const ss      = SS();
  const tmpl    = ss.getSheetByName(SHEET.TEMPLATE || 'TEMPLATE_Match');
  if (!tmpl) return { error: 'TEMPLATE_Match sheet not found' };

  const matchId = generateId('Match');
  const newWs   = tmpl.copyTo(ss);
  newWs.setName(matchId);

  // Isi info event (col C, row 3-13)
  const vals = [
    matchId, p.tanggal, p.venue, p.jenisGame, p.tipeEvent,
    p.tipeLapangan, p.htmPlayer || 95000, p.htmGk || 35000,
    p.kuotaKiper || 4, p.kuotaField || 24, 'Aktif'
  ];
  vals.forEach((v, i) => { newWs.getRange(3+i, 3).setValue(v); });

  return { ok: true, matchId };
}

// ── POST: Add Player to Match ───────────────────────────────
function addPlayerToMatch(p) {
  const ws   = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found: ' + p.matchId };

  // Cari baris pertama kosong setelah header player
  const data = ws.getDataRange().getValues();
  let insertRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (data[r][0] === 'No' && data[r][1] === 'Nama Player') {
      // Cari baris kosong setelah header
      for (let rr = r+1; rr < data.length; rr++) {
        if (!data[rr][1]) { insertRow = rr+1; break; }
      }
      if (insertRow < 0) insertRow = data.length + 1;
      break;
    }
  }
  if (insertRow < 0) return { error: 'Player table not found in sheet' };

  const rowData = [
    p.no || '', p.nama, p.statusBayar || 'Sementara',
    p.htm, p.labelHarga || 'Normal',
    p.depositDipakai || 0, p.metodeBayar || '— (belum)',
    p.jumlahBayar || 0, p.tglBayar || '',
    p.posisi || 'Field', p.jerseyWarna || '', p.ukuran || '',
    p.tim || '', p.formasi || '', p.shtmFlag || '—', p.catatan || ''
  ];
  ws.getRange(insertRow, 1, 1, rowData.length).setValues([rowData]);

  // Kurangi stok jersey jika warna dipilih
  if (p.jerseyWarna && p.jerseyWarna !== 'Pribadi') {
    updateJerseyStok(p.jerseyWarna, p.ukuran, 1);
  }
  // Pakai deposit jika metode = Deposit
  if (p.metodeBayar === 'Deposit' && p.depositDipakai > 0) {
    useDeposit({ namaPemain: p.nama, jumlah: p.depositDipakai, matchId: p.matchId });
  }
  return { ok: true, row: insertRow };
}

// ── POST: Update Player ─────────────────────────────────────
function updatePlayerInMatch(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  const rowData = [
    p.no, p.nama, p.statusBayar, p.htm, p.labelHarga,
    p.depositDipakai || 0, p.metodeBayar, p.jumlahBayar || 0, p.tglBayar || '',
    p.posisi, p.jerseyWarna, p.ukuran, p.tim, p.formasi, p.shtmFlag, p.catatan || ''
  ];
  ws.getRange(p.rowIdx, 1, 1, rowData.length).setValues([rowData]);
  return { ok: true };
}

// ── POST: Delete Player ─────────────────────────────────────
function deletePlayerFromMatch(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  // Kembalikan stok jersey
  const data = ws.getRange(p.rowIdx, 1, 1, 16).getValues()[0];
  const warna = data[10]; const uk = data[11];
  if (warna && warna !== 'Pribadi') updateJerseyStok(warna, uk, -1);
  ws.deleteRow(p.rowIdx);
  return { ok: true };
}

// ── POST: Biaya ─────────────────────────────────────────────
function addBiaya(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  // Cari section biaya di sheet (label "Item Biaya" dst.)
  // Append di baris kosong pertama di section biaya
  const data = ws.getDataRange().getValues();
  let insertRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (String(data[r][0]).toLowerCase().includes('item biaya') ||
        String(data[r][0]).toLowerCase().includes('nama item')) {
      for (let rr = r+1; rr < data.length; rr++) {
        if (!data[rr][0]) { insertRow = rr+1; break; }
      }
      break;
    }
  }
  if (insertRow < 0) insertRow = ws.getLastRow() + 1;
  const rowData = [p.nama, p.kategori, p.tipe, p.rencana || 0, p.realisasi || 0,
                   p.skema || '', p.jmlCicilan || '', p.catatan || ''];
  ws.getRange(insertRow, 1, 1, rowData.length).setValues([rowData]);
  return { ok: true, row: insertRow };
}

function updateBiaya(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  const rowData = [p.nama, p.kategori, p.tipe, p.rencana||0, p.realisasi||0,
                   p.skema||'', p.jmlCicilan||'', p.catatan||''];
  ws.getRange(p.rowIdx, 1, 1, rowData.length).setValues([rowData]);
  return { ok: true };
}

function deleteBiaya(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  ws.deleteRow(p.rowIdx);
  return { ok: true };
}

// ── POST: Goal ──────────────────────────────────────────────
function addGol(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  // Simpan gol di section goal scorer
  const data = ws.getDataRange().getValues();
  let insertRow = ws.getLastRow() + 1;
  for (let r = 0; r < data.length; r++) {
    if (String(data[r][0]).toLowerCase().includes('pencetak') ||
        String(data[r][0]).toLowerCase().includes('goal scorer')) {
      for (let rr = r+1; rr < data.length; rr++) {
        if (!data[rr][0]) { insertRow = rr+1; break; }
      }
      break;
    }
  }
  ws.getRange(insertRow, 1, 1, 3).setValues([[p.nama, p.gol, p.keterangan || '']]);
  return { ok: true };
}

function deleteGol(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };
  ws.deleteRow(p.rowIdx);
  return { ok: true };
}

// ── POST: Flag SHTM ─────────────────────────────────────────
function flagShtm(p) {
  const ws  = getSheet(SHEET.SHTM);
  const lastRow = ws.getLastRow() + 1;
  const no  = lastRow - 2; // nomor urut
  ws.getRange(lastRow, 1, 1, 7).setValues([[
    no, p.nama, p.tglMatch, null, 'Belum dipakai',
    p.matchLabel || p.tglMatch, p.catatan || 'Flag dari Slot Player'
  ]]);
  return { ok: true };
}

// ── POST: Use SHTM (saat player hadir dgn SHTM di match berikutnya) ──
function useShtm(p) {
  const ws   = getSheet(SHEET.SHTM);
  const data = ws.getDataRange().getValues();
  for (let r = 2; r < data.length; r++) {
    if (data[r][1] === p.nama && data[r][4] === 'Belum dipakai') {
      ws.getRange(r+1, 4).setValue(p.tglDipakai);
      ws.getRange(r+1, 5).setValue('Sudah dipakai');
      ws.getRange(r+1, 6).setValue(p.matchLabel || p.tglDipakai);
      return { ok: true, row: r+1 };
    }
  }
  return { error: 'SHTM aktif tidak ditemukan untuk: ' + p.nama };
}

// ── POST: Deposit ───────────────────────────────────────────
function addDeposit(p) {
  const ws      = getSheet(SHEET.DEPOSIT);
  const lastRow = ws.getLastRow() + 1;
  ws.getRange(lastRow, 1, 1, 8).setValues([[
    p.tgl || today(), p.nama, p.matchId || '—',
    p.sumber || 'Cancel → Deposit', p.keterangan || '',
    p.jumlah, 0, p.jumlah
  ]]);
  return { ok: true };
}

function useDeposit(p) {
  const ws   = getSheet(SHEET.DEPOSIT);
  const data = ws.getDataRange().getValues();
  // Hitung saldo aktif player
  let saldo = 0;
  for (let r = 2; r < data.length; r++) {
    if (data[r][1] === p.namaPemain) {
      saldo += (data[r][5] || 0) - (data[r][6] || 0);
    }
  }
  if (saldo < p.jumlah) return { error: 'Saldo deposit tidak cukup' };
  const lastRow = ws.getLastRow() + 1;
  ws.getRange(lastRow, 1, 1, 8).setValues([[
    today(), p.namaPemain, p.matchId,
    'Digunakan', 'Bayar HTM', 0, p.jumlah, saldo - p.jumlah
  ]]);
  return { ok: true, saldoSisa: saldo - p.jumlah };
}

// ── POST: Stok Jersey ───────────────────────────────────────
function updateJerseyStok(warna, ukuran, delta) {
  // delta: +1 = tambah dipakai, -1 = kembalikan
  const ws   = getSheet(SHEET.JERSEY);
  const data = ws.getDataRange().getValues();
  for (let r = 2; r < data.length; r++) {
    if (data[r][0] === warna && data[r][2] === ukuran) {
      const dipakai = (data[r][4] || 0) + delta;
      ws.getRange(r+1, 5).setValue(Math.max(0, dipakai));
      ws.getRange(r+1, 6).setValue((data[r][3] || 0) - Math.max(0, dipakai));
      const sisa = (data[r][3] || 0) - Math.max(0, dipakai);
      ws.getRange(r+1, 7).setValue(sisa === 0 ? 'HABIS' : sisa <= 1 ? 'LOW' : 'OK');
      return { ok: true };
    }
  }
  return { error: 'Jersey tidak ditemukan: ' + warna + ' ' + ukuran };
}

function updateStokJersey(p) {
  const ws   = getSheet(SHEET.JERSEY);
  const data = ws.getDataRange().getValues();
  // Cari baris yang ada
  for (let r = 2; r < data.length; r++) {
    if (data[r][0] === p.warna && data[r][1] === p.tipe && data[r][2] === p.ukuran) {
      const newStok = (data[r][3] || 0) + (p.tambah || 0);
      ws.getRange(r+1, 4).setValue(newStok);
      const sisa = newStok - (data[r][4] || 0);
      ws.getRange(r+1, 6).setValue(sisa);
      ws.getRange(r+1, 7).setValue(sisa === 0 ? 'HABIS' : sisa <= 1 ? 'LOW' : 'OK');
      if (p.keterangan) ws.getRange(r+1, 8).setValue(p.keterangan);
      return { ok: true, newStok };
    }
  }
  // Warna/ukuran baru — tambah baris
  const lastRow = ws.getLastRow() + 1;
  ws.getRange(lastRow, 1, 1, 8).setValues([[
    p.warna, p.tipe || 'Field', p.ukuran,
    p.tambah || 1, 0, p.tambah || 1,
    p.tambah > 1 ? 'OK' : 'LOW',
    p.keterangan || ''
  ]]);
  return { ok: true, newRow: lastRow };
}

// ── POST: Close Match ───────────────────────────────────────
function closeMatch(p) {
  const ws = SS().getSheetByName(p.matchId);
  if (!ws) return { error: 'Sheet not found' };

  // 1. Set status match = Ditutup
  const data = ws.getDataRange().getValues();
  for (let r = 0; r < data.length; r++) {
    if (data[r][0] === 'Status Match') {
      ws.getRange(r+1, 3).setValue('Ditutup');
      break;
    }
  }

  // 2. Hitung sales, cost, margin dari data player & biaya
  const sales  = p.totalSales  || 0;
  const cost   = p.totalCost   || 0;
  const margin = sales - cost;
  const sm     = Math.round(margin * 0.6);
  const sf     = Math.round(margin * 0.4);
  const inv    = p.totalInv    || 0;
  const nm     = sm - Math.round(inv * 0.6);
  const nf     = sf - Math.round(inv * 0.4);

  // 3. Tambah baris ke Rekap_Keuangan
  const rekap  = getSheet(SHEET.REKAP);
  // Cari baris BASELINE atau last row
  const rdata  = rekap.getDataRange().getValues();
  let insertRow = rekap.getLastRow() + 1;
  for (let r = rdata.length-1; r >= 0; r--) {
    if (rdata[r][0] === 'BASELINE') { insertRow = r+1; break; }
  }

  // Kumul dari baris sebelumnya
  const prevRow  = rdata[insertRow-2] || [];
  const prevKumM = (prevRow[14] && typeof prevRow[14] === 'number') ? prevRow[14] : 8147730;
  const prevKumF = (prevRow[15] && typeof prevRow[15] === 'number') ? prevRow[15] : 5372762;
  const kumM     = prevKumM + nm;
  const kumF     = prevKumF + nf;

  const matchNo = (insertRow - 3); // nomor match baru
  rekap.getRange(insertRow, 1, 1, 22).setValues([[
    matchNo, p.tglMatch, p.jenisGame, p.tipeEvent,
    p.venue, p.tipeLapangan,
    sales, cost, margin, sm, sf,
    inv, nm, nf, kumM, kumF,
    p.bca||0, p.bsi||0, p.mandiri||0, p.bri||0, p.cash||0,
    ''
  ]]);
  // Kolom 23 "Tahun" — dipakai grafik "Profit per match" biar bisa
  // dikelompokin per kuartal/tahun. Match lama (sebelum kolom ini ada)
  // diisi manual sama Sam; match baru yang ditutup lewat sistem ini
  // otomatis kepakai tahun berjalan saat ditutup.
  rekap.getRange(insertRow, 23).setValue(new Date().getFullYear());

  // 4. Kembalikan semua stok jersey dari match ini
  // (data player sudah dikurangi saat addPlayer, dikembalikan jika deletePlayer)
  // Saat close, jersey dianggap dikembalikan (stok dipakai → 0 lagi)
  if (p.jerseyUsed && Array.isArray(p.jerseyUsed)) {
    p.jerseyUsed.forEach(j => updateJerseyStok(j.warna, j.uk, -1));
  }

  return {
    ok: true, matchNo, sales, cost, margin,
    splitMalik: sm, splitFilan: sf,
    kumulMalik: kumM, kumulFilan: kumF
  };
}

// ── POST: Update Config ─────────────────────────────────────
function updateConfig(p) {
  // Simpan config ke PropertiesService (persisten, tidak di Sheets)
  const props = PropertiesService.getScriptProperties();
  if (p.htmPlayer)  props.setProperty('HTM_PLAYER',  String(p.htmPlayer));
  if (p.htmGk)      props.setProperty('HTM_GK',      String(p.htmGk));
  if (p.htmShtm)    props.setProperty('HTM_SHTM',    String(p.htmShtm));
  if (p.splitMalik) props.setProperty('SPLIT_MALIK', String(p.splitMalik));
  return { ok: true };
}

// ============================================================
// ── SETUP: jalankan SEKALI aja dari editor Apps Script ──────
// Cara pakai: buka Extensions > Apps Script di Sheet "Minisoccer_DB",
// pilih function "setupSheets" di dropdown atas, klik Run (▶),
// izinkan akses saat diminta. Ini otomatis bikin semua tab + header
// yang dibutuhin backend ini (DB_Player, Stok_Jersey, SHTM_Log,
// Rekap_Keuangan, Deposit_Log, Inventaris, TEMPLATE_Match).
// Aman dijalankan berkali-kali — kalau tab udah ada, dilewatin
// (isinya TIDAK ditimpa), cuma tab yang belum ada yang dibikin.
// ============================================================
function setupSheets() {
  const ss = SS();
  const made = [];
  const skipped = [];

  function ensureSheet(name, rows, cols) {
    let ws = ss.getSheetByName(name);
    if (ws) { skipped.push(name); return ws; }
    ws = ss.insertSheet(name);
    if (rows || cols) ws.getRange(1,1, rows||1, cols||1); // no-op touch, avoids edge cases on some accounts
    made.push(name);
    return ws;
  }

  function writeHeader(ws, title, headers) {
    ws.getRange(1,1).setValue(title);
    ws.getRange(2,1,1,headers.length).setValues([headers]);
    ws.getRange(2,1,1,headers.length).setFontWeight('bold');
    ws.setFrozenRows(2);
  }

  // ── DB_Player ──
  let ws = ensureSheet(SHEET.PLAYER);
  if (ws.getLastRow() < 2) writeHeader(ws, 'DB_Player — master data player', [
    'No','Nama Player','Kategori','Total Hadir','Total Gol','Total SHTM',
    'Ukuran Jersey','Warna Favorit','Deposit Saldo (Rp)','Metode Favorit',
    'Bank','Kontak','Jersey Pribadi','Alias','Status','Catatan','Tgl Daftar'
  ]);

  // ── Stok_Jersey ── (diisi angka stok yang sekarang dipakai di index.html
  // supaya begitu tersambung, angkanya nyambung — bukan mulai dari 0)
  ws = ensureSheet(SHEET.JERSEY);
  if (ws.getLastRow() < 2) {
    writeHeader(ws, 'Stok_Jersey — stok jersey Field & GK', [
      'Warna','Tipe','Ukuran','Stok Total','Dipakai','Sisa','Status','Keterangan'
    ]);
    const stokField = {
      'Merah': {M:2,L:2,XL:0,'2XL':1}, 'Kuning': {M:1,L:3,XL:1,'2XL':1},
      'Biru': {M:2,L:3,XL:3,'2XL':1}, 'Hijau': {M:2,L:1,XL:4,'2XL':1}
    };
    const stokGk = { 'GK': {M:2,L:2,XL:2} };
    const rows = [];
    function pushRows(map, tipe) {
      Object.keys(map).forEach(warna => {
        Object.keys(map[warna]).forEach(uk => {
          const stok = map[warna][uk];
          const status = stok===0 ? 'HABIS' : stok<=1 ? 'LOW' : 'OK';
          rows.push([warna, tipe, uk, stok, 0, stok, status, '']);
        });
      });
    }
    pushRows(stokField, 'Field');
    pushRows(stokGk, 'GK');
    ws.getRange(3,1,rows.length,8).setValues(rows);
  }

  // ── SHTM_Log ──
  ws = ensureSheet(SHEET.SHTM);
  if (ws.getLastRow() < 2) writeHeader(ws, 'SHTM_Log — riwayat Sistem Half-price Top Match', [
    'No','Nama Player','Tgl Mendapat SHTM','Tgl SHTM Dipakai','Status','Match','Catatan'
  ]);

  // ── Rekap_Keuangan ──
  ws = ensureSheet(SHEET.REKAP);
  if (ws.getLastRow() < 2) {
    writeHeader(ws, 'Rekap_Keuangan — 1 baris per match yang ditutup', [
      'No','Match / Tanggal','Jenis Game','Tipe Event','Venue','Tipe Lapangan',
      'Sales (Rp)','Cost (Rp)','Margin (Rp)','Split Malik (Rp)','Split Filan (Rp)',
      'Inventaris (Rp)','Net Malik (Rp)','Net Filan (Rp)','Kumul Malik (Rp)','Kumul Filan (Rp)',
      'BCA (Rp)','BSI (Rp)','Mandiri (Rp)','BRI (Rp)','Cash (Rp)','Lap.'
    ]);
    // Baris BASELINE — titik awal hitungan kumulatif sebelum match pertama
    // ditutup lewat sistem ini. Ganti angka O/P di baris ini (Kumul Malik /
    // Kumul Filan) ke saldo riil kamu sebelum mulai pakai backend ini.
    const baseline = new Array(22).fill('');
    baseline[0] = 'BASELINE';
    baseline[14] = 8147730; // Kumul Malik (Rp) — sesuaikan ke saldo riil
    baseline[15] = 5372762; // Kumul Filan (Rp) — sesuaikan ke saldo riil
    ws.getRange(3,1,1,22).setValues([baseline]);
  } else {
    // Tab udah ada isinya (kasus real kamu) — kolom "Match / Tanggal" cuma
    // simpan tanggal singkat tanpa tahun ("21 Des", "12 April"), jadi grafik
    // yg butuh dikelompokin per kuartal/tahun (Profit per match) nggak bisa
    // dihitung akurat. Tambahin kolom "Tahun" di ujung (kolom 23) — ADDITIF,
    // nggak nimpa apa pun yang udah ada. Header row = row 2, sama kayak
    // header lainnya di sheet ini.
    const rHdr = ws.getRange(2,1,1,ws.getLastColumn()).getValues()[0];
    const hasTahun = rHdr.some(h => String(h).trim() === 'Tahun');
    if (!hasTahun) {
      ws.getRange(2, 23).setValue('Tahun');
    }
  }

  // ── Deposit_Log ──
  ws = ensureSheet(SHEET.DEPOSIT);
  if (ws.getLastRow() < 2) writeHeader(ws, 'Deposit_Log — riwayat deposit player', [
    'Tanggal','Nama Player','Match ID','Sumber','Keterangan',
    'Jumlah Masuk (Rp)','Jumlah Keluar (Rp)','Saldo (Rp)'
  ]);

  // ── Inventaris ──
  ws = ensureSheet(SHEET.INVENTARIS);
  if (ws.getLastRow() < 2) writeHeader(ws, 'Inventaris — belanja alat, potong kumulatif Malik/Filan', [
    'Tanggal','Item','Jumlah (Rp)','Keterangan','Split Malik (Rp)','Split Filan (Rp)'
  ]);

  // ── TEMPLATE_Match — sheet ini DI-COPY tiap kali "Buat match baru" ──
  // Kalau tab ini BELUM ada sama sekali: bikin dari nol, layout persis
  // sama kayak TEMPLATE_Match yang udah ada di sheet data riil kamu (info
  // row 3-13, judul "ESTIMASI MARGIN" row 15, 6 field margin row 16-21,
  // "DAFTAR PLAYER" row 23, header tabel player row 24) — supaya
  // parseMatchSheet() di kode ini baca row yang bener baik di sheet baru
  // maupun yang lama.
  ws = ss.getSheetByName('TEMPLATE_Match');
  if (!ws) {
    ws = ss.insertSheet('TEMPLATE_Match');
    made.push('TEMPLATE_Match');

    ws.getRange(1,1).setValue('TEMPLATE MATCH — jangan diisi manual, disalin otomatis tiap "Buat match baru"');
    ws.getRange(1,1).setFontWeight('bold');
    ws.getRange(2,1).setValue('INFO EVENT');
    ws.getRange(2,1).setFontWeight('bold');

    const infoLabels = ['Match ID','Tanggal','Venue','Jenis Game','Tipe Event',
      'Tipe Lapangan','HTM Player','HTM GK','Kuota Kiper','Kuota Field','Status Match'];
    infoLabels.forEach((lbl,i) => ws.getRange(3+i,1).setValue(lbl));

    ws.getRange(15,1).setValue('ESTIMASI MARGIN');
    ws.getRange(15,1).setFontWeight('bold');
    const marginLabels = ['Sales Pasti (Rp)','Sales Estimasi (Rp)','Total Cost (Rp)',
      'Margin Proyeksi (Rp)','Split Malik (Rp)','Split Filan (Rp)'];
    marginLabels.forEach((lbl,i) => ws.getRange(16+i,1).setValue(lbl));

    ws.getRange(23,1).setValue('DAFTAR PLAYER');
    ws.getRange(23,1).setFontWeight('bold');
    ws.getRange(24,1,1,16).setValues([[
      'No','Nama Player','Status Bayar','HTM','Label Harga','Deposit Dipakai',
      'Metode Bayar','Jumlah Bayar','Tgl Bayar','Posisi','Jersey Warna','Ukuran',
      'Tim','Formasi','Status SHTM','Catatan'
    ]]);
    ws.getRange(24,1,1,16).setFontWeight('bold');

    ws.getRange(60,1).setValue('BIAYA OPERASIONAL');
    ws.getRange(60,1).setFontWeight('bold');
    ws.getRange(61,1,1,8).setValues([[
      'Item Biaya','Kategori','Tipe','Rencana (Rp)','Realisasi (Rp)','Skema Split','Jml Cicilan','Catatan'
    ]]);
    ws.getRange(61,1,1,8).setFontWeight('bold');

    ws.getRange(85,1).setValue('PENCETAK GOL');
    ws.getRange(85,1).setFontWeight('bold');
    ws.getRange(86,1,1,3).setValues([['Nama','Gol','Keterangan']]);
    ws.getRange(86,1,1,3).setFontWeight('bold');
  } else {
    // Tab ini udah ada (kasus sheet data riil kamu) — JANGAN disentuh isi
    // row 1-24-nya sama sekali. Cuma cek: apa udah ada section "Biaya"
    // dan "Pencetak Gol"? Kalau belum (dan memang belum, di sheet riil
    // kamu TEMPLATE_Match cuma sampai row 24 = header Daftar Player),
    // ditambahin di BAWAH baris terakhir yang udah kepake — addBiaya()/
    // addGol() di kode ini butuh baris berlabel itu buat tau mau nulis
    // baris baru di mana.
    skipped.push('TEMPLATE_Match (isi lama tidak diubah)');
    const data = ws.getDataRange().getValues();
    const hasBiaya = data.some(row => {
      const c0 = String(row[0] || '').toLowerCase();
      return c0.includes('item biaya') || c0.includes('biaya operasional');
    });
    const hasGol = data.some(row => {
      const c0 = String(row[0] || '').toLowerCase();
      return c0.includes('pencetak') || c0.includes('goal scorer');
    });
    if (!hasBiaya || !hasGol) {
      // JANGAN taruh section baru cuma "3 baris di bawah baris terakhir
      // yang kepake" — baris terakhir yang kepake itu justru HEADER tabel
      // player (row 24), dan addPlayerToMatch() nyari baris kosong per
      // KOLOM B buat nyisipin player baru; kalau section baru ditaruh
      // kepepet deket situ, player ke-3/4 bisa numpuk nimpa label section.
      // Jadi kasih jarak cukup dari baris HEADER PLAYER (bukan dari baris
      // terakhir), muat sampai puluhan player per match.
      let playerHeaderRow = -1;
      for (let r = 0; r < data.length; r++) {
        if (data[r][0] === 'No' && data[r][1] === 'Nama Player') { playerHeaderRow = r + 1; break; }
      }
      let nextRow = playerHeaderRow > 0 ? playerHeaderRow + 40 : ws.getLastRow() + 3;
      if (!hasBiaya) {
        ws.getRange(nextRow,1).setValue('BIAYA OPERASIONAL');
        ws.getRange(nextRow,1).setFontWeight('bold');
        nextRow += 1;
        ws.getRange(nextRow,1,1,8).setValues([[
          'Item Biaya','Kategori','Tipe','Rencana (Rp)','Realisasi (Rp)','Skema Split','Jml Cicilan','Catatan'
        ]]);
        ws.getRange(nextRow,1,1,8).setFontWeight('bold');
        made.push('TEMPLATE_Match: section Biaya Operasional (ditambah di row ' + (nextRow-1) + ')');
        nextRow += 24; // headroom baris biaya
      }
      if (!hasGol) {
        nextRow += 2;
        ws.getRange(nextRow,1).setValue('PENCETAK GOL');
        ws.getRange(nextRow,1).setFontWeight('bold');
        nextRow += 1;
        ws.getRange(nextRow,1,1,3).setValues([['Nama','Gol','Keterangan']]);
        ws.getRange(nextRow,1,1,3).setFontWeight('bold');
        made.push('TEMPLATE_Match: section Pencetak Gol (ditambah di row ' + (nextRow-1) + ')');
      }
    }
  }

  // ── Rekap_Skorer ── (top skor all time, manual — field-nya sama persis
  // dgn tabel "Kelola Rekap Skorer" di Admin: Nama/Kategori/Tahun/Gol/
  // Uraian. BELUM ada di sheet data riil kamu juga, jadi ini tab baru,
  // aman dibikin kapan aja — nggak nyentuh 6 tab lain yang udah ada.
  // Diisi otomatis dari 7 data yang udah kamu masukin manual di Admin.
  ws = ensureSheet(SHEET.SCORER);
  if (ws.getLastRow() < 2) {
    writeHeader(ws, 'Rekap_Skorer — top skor all time (tampil di Dashboard & Goal scorer)', [
      'No','Nama','Kategori','Tahun','Total Gol','Uraian'
    ]);
    const seedScorer = [
      [1,'Annas','BMR','2026',19,''],
      [2,'Bambang','SMS','2026',13,''],
      [3,'Jayadi','Umum','2026',10,''],
      [4,'Ibnu','Umum','2026',7,''],
      [5,'Wasis','Umum','2026',7,''],
      [6,'Arsya','Umum','2026',4,''],
      [7,'Eri','SMS','2026',4,'']
    ];
    ws.getRange(3,1,seedScorer.length,6).setValues(seedScorer);
  }

  // Hapus "Sheet1" bawaan Google kalau masih ada dan kosong
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) {
    ss.deleteSheet(def);
  }

  const msg = 'Selesai. Tab dibuat: ' + (made.join(', ') || '(tidak ada, semua udah ada)') +
    ' | Tab dilewatin (udah ada sebelumnya): ' + (skipped.join(', ') || '(tidak ada)');
  Logger.log(msg);
  return msg;
}

// ── GET: Rekap Skorer (top skor all time, manual) ───────────
function getScorer() {
  return sheetToObjects(SHEET.SCORER, 2);
}

// ── POST: Rekap Skorer ───────────────────────────────────────
function addScorer(p) {
  const ws      = getSheet(SHEET.SCORER);
  const lastRow = ws.getLastRow() + 1;
  const no      = lastRow - 2;
  ws.getRange(lastRow, 1, 1, 6).setValues([[
    no, p.nama, p.kategori || '—', p.tahun || '', p.gol || 0, p.uraian || ''
  ]]);
  return { ok: true, row: lastRow };
}

function updateScorer(p) {
  const ws = getSheet(SHEET.SCORER);
  if (!ws) return { error: 'Sheet not found' };
  ws.getRange(p.rowIdx, 1, 1, 6).setValues([[
    p.no, p.nama, p.kategori || '—', p.tahun || '', p.gol || 0, p.uraian || ''
  ]]);
  return { ok: true };
}

function deleteScorer(p) {
  const ws = getSheet(SHEET.SCORER);
  if (!ws) return { error: 'Sheet not found' };
  ws.deleteRow(p.rowIdx);
  return { ok: true };
}

// ── POST: Add New Player to DB ──────────────────────────────
function addNewPlayer(p) {
  const ws      = getSheet(SHEET.PLAYER);
  const lastRow = ws.getLastRow() + 1;
  const no      = lastRow - 2;
  ws.getRange(lastRow, 1, 1, 17).setValues([[
    no, p.nama, p.kategori || '—', 0, 0, 0,
    p.ukuran || '—', '—', 0, '—', '—', '—', p.jerseyPribadi || '—',
    '—', 'Aktif', p.catatan || '', today()
  ]]);
  return { ok: true, row: lastRow };
}
