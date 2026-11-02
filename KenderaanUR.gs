// Tambah fail ini bersama Code.gs terkini; jangan gantikan Code.gs sedia ada.
const URV_HEADERS = ['JENIS KENDERAAN','NO PENDAFTARAN','JENIS MODEL KENDERAAN',
  'KUASA CC','KM PERGI','KM BALIK','SEBAB KENDERAAN SENDIRI','PEMANDU'];

function urvSchema_(write) {
  const sh = sheet_('BORANG_UR');
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
  if (h[0] !== 'REQ_ID' || h[15] !== 'STATUS SOKONGAN' || h[22] !== 'TARIKH LULUS') {
    throw Error('Struktur BORANG_UR tidak sepadan dengan e-MANDAT V6.1.');
  }
  const missing = URV_HEADERS.filter(x => !h.includes(x));
  if (write && missing.length) {
    const backup = sh.copyTo(ss_()).setName('BU_UR_KEND_' +
      Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyMMdd_HHmmss') + '_' + uuid_().slice(0,4));
    const first = h.length + 1;
    if (sh.getMaxColumns() < h.length + missing.length) {
      sh.insertColumnsAfter(sh.getMaxColumns(), h.length + missing.length - sh.getMaxColumns());
    }
    sh.getRange(1, first, 1, missing.length).setValues([missing]);
    h.push(...missing);
    Logger.log('Lajur kenderaan ditambah. Sandaran: ' + backup.getName());
  }
  return { sh: sh, headers: h };
}

function urvValidate_(raw) {
  const p = raw || {}, jenis = norm_(p.jenis);
  if (!['JABATAN','SENDIRI'].includes(jenis)) throw Error('Pilih Kenderaan Jabatan atau Kenderaan Sendiri.');
  const v = {jenis:jenis, nombor:norm_(p.nombor), model:'', cc:'', kmPergi:'', kmBalik:'', sebab:'', pemandu:''};
  if (v.nombor.length > 30) throw Error('Nombor pendaftaran terlalu panjang.');
  if (jenis === 'JABATAN') {
    v.pemandu = norm_(p.pemandu);
    if (v.pemandu.length > 200) throw Error('Nama pemandu terlalu panjang.');
  } else {
    v.model = norm_(p.model); v.sebab = norm_(p.sebab);
    if (!v.nombor || !v.model || !v.sebab) throw Error('Nombor pendaftaran, jenis/model dan sebab menggunakan kenderaan sendiri wajib diisi.');
    if (v.model.length > 150 || v.sebab.length > 2000) throw Error('Maklumat kenderaan terlalu panjang.');
    const cc = Number(p.cc), pergi = Number(p.kmPergi), balik = Number(p.kmBalik);
    if (p.cc === '' || p.cc == null || !Number.isFinite(cc) || cc <= 0 || cc > 20000 || !Number.isInteger(cc)) throw Error('Isikan kuasa enjin dalam CC yang sah.');
    if (p.kmPergi === '' || p.kmPergi == null || p.kmBalik === '' || p.kmBalik == null ||
      !Number.isFinite(pergi) || !Number.isFinite(balik) || pergi < 0 || balik < 0 ||
      pergi + balik <= 0 || pergi + balik > 20000) throw Error('Anggaran KM pergi/balik tidak sah.');
    v.cc = cc; v.kmPergi = Math.round(pergi*10)/10; v.kmBalik = Math.round(balik*10)/10;
  }
  ['nombor','model','sebab','pemandu'].forEach(k => v[k] = safe_(v[k]));
  return v;
}

function urvFromRow_(r, headers) {
  const at = x => { const i = headers.indexOf(x); return i < 0 ? '' : (r[i] == null ? '' : r[i]); };
  return {jenis:at(URV_HEADERS[0]), nombor:at(URV_HEADERS[1]), model:at(URV_HEADERS[2]),
    cc:at(URV_HEADERS[3]), kmPergi:at(URV_HEADERS[4]), kmBalik:at(URV_HEADERS[5]),
    sebab:at(URV_HEADERS[6]), pemandu:at(URV_HEADERS[7])};
}

function urvMap_() {
  const schema = urvSchema_(false);
  return new Map(rows_(schema.sh, schema.headers.length).map(r => [r[0],urvFromRow_(r,schema.headers)]));
}

function submitBorangURKenderaan(token, p) {
  return locked_(() => {
    const u = session_(token);
    validateForm_(u,p,'UR');
    const v = urvValidate_(p.kenderaan);
    const schema = urvSchema_(true), id = 'UR-' + uuid_();
    const now = Utilities.formatDate(new Date(),'Asia/Kuala_Lumpur','yyyy-MM-dd HH:mm:ss');
    const r = [id,now,u.id,u.name,u.unit,u.jawatan,u.gred,p.perihal,p.mula,p.akhir,
      p.lokasi,p.tempatLuar,p.rujukan,p.tarikhSurat,p.diminitkan,
      'MENUNGGU SOKONGAN','','','','MENUNGGU KELULUSAN','','',''];
    while (r.length < schema.headers.length) r.push('');
    const values = [v.jenis,v.nombor,v.model,v.cc,v.kmPergi,v.kmBalik,v.sebab,v.pemandu];
    URV_HEADERS.forEach((h,i) => r[schema.headers.indexOf(h)] = values[i]);
    schema.sh.appendRow(r);
    for (let day=p.mula;day<=p.akhir;day=addDays_(day,1)) {
      const dow = parseDay_(day).getUTCDay();
      if (dow>=1 && dow<=5) upsertRecord_(u.id,day,'UR (Sokongan)','Rujuk UR: '+id,u.id);
    }
    return {ok:true,reqId:id};
  });
}

function getMyFormsKenderaan(token) {
  const result = getMyForms(token), map = urvMap_();
  result.ur.forEach(p => p.kenderaan = map.get(p.reqId) || {});
  return result;
}

function adminGetPendingURKenderaan(token) {
  const actor = session_(token), result = adminGetPendingUR(token), map = urvMap_();
  result.sokong = result.sokong.filter(p => adminScope_(actor,{unit:p.unit}));
  result.lulus = result.lulus.filter(p => adminScope_(actor,{unit:p.unit}));
  result.sokong.forEach(p => p.kenderaan = map.get(p.reqId) || {});
  result.lulus.forEach(p => p.kenderaan = map.get(p.reqId) || {});
  return result;
}

function getBorangKenderaanUR(token, reqId) {
  const u = session_(token), schema = urvSchema_(false);
  const r = rows_(schema.sh,schema.headers.length).find(r => r[0]===norm_(reqId));
  if (!r) throw Error('Permohonan tidak ditemui.');
  if (u.id!==r[2] && !((checkAccess_(u,'SOKONG_UR') || checkAccess_(u,'LULUS_UR')) && adminScope_(u,{unit:r[4]}))) throw Error('Tiada akses kepada borang ini.');
  const k = urvFromRow_(r,schema.headers);
  if (k.jenis!=='SENDIRI') throw Error('Borang ini hanya untuk permohonan menggunakan kenderaan sendiri.');
  const p = {reqId:r[0],tMohon:r[1],name:r[3],unit:r[4],jawatan:r[5],gred:r[6],perihal:r[7],
    mula:r[8],akhir:r[9],lokasi:r[10],tempatLuar:r[11],rujukan:r[12],tarikhSurat:r[13],
    statusSokong:r[15],idSokong:r[16],ulasanSokong:r[17],tSokong:r[18],
    statusLulus:r[19],idLulus:r[20],ulasanLulus:r[21],tLulus:r[22],kenderaan:k};
  const applicant = userRow_(r[2]); p.mySign = applicant.row ? applicant.row[7] : '';
  const fill = (id, prefix, approved) => {
    if (!id || !approved) return;
    const staff = staff_(id)[0], account = userRow_(id);
    p[prefix+'Sign'] = account.row ? account.row[7] : '';
    p[prefix+'Stamp'] = account.row ? account.row[8] : '';
    p[prefix==='sokong'?'namaSokong':'namaLulus'] = staff ? staff[1] : id;
    p[prefix==='sokong'?'jwtSokong':'jwtLulus'] = staff ? staff[3]+' '+staff[4] : '';
  };
  fill(r[16],'sokong',r[15]==='DISOKONG');
  fill(r[20],'lulus',r[19]==='DILULUSKAN');
  return p;
}
