/* ====================== LAPORAN KUNJUNGAN TEKNISI ======================
   Single-page app: form -> generate .docx (client-side) -> upload to
   Google Drive folder via Google Identity Services + Drive API v3.
   No backend/server required. All logic runs in the visitor's browser.
========================================================================= */

/* ---------------------- CONFIG (persisted locally) --------------------- */
const CFG_KEYS = { clientId: 'lkt_client_id', folderId: 'lkt_folder_id' };
function getCfg(key, fallback){ try{ return localStorage.getItem(key) || fallback; }catch(e){ return fallback; } }
function setCfg(key, val){ try{ localStorage.setItem(key, val); }catch(e){} }

function toggleConfig(){
  document.getElementById('configPanel').classList.toggle('show');
}
function saveConfig(){
  const cid = document.getElementById('cfgClientId').value.trim();
  const fid = document.getElementById('cfgFolderId').value.trim();
  setCfg(CFG_KEYS.clientId, cid);
  setCfg(CFG_KEYS.folderId, fid);
  document.getElementById('cfgSavedMsg').textContent = 'Tersimpan ✓';
  setTimeout(()=>{ document.getElementById('cfgSavedMsg').textContent=''; }, 2000);
  initGoogleAuth();
}

/* ------------------------------ STATE ----------------------------------- */
const state = {
  fields: {},            // simple text/select field values, keyed by id
  compare: {},           // status compare items -> {before:'baik'|'buruk', after:'baik'|'buruk', extra:{...}}
  customRows: {           // dynamic user-added rows
    outdoor: [], indoor: [], device: []
  },
  photos: {}              // key -> {file, dataUrl, name}
};

/* ------------------------- FIELD DEFINITIONS ----------------------------- */

const OUTDOOR_ITEMS = [
  { key:'dish', label:'Dish / Antenna' },
  { key:'pedestal', label:'Pedestal Antena' },
  { key:'los', label:'Antenna Line of Sight (LOS)' },
  { key:'baut', label:'Kekencangan Baut Antena' },
  { key:'feedhorn', label:'Feed Horn' },
  { key:'lnalnb', label:'LNA / LNB' },
  { key:'rfbucodu', label:'RF / BUC / ODU' },
  { key:'kabelrf', label:'Kabel RF (RG6/RG11/RG8)', meterLabel:'Panjang kabel (meter)' },
  { key:'konektor', label:'Konektor' },
  { key:'groundingodu', label:'Kabel Grounding ODU' },
];

const INDOOR_ITEMS = [
  { key:'modem', label:'Modem' },
  { key:'router', label:'Router' },
  { key:'poe', label:'POE' },
  { key:'ap1', label:'Access Point #1' },
  { key:'ap2', label:'Access Point #2' },
  { key:'utp_modem_router', label:'Kabel UTP Modem - Router' },
  { key:'utp_router_ap1', label:'Kabel UTP Router - AP1' },
  { key:'utp_router_ap2', label:'Kabel UTP Router - AP2' },
  { key:'rj45', label:'Konektor RJ45' },
  { key:'rg', label:'Konektor RG11 / RG8' },
  { key:'grounding_indoor', label:'Kabel Grounding Indoor' },
  { key:'label_perangkat', label:'Label Perangkat' },
  { key:'stabilizer_indoor', label:'Stabilizer' },
];

const PARAMETER_ITEMS = [
  { key:'ping', label:'Ping (Min / Max / Avg / Packet Loss)', kind:'ping' },
  { key:'sqf', label:'SQF', kind:'value' },
  { key:'esno', label:'Es/No', kind:'value' },
  { key:'cn_cpi', label:'C/N - CPI', kind:'value' },
  { key:'power_atten', label:'Power Attenuation', kind:'value' },
  { key:'rx_modcod', label:'Rx Modcod', kind:'value' },
  { key:'tx_modcod', label:'Tx Modcod', kind:'value' },
];

const POWER_ITEMS = [
  { key:'stabilizer_power', label:'Stabilizer' },
];

const POWER_PARAMS = [
  { key:'source_power', label:'Sumber Power (PLN/Genset/Solar/dll)', kind:'text' },
  { key:'waktu_operasional', label:'Waktu Operasional (hari & jam)', kind:'text' },
  { key:'teg_input_source', label:'Tegangan Input Source Power (Volt)', kind:'value' },
  { key:'teg_output_ups', label:'Tegangan Output UPS (Volt)', kind:'value' },
  { key:'teg_output_stabilizer', label:'Tegangan Output Stabilizer (Volt)', kind:'value' },
  { key:'teg_grounding', label:'Tegangan Grounding Power Source (Volt)', kind:'value' },
];

const ALT_NETWORK_ITEMS = [
  { key:'selular4g', label:'Selular 4G' },
  { key:'selular3g', label:'Selular 3G' },
  { key:'selular2g', label:'Selular 2G' },
  { key:'rl', label:'Radio Link (RL)' },
  { key:'fo_sekitar', label:'Fiber Optik di Sekitar Lokasi' },
  { key:'fo_terpasang', label:'Fiber Optik Terpasang' },
  { key:'jaringan_lain', label:'Jaringan Lainnya' },
];

const DEVICE_ITEMS = [
  { key:'modem_dev', label:'Modem' },
  { key:'transiver_dev', label:'Transiver' },
  { key:'router_dev', label:'Router' },
  { key:'ap1_dev', label:'Access Point #1' },
  { key:'ap2_dev', label:'Access Point #2' },
  { key:'stabilizer_dev', label:'Stabilizer' },
];

const PHOTO_ITEMS = [
  ['sumber_listrik','Sumber Listrik (PLN / Genset / PLTS / Solar Cell / UPS)'],
  ['teknisi_lokasi','Foto Teknisi di Depan Informasi Nama Lokasi'],
  ['antena_odu_sebelum','Foto Antenna (Sebelum) — Tampak Seluruh ODU'],
  ['antena_odu_sesudah','Foto Antenna (Sesudah) — Tampak Seluruh ODU'],
  ['antena_los_sebelum','Foto Antenna (Sebelum) — Posisi / Tampak LOS'],
  ['antena_los_sesudah','Foto Antenna (Sesudah) — Posisi / Tampak LOS'],
  ['transceiver_sebelum','Foto Transceiver (Sebelum)'],
  ['transceiver_sesudah','Foto Transceiver (Sesudah)'],
  ['grounding_sebelum','Foto Grounding (Sebelum)'],
  ['grounding_sesudah','Foto Grounding (Sesudah)'],
  ['modem_sebelum','Foto Modem (Sebelum)'],
  ['modem_sesudah','Foto Modem (Sesudah)'],
  ['router_sebelum','Foto Router (Sebelum)'],
  ['router_sesudah','Foto Router (Sesudah)'],
  ['wallmount_sebelum','Foto Wallmount Rack (Sebelum)'],
  ['wallmount_sesudah','Foto Wallmount Rack (Sesudah)'],
  ['ap1_sebelum','Foto Access Point #1 (Sebelum)'],
  ['ap1_sesudah','Foto Access Point #1 (Sesudah)'],
  ['ap2_sebelum','Foto Access Point #2 (Sebelum)'],
  ['ap2_sesudah','Foto Access Point #2 (Sesudah)'],
  ['stabilizer_sebelum','Foto Stabilizer (Sebelum)'],
  ['stabilizer_sesudah','Foto Stabilizer (Sesudah)'],
  ['lain_sebelum','Foto Lain-Lain (Sebelum)'],
  ['lain_sesudah','Foto Lain-Lain (Sesudah)'],
  ['speedtest_lokasi_sebelum','Capture Speedtest SSID Nama Lokasi (Sebelum)'],
  ['speedtest_lokasi_sesudah','Capture Speedtest SSID Nama Lokasi (Sesudah)'],
  ['speedtest_baktiaksi_sebelum','Capture Speedtest SSID BAKTI AKSI (Sebelum)'],
  ['speedtest_baktiaksi_sesudah','Capture Speedtest SSID BAKTI AKSI (Sesudah)'],
  ['wifi_analyzer_ap1','Foto Wifi Analyzer AP#1'],
  ['wifi_analyzer_ap2','Foto Wifi Analyzer AP#2'],
  ['wifi_analyzer_ap1_baktiaksi','Foto Wifi Analyzer AP#1 — BAKTI AKSI'],
  ['wifi_analyzer_ap2_baktiaksi','Foto Wifi Analyzer AP#2 — BAKTI AKSI'],
  ['berita_acara','Lampiran Foto Berita Acara'],
];

/* --------------------------- RENDER HELPERS ------------------------------ */

function el(tag, attrs, ...children){
  const e = document.createElement(tag);
  if(attrs) for(const k in attrs){
    if(k === 'html') e.innerHTML = attrs[k];
    else if(k.startsWith('on')) e.addEventListener(k.slice(2), attrs[k]);
    else e.setAttribute(k, attrs[k]);
  }
  children.flat().forEach(c=>{
    if(c==null) return;
    e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
  });
  return e;
}

function sectionCard(num, title, sub, contentEl){
  return el('div',{class:'card'},
    el('h2',{}, el('span',{class:'section-num'}, String(num)), title),
    sub ? el('p',{class:'sub'}, sub) : null,
    contentEl
  );
}

function simpleField(id, labelText, type, opts){
  opts = opts || {};
  state.fields[id] = state.fields[id] || opts.default || '';
  let inputEl;
  if(type === 'textarea'){
    inputEl = el('textarea',{id:'f_'+id, oninput:e=>state.fields[id]=e.target.value}, state.fields[id]);
  } else if(type === 'select'){
    inputEl = el('select',{id:'f_'+id, onchange:e=>state.fields[id]=e.target.value});
    (opts.options||[]).forEach(o=>{
      const optEl = el('option',{value:o}, o);
      if(o === state.fields[id]) optEl.setAttribute('selected','selected');
      inputEl.appendChild(optEl);
    });
  } else {
    inputEl = el('input',{id:'f_'+id, type: type||'text', value: state.fields[id], oninput:e=>state.fields[id]=e.target.value});
    if(opts.placeholder) inputEl.setAttribute('placeholder', opts.placeholder);
  }
  return el('div',{},
    el('label',{for:'f_'+id}, labelText, opts.optional ? el('span',{class:'opt'},' (opsional)') : null),
    inputEl,
    opts.hint ? el('p',{class:'hint'}, opts.hint) : null
  );
}

function compareItemEl(itemKey, label, opts){
  opts = opts || {};
  state.compare[itemKey] = state.compare[itemKey] || { before:'', after:'', extra:{} };
  const c = state.compare[itemKey];

  function toggleGroup(side){
    const grp = el('div',{class:'toggle-group'});
    ['baik','buruk'].forEach(v=>{
      const btn = el('button',{type:'button', class:'toggle-btn'+(c[side]===v?(' active-'+v):'')},
        v==='baik' ? '✓ Baik' : '✕ Buruk');
      btn.addEventListener('click', ()=>{
        c[side] = v;
        grp.querySelectorAll('.toggle-btn').forEach(b=>b.className='toggle-btn');
        grp.querySelectorAll('.toggle-btn').forEach((b,i)=>{ if((i===0&&v==='baik')||(i===1&&v==='buruk')) b.classList.add('active-'+v); });
      });
      grp.appendChild(btn);
    });
    return grp;
  }

  const extraEls = [];
  if(opts.meterLabel){
    c.extra.meter_before = c.extra.meter_before || '';
    c.extra.meter_after = c.extra.meter_after || '';
  }

  const beforeCol = el('div',{class:'ci-col'},
    el('span',{class:'ci-tag'},'Sebelum'),
    toggleGroup('before'),
    opts.meterLabel ? el('div',{class:'extra-field'},
      el('input',{type:'text', placeholder:opts.meterLabel, value:c.extra.meter_before,
        oninput:e=>c.extra.meter_before=e.target.value})) : null
  );
  const afterCol = el('div',{class:'ci-col'},
    el('span',{class:'ci-tag'},'Sesudah'),
    toggleGroup('after'),
    opts.meterLabel ? el('div',{class:'extra-field'},
      el('input',{type:'text', placeholder:opts.meterLabel, value:c.extra.meter_after,
        oninput:e=>c.extra.meter_after=e.target.value})) : null
  );

  return el('div',{class:'compare-item'},
    el('div',{class:'ci-label'}, label),
    el('div',{class:'ci-grid'}, beforeCol, afterCol)
  );
}

function customRowsBlock(groupName, addLabel, placeholder){
  const wrap = el('div',{class:'custom-row-list'});
  function redraw(){
    wrap.innerHTML = '';
    state.customRows[groupName].forEach((row, idx)=>{
      const key = groupName+'_custom_'+idx;
      if(!row.nameInputBound){
        // ensure compare state exists under a stable key
      }
      const nameInput = el('input',{type:'text', placeholder:placeholder, value:row.name,
        oninput:e=>row.name=e.target.value, style:'margin-bottom:8px;'});
      const compareBlock = compareItemEl('custom_'+groupName+'_'+row.id, '', {});
      // remove default label line since name is editable above
      compareBlock.querySelector('.ci-label').remove();
      const item = el('div',{class:'compare-item', style:'position:relative;'},
        el('button',{type:'button', class:'remove-x', onclick:()=>{ state.customRows[groupName].splice(idx,1); redraw(); }}, '✕'),
        nameInput,
        compareBlock.querySelector('.ci-grid')
      );
      wrap.appendChild(item);
    });
    wrap.appendChild(el('button',{type:'button', class:'add-link', onclick:()=>{
      state.customRows[groupName].push({ id: Date.now()+Math.random().toString(16).slice(2), name:'' });
      redraw();
    }}, '+ '+addLabel));
  }
  redraw();
  return wrap;
}

/* ------------------------------ PARAMETER ITEM (ping/value) -------------- */

function parameterItemEl(item){
  state.compare[item.key] = state.compare[item.key] || { before:{}, after:{} };
  const c = state.compare[item.key];

  function valueInputs(side){
    if(item.kind === 'ping'){
      c[side] = c[side] || {};
      ['min','max','avg','loss'].forEach(f=> c[side][f] = c[side][f] || '');
      const labels = {min:'Min (ms)', max:'Max (ms)', avg:'Avg (ms)', loss:'Packet Loss (%)'};
      return el('div',{},
        Object.keys(labels).map(f=>
          el('div',{class:'extra-field'},
            el('input',{type:'text', placeholder:labels[f], value:c[side][f],
              oninput:e=>c[side][f]=e.target.value}))
        )
      );
    } else if(item.kind === 'text'){
      c[side+'_val'] = c[side+'_val'] || '';
      return el('input',{type:'text', value:c[side+'_val'], placeholder:item.label,
        oninput:e=>c[side+'_val']=e.target.value});
    } else {
      c[side+'_val'] = c[side+'_val'] || '';
      return el('input',{type:'text', value:c[side+'_val'], placeholder:'Nilai',
        oninput:e=>c[side+'_val']=e.target.value});
    }
  }

  return el('div',{class:'compare-item'},
    el('div',{class:'ci-label'}, item.label),
    el('div',{class:'ci-grid'},
      el('div',{class:'ci-col'}, el('span',{class:'ci-tag'},'Sebelum'), valueInputs('before')),
      el('div',{class:'ci-col'}, el('span',{class:'ci-tag'},'Sesudah'), valueInputs('after'))
    )
  );
}

function altNetworkItemEl(item){
  state.compare[item.key] = state.compare[item.key] || { status:'', operator:'' };
  const c = state.compare[item.key];
  const sel = el('select',{onchange:e=>c.status=e.target.value},
    el('option',{value:''},'— Pilih —'),
    el('option',{value:'Ada'},'Ada'),
    el('option',{value:'Tidak Ada'},'Tidak Ada')
  );
  Array.from(sel.options).forEach(o=>{ if(o.value===c.status) o.selected = true; });
  const opInput = el('input',{type:'text', placeholder:'Nama operator / keterangan', value:c.operator,
    oninput:e=>c.operator=e.target.value});
  return el('div',{class:'compare-item'},
    el('div',{class:'ci-label'}, item.label),
    el('div',{class:'row2'}, sel, opInput)
  );
}

/* --------------------------------- PHOTOS --------------------------------- */

function compressImage(file, maxDim, quality){
  return new Promise((resolve, reject)=>{
    const img = new Image();
    const reader = new FileReader();
    reader.onload = ()=>{ img.src = reader.result; };
    reader.onerror = reject;
    img.onload = ()=>{
      let { width, height } = img;
      if(width > height && width > maxDim){ height = Math.round(height*maxDim/width); width = maxDim; }
      else if(height > maxDim){ width = Math.round(width*maxDim/height); height = maxDim; }
      const canvas = document.createElement('canvas');
      canvas.width = width; canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, width, height);
      canvas.toBlob(blob=>{
        const fr = new FileReader();
        fr.onload = ()=> resolve({ blob, dataUrl: fr.result, width, height });
        fr.onerror = reject;
        fr.readAsDataURL(blob);
      }, 'image/jpeg', quality || 0.72);
    };
    img.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function photoSlotEl(key, label){
  const inputId = 'photo_'+key;
  const preview = el('div',{class:'photo-preview', id:'prev_'+key}, 'Belum ada foto');
  const slot = el('div',{class:'photo-slot', id:'slot_'+key},
    el('div',{class:'ps-label'}, label),
    preview,
    el('div',{class:'photo-actions'},
      el('label',{class:'upl', for:inputId}, '📷 Ambil / Upload'),
      el('input',{type:'file', id:inputId, accept:'image/*', capture:'environment',
        onchange: async (e)=>{
          const file = e.target.files[0];
          if(!file) return;
          preview.textContent = 'Memproses...';
          try{
            const { blob, dataUrl, width, height } = await compressImage(file, 1400, 0.72);
            state.photos[key] = { blob, dataUrl, width, height, name: key+'.jpg' };
            preview.innerHTML = '';
            preview.appendChild(el('img',{src:dataUrl}));
            slot.classList.add('filled');
          }catch(err){
            preview.textContent = 'Gagal memuat foto';
          }
        }
      }),
      el('button',{type:'button', class:'clear-btn', onclick:()=>{
        delete state.photos[key];
        preview.innerHTML = ''; preview.textContent = 'Belum ada foto';
        slot.classList.remove('filled');
        document.getElementById(inputId).value = '';
      }}, '✕')
    )
  );
  return slot;
}

/* ------------------------------ MAIN RENDER ------------------------------- */

function renderForm(){
  const root = document.getElementById('formSections');
  root.innerHTML = '';

  // 1. Info Lokasi
  const infoWrap = el('div',{},
    el('div',{class:'row2'},
      simpleField('nama_lokasi','Nama Lokasi','text',{placeholder:'contoh: KANTOR DESA'}),
      simpleField('site_id','Site ID','text',{placeholder:'contoh: ANDKN123N4'})
    ),
    simpleField('alamat','Alamat / Lokasi Site','textarea'),
    el('div',{class:'row3'},
      simpleField('provinsi','Provinsi','text'),
      simpleField('kabupaten','Kabupaten','text'),
      simpleField('kecamatan','Kecamatan','text')
    ),
    el('div',{class:'row2'},
      simpleField('desa','Kelurahan / Desa','text'),
      simpleField('koordinat','Koordinat (Lat, Long)','text',{hint:'Tekan tombol di bawah untuk isi otomatis dari lokasi HP.'})
    ),
    el('button',{type:'button', class:'btn btn-outline btn-sm', onclick:fillGeolocation}, '📍 Ambil Koordinat Saat Ini'),
    el('div',{class:'row2', style:'margin-top:10px;'},
      simpleField('petugas','Nama Petugas / Teknisi','text'),
      simpleField('kegiatan','Jenis Kegiatan','select',{options:['CM (Corrective Maintenance)','PM (Preventive Maintenance)']})
    ),
    el('div',{class:'row3'},
      simpleField('tanggal','Tanggal Kunjungan','date'),
      simpleField('waktu_datang','Waktu Kedatangan','time'),
      simpleField('waktu_selesai','Waktu Penyelesaian','time')
    ),
    simpleField('solusi_teknologi','Solusi Teknologi','text',{default:'Akses Internet VSAT', hint:'Boleh diubah sesuai kondisi site.'}),
    simpleField('ringkasan','Penjelasan Ringkas Kondisi / Gangguan','textarea')
  );
  root.appendChild(sectionCard(1,'Informasi Lokasi & Kunjungan','Data dasar site yang dikunjungi', infoWrap));

  // 2. Status Outdoor
  const outdoorWrap = el('div',{});
  OUTDOOR_ITEMS.forEach(it=> outdoorWrap.appendChild(compareItemEl(it.key, it.label, {meterLabel: it.meterLabel})));
  outdoorWrap.appendChild(customRowsBlock('outdoor','Tambah Perangkat Outdoor Lain','Nama perangkat outdoor lainnya'));
  root.appendChild(sectionCard(2,'Status Perangkat Outdoor','Kondisi sebelum & sesudah kunjungan', outdoorWrap));

  // 3. Status Indoor
  const indoorWrap = el('div',{});
  INDOOR_ITEMS.forEach(it=> indoorWrap.appendChild(compareItemEl(it.key, it.label, {})));
  indoorWrap.appendChild(customRowsBlock('indoor','Tambah Perangkat Indoor Lain','Nama perangkat indoor lainnya'));
  root.appendChild(sectionCard(3,'Status Perangkat Indoor','Kondisi sebelum & sesudah kunjungan', indoorWrap));

  // 4. Parameter Jaringan
  const paramWrap = el('div',{});
  PARAMETER_ITEMS.forEach(it=> paramWrap.appendChild(parameterItemEl(it)));
  root.appendChild(sectionCard(4,'Parameter Jaringan','Hasil pengukuran sebelum & sesudah', paramWrap));

  // 5. Status Power
  const powerWrap = el('div',{});
  POWER_ITEMS.forEach(it=> powerWrap.appendChild(compareItemEl(it.key, it.label, {})));
  POWER_PARAMS.forEach(it=> powerWrap.appendChild(parameterItemEl(it)));
  root.appendChild(sectionCard(5,'Status Power','Kondisi & parameter kelistrikan', powerWrap));

  // 6. Alternatif Jaringan
  const altWrap = el('div',{});
  ALT_NETWORK_ITEMS.forEach(it=> altWrap.appendChild(altNetworkItemEl(it)));
  root.appendChild(sectionCard(6,'Jaringan Alternatif di Sekitar Lokasi','Ketersediaan jaringan lain', altWrap));

  // 7. Perangkat Terpasang
  const devWrap = el('div',{});
  DEVICE_ITEMS.forEach(it=>{
    state.fields['dev_sn_before_'+it.key] = state.fields['dev_sn_before_'+it.key] || '';
    state.fields['dev_sn_after_'+it.key] = state.fields['dev_sn_after_'+it.key] || '';
    devWrap.appendChild(el('div',{class:'compare-item'},
      el('div',{class:'ci-label'}, it.label + ' — S/N / MAC / Alias'),
      el('div',{class:'ci-grid'},
        el('div',{class:'ci-col'}, el('span',{class:'ci-tag'},'Sebelum'),
          el('input',{type:'text', value:state.fields['dev_sn_before_'+it.key],
            oninput:e=>state.fields['dev_sn_before_'+it.key]=e.target.value})),
        el('div',{class:'ci-col'}, el('span',{class:'ci-tag'},'Sesudah'),
          el('input',{type:'text', value:state.fields['dev_sn_after_'+it.key],
            oninput:e=>state.fields['dev_sn_after_'+it.key]=e.target.value}))
      )
    ));
  });
  const devCustomWrap = el('div',{class:'custom-row-list'});
  function redrawDevCustom(){
    devCustomWrap.innerHTML = '';
    state.customRows.device.forEach((row, idx)=>{
      row.snBefore = row.snBefore || ''; row.snAfter = row.snAfter || '';
      devCustomWrap.appendChild(el('div',{class:'compare-item', style:'position:relative;'},
        el('button',{type:'button', class:'remove-x', onclick:()=>{ state.customRows.device.splice(idx,1); redrawDevCustom(); }}, '✕'),
        el('input',{type:'text', placeholder:'Nama perangkat lainnya', value:row.name, style:'margin-bottom:8px;',
          oninput:e=>row.name=e.target.value}),
        el('div',{class:'ci-grid'},
          el('div',{class:'ci-col'}, el('span',{class:'ci-tag'},'Sebelum'),
            el('input',{type:'text', value:row.snBefore, oninput:e=>row.snBefore=e.target.value})),
          el('div',{class:'ci-col'}, el('span',{class:'ci-tag'},'Sesudah'),
            el('input',{type:'text', value:row.snAfter, oninput:e=>row.snAfter=e.target.value}))
        )
      ));
    });
    devCustomWrap.appendChild(el('button',{type:'button', class:'add-link', onclick:()=>{
      state.customRows.device.push({ id: Date.now()+Math.random().toString(16).slice(2), name:'' });
      redrawDevCustom();
    }}, '+ Tambah Perangkat Lain'));
  }
  redrawDevCustom();
  devWrap.appendChild(devCustomWrap);
  root.appendChild(sectionCard(7,'Perangkat Terpasang','Detail Serial Number / MAC / Alias', devWrap));

  // 8. Photos
  const photoWrap = el('div',{class:'photo-grid'});
  PHOTO_ITEMS.forEach(([key,label])=> photoWrap.appendChild(photoSlotEl(key,label)));
  root.appendChild(sectionCard(8,'Dokumentasi Foto','Tekan tombol kamera untuk memotret langsung atau upload dari galeri', photoWrap));
}

function fillGeolocation(){
  if(!navigator.geolocation){ alert('Perangkat tidak mendukung geolocation.'); return; }
  navigator.geolocation.getCurrentPosition(pos=>{
    const val = pos.coords.latitude.toFixed(6)+', '+pos.coords.longitude.toFixed(6);
    state.fields['koordinat'] = val;
    const inp = document.getElementById('f_koordinat');
    if(inp) inp.value = val;
  }, err=>{ alert('Gagal mengambil lokasi: '+err.message); });
}

document.addEventListener('DOMContentLoaded', ()=>{
  document.getElementById('cfgClientId').value = getCfg(CFG_KEYS.clientId, '');
  document.getElementById('cfgFolderId').value = getCfg(CFG_KEYS.folderId, '1LfDpdzDnpj2FuQ0bz3Q9xCMRBFPbDYVa');
  renderForm();
  initGoogleAuth();
});

/* --------------------------- GOOGLE AUTH / DRIVE --------------------------- */

let tokenClient = null;
let accessToken = null;
let tokenExpiry = 0;

function log(msg){
  const box = document.getElementById('statusLog');
  box.classList.add('show');
  box.textContent += msg + '\n';
  box.scrollTop = box.scrollHeight;
}

function initGoogleAuth(){
  const clientId = getCfg(CFG_KEYS.clientId, '');
  const holder = document.getElementById('gsiBtnHolder');
  holder.innerHTML = '';
  if(!clientId){
    holder.appendChild(el('span',{class:'hint'}, 'Isi Client ID di Pengaturan dulu ⚙'));
    return;
  }
  if(!window.google || !google.accounts || !google.accounts.oauth2){
    holder.appendChild(el('span',{class:'hint'}, 'Memuat Google...'));
    setTimeout(initGoogleAuth, 500);
    return;
  }
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: clientId,
    scope: 'https://www.googleapis.com/auth/drive',
    callback: (resp)=>{
      if(resp.error){ log('Login gagal: '+resp.error); return; }
      accessToken = resp.access_token;
      tokenExpiry = Date.now() + (resp.expires_in*1000 - 60000);
      document.getElementById('authStatusText').textContent = 'Login Google berhasil ✓';
    }
  });
  const btn = el('button',{type:'button', class:'btn btn-outline btn-sm', onclick: signIn}, 'Login dengan Google');
  holder.appendChild(btn);
}

function signIn(){
  if(!tokenClient){ initGoogleAuth(); return; }
  tokenClient.requestAccessToken({ prompt: accessToken ? '' : 'consent' });
}

function ensureSignedIn(){
  return new Promise((resolve, reject)=>{
    if(accessToken && Date.now() < tokenExpiry){ resolve(accessToken); return; }
    if(!tokenClient){ reject(new Error('Client ID belum diatur. Buka Pengaturan.')); return; }
    tokenClient.callback = (resp)=>{
      if(resp.error){ reject(new Error('Login gagal: '+resp.error)); return; }
      accessToken = resp.access_token;
      tokenExpiry = Date.now() + (resp.expires_in*1000 - 60000);
      document.getElementById('authStatusText').textContent = 'Login Google berhasil ✓';
      resolve(accessToken);
    };
    tokenClient.requestAccessToken({ prompt: accessToken ? '' : 'consent' });
  });
}

async function driveFindFolder(name, parentId){
  const q = encodeURIComponent(`name='${name.replace(/'/g,"\\'")}' and '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`);
  const res = await fetch(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`, {
    headers: { Authorization: 'Bearer '+accessToken }
  });
  const data = await res.json();
  if(data.files && data.files.length) return data.files[0].id;
  return null;
}

async function driveCreateFolder(name, parentId){
  const res = await fetch('https://www.googleapis.com/drive/v3/files', {
    method:'POST',
    headers: { Authorization:'Bearer '+accessToken, 'Content-Type':'application/json' },
    body: JSON.stringify({ name, mimeType:'application/vnd.google-apps.folder', parents:[parentId] })
  });
  const data = await res.json();
  if(!res.ok) throw new Error('Gagal buat folder: '+(data.error && data.error.message));
  return data.id;
}

async function ensureFolder(name, parentId){
  let id = await driveFindFolder(name, parentId);
  if(!id) id = await driveCreateFolder(name, parentId);
  return id;
}

async function driveUploadFile(blob, filename, parentId, mimeType){
  const metadata = { name: filename, parents: [parentId] };
  const form = new FormData();
  form.append('metadata', new Blob([JSON.stringify(metadata)], {type:'application/json'}));
  form.append('file', blob, filename);
  const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink', {
    method:'POST',
    headers: { Authorization:'Bearer '+accessToken },
    body: form
  });
  const data = await res.json();
  if(!res.ok) throw new Error('Gagal upload file: '+(data.error && data.error.message));
  return data;
}

/* -------------------------------- DOCX BUILD -------------------------------- */

const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        WidthType, AlignmentType, HeadingLevel, ImageRun, ShadingType,
        BorderStyle, VerticalAlign } = docx;

const CELL_BORDER = {
  top:{style:BorderStyle.SINGLE,size:2,color:'CCCCCC'},
  bottom:{style:BorderStyle.SINGLE,size:2,color:'CCCCCC'},
  left:{style:BorderStyle.SINGLE,size:2,color:'CCCCCC'},
  right:{style:BorderStyle.SINGLE,size:2,color:'CCCCCC'},
};

function tCell(text, opts){
  opts = opts || {};
  return new TableCell({
    width:{ size: opts.width || 3000, type: WidthType.DXA },
    shading: opts.shade ? { type: ShadingType.CLEAR, fill: opts.shade } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    borders: CELL_BORDER,
    children:[ new Paragraph({
      children:[ new TextRun({ text: String(text==null?'':text), bold: !!opts.bold, color: opts.color }) ]
    }) ]
  });
}

function headingPara(text, level){
  return new Paragraph({ heading: level || HeadingLevel.HEADING_2, spacing:{ before:280, after:120 },
    children:[ new TextRun({ text, bold:true, color:'12395C' }) ] });
}

function compareTable(rows){
  // rows: [{label, before, after}]
  const trs = [ new TableRow({ tableHeader:true, children:[
    tCell('Item', {bold:true, shade:'D9E4F5', width:4500}),
    tCell('Sebelum', {bold:true, shade:'D9E4F5', width:2500}),
    tCell('Sesudah', {bold:true, shade:'D9E4F5', width:2500}),
  ]})];
  rows.forEach(r=>{
    trs.push(new TableRow({ children:[
      tCell(r.label, {width:4500}),
      tCell(r.before || '-', {width:2500}),
      tCell(r.after || '-', {width:2500}),
    ]}));
  });
  return new Table({ width:{size:100,type:WidthType.PERCENTAGE}, rows: trs });
}

function fmtCompare(itemKey){
  const c = state.compare[itemKey] || {};
  return { before: c.before ? c.before.toUpperCase() : '-', after: c.after ? c.after.toUpperCase() : '-' };
}

function buildOutdoorRows(){
  const rows = OUTDOOR_ITEMS.map(it=>{
    const f = fmtCompare(it.key);
    const c = state.compare[it.key] || {};
    let before = f.before, after = f.after;
    if(it.meterLabel){
      before += c.extra && c.extra.meter_before ? ` (${c.extra.meter_before} m)` : '';
      after += c.extra && c.extra.meter_after ? ` (${c.extra.meter_after} m)` : '';
    }
    return { label: it.label, before, after };
  });
  state.customRows.outdoor.forEach(row=>{
    const f = fmtCompare('custom_outdoor_'+row.id);
    rows.push({ label: row.name || '(perangkat lain)', before: f.before, after: f.after });
  });
  return rows;
}

function buildIndoorRows(){
  const rows = INDOOR_ITEMS.map(it=>{
    const f = fmtCompare(it.key);
    return { label: it.label, before: f.before, after: f.after };
  });
  state.customRows.indoor.forEach(row=>{
    const f = fmtCompare('custom_indoor_'+row.id);
    rows.push({ label: row.name || '(perangkat lain)', before: f.before, after: f.after });
  });
  return rows;
}

function buildParamRows(items){
  return items.map(it=>{
    const c = state.compare[it.key] || {};
    let before, after;
    if(it.kind === 'ping'){
      const b = c.before || {}, a = c.after || {};
      before = `Min:${b.min||'-'} Max:${b.max||'-'} Avg:${b.avg||'-'} Loss:${b.loss||'-'}%`;
      after = `Min:${a.min||'-'} Max:${a.max||'-'} Avg:${a.avg||'-'} Loss:${a.loss||'-'}%`;
    } else {
      before = c.before_val || '-';
      after = c.after_val || '-';
    }
    return { label: it.label, before, after };
  });
}

function altNetworkTable(){
  const trs = [ new TableRow({ tableHeader:true, children:[
    tCell('Jenis Jaringan', {bold:true, shade:'D9E4F5', width:3500}),
    tCell('Status', {bold:true, shade:'D9E4F5', width:2000}),
    tCell('Operator / Keterangan', {bold:true, shade:'D9E4F5', width:3500}),
  ]})];
  ALT_NETWORK_ITEMS.forEach(it=>{
    const c = state.compare[it.key] || {};
    trs.push(new TableRow({ children:[
      tCell(it.label, {width:3500}),
      tCell(c.status || '-', {width:2000}),
      tCell(c.operator || '-', {width:3500}),
    ]}));
  });
  return new Table({ width:{size:100,type:WidthType.PERCENTAGE}, rows: trs });
}

function deviceTable(){
  const trs = [ new TableRow({ tableHeader:true, children:[
    tCell('Perangkat', {bold:true, shade:'D9E4F5', width:3000}),
    tCell('S/N / MAC / Alias (Sebelum)', {bold:true, shade:'D9E4F5', width:3500}),
    tCell('S/N / MAC / Alias (Sesudah)', {bold:true, shade:'D9E4F5', width:3500}),
  ]})];
  DEVICE_ITEMS.forEach(it=>{
    trs.push(new TableRow({ children:[
      tCell(it.label, {width:3000}),
      tCell(state.fields['dev_sn_before_'+it.key] || '-', {width:3500}),
      tCell(state.fields['dev_sn_after_'+it.key] || '-', {width:3500}),
    ]}));
  });
  state.customRows.device.forEach(row=>{
    trs.push(new TableRow({ children:[
      tCell(row.name || '(perangkat lain)', {width:3000}),
      tCell(row.snBefore || '-', {width:3500}),
      tCell(row.snAfter || '-', {width:3500}),
    ]}));
  });
  return new Table({ width:{size:100,type:WidthType.PERCENTAGE}, rows: trs });
}

async function photoParagraphs(){
  const out = [];
  for(const [key,label] of PHOTO_ITEMS){
    out.push(new Paragraph({ spacing:{before:200,after:80}, children:[ new TextRun({ text: label, bold:true, size:20 }) ] }));
    const p = state.photos[key];
    if(p && p.blob){
      const buf = new Uint8Array(await p.blob.arrayBuffer());
      const maxW = 380;
      const ratio = (p.width && p.height) ? p.height/p.width : 0.75;
      const w = maxW, h = Math.round(maxW*ratio);
      out.push(new Paragraph({ children:[ new ImageRun({ data: buf, transformation:{ width:w, height:h }, type:'jpg' }) ] }));
    } else {
      out.push(new Paragraph({ children:[ new TextRun({ text:'(Foto belum diisi)', italics:true, color:'999999' }) ] }));
    }
  }
  return out;
}

async function buildDocx(){
  const f = state.fields;
  const infoRows = [
    ['Nama Lokasi', f.nama_lokasi], ['Site ID', f.site_id],
    ['Alamat / Lokasi', f.alamat], ['Provinsi', f.provinsi],
    ['Kabupaten', f.kabupaten], ['Kecamatan', f.kecamatan],
    ['Kelurahan / Desa', f.desa], ['Koordinat', f.koordinat],
    ['Nama Petugas', f.petugas], ['Jenis Kegiatan', f.kegiatan],
    ['Tanggal Kunjungan', f.tanggal], ['Waktu Kedatangan', f.waktu_datang],
    ['Waktu Penyelesaian', f.waktu_selesai], ['Solusi Teknologi', f.solusi_teknologi],
  ];
  const infoTrs = infoRows.map(([label,val])=> new TableRow({ children:[
    tCell(label, {bold:true, width:3000, shade:'F2F5FA'}),
    tCell(val || '-', {width:6000}),
  ]}));
  const infoTable = new Table({ width:{size:100,type:WidthType.PERCENTAGE}, rows: infoTrs });

  const photoParas = await photoParagraphs();

  const doc = new Document({
    sections: [{
      properties:{},
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:60}, children:[
          new TextRun({ text:'LAPORAN KUNJUNGAN TEKNISI', bold:true, size:32, color:'12395C' }) ] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:240}, children:[
          new TextRun({ text: (f.nama_lokasi||'-') + ' — ' + (f.tanggal||'-'), size:22, color:'555555' }) ] }),

        headingPara('I. Informasi Lokasi & Kunjungan'), infoTable,

        headingPara('II. Penjelasan Ringkas Kondisi / Gangguan'),
        new Paragraph({ spacing:{after:200}, children:[ new TextRun({ text: f.ringkasan || '-' }) ] }),

        headingPara('III. Status Perangkat Outdoor'), compareTable(buildOutdoorRows()),
        new Paragraph({ text:'', spacing:{after:160} }),

        headingPara('IV. Status Perangkat Indoor'), compareTable(buildIndoorRows()),
        new Paragraph({ text:'', spacing:{after:160} }),

        headingPara('V. Parameter Jaringan'), compareTable(buildParamRows(PARAMETER_ITEMS)),
        new Paragraph({ text:'', spacing:{after:160} }),

        headingPara('VI. Status Power'),
        compareTable(POWER_ITEMS.map(it=>{ const fc = fmtCompare(it.key); return { label:it.label, before:fc.before, after:fc.after }; })),
        new Paragraph({ text:'', spacing:{after:120} }),
        compareTable(buildParamRows(POWER_PARAMS)),
        new Paragraph({ text:'', spacing:{after:160} }),

        headingPara('VII. Jaringan Alternatif di Sekitar Lokasi'), altNetworkTable(),
        new Paragraph({ text:'', spacing:{after:160} }),

        headingPara('VIII. Perangkat Terpasang'), deviceTable(),
        new Paragraph({ text:'', spacing:{after:160} }),

        headingPara('IX. Dokumentasi Foto'),
        ...photoParas,
      ]
    }]
  });
  return doc;
}

/* ------------------------------- SUBMIT FLOW -------------------------------- */

const BULAN_ID = ['JANUARI','FEBRUARI','MARET','APRIL','MEI','JUNI','JULI','AGUSTUS','SEPTEMBER','OKTOBER','NOVEMBER','DESEMBER'];

function namaBulanDariTanggal(tglStr){
  if(!tglStr) return '';
  const d = new Date(tglStr);
  if(isNaN(d.getTime())) return '';
  return BULAN_ID[d.getMonth()];
}

function validateRequired(){
  const f = state.fields;
  if(!f.nama_lokasi || !f.nama_lokasi.trim()) return 'Nama Lokasi wajib diisi.';
  if(!f.site_id || !f.site_id.trim()) return 'Site ID wajib diisi.';
  if(!f.tanggal) return 'Tanggal Kunjungan wajib diisi.';
  return null;
}

function setProgress(pct){
  const wrap = document.getElementById('progressWrap');
  wrap.style.display = pct>0 && pct<100 ? 'block' : (pct>=100 ? 'none' : 'none');
  document.getElementById('progressBar').style.width = pct+'%';
}

function fileBaseName(){
  const f = state.fields;
  return `LAPORAN KUNJUNGAN ${f.tanggal || ''} ${(f.nama_lokasi||'').toUpperCase()}`.trim();
}

function folderName(){
  const f = state.fields;
  const bulan = namaBulanDariTanggal(f.tanggal);
  return `LAPORAN KUNJUNGAN ${bulan} ${(f.nama_lokasi||'').toUpperCase()}`.trim();
}

async function downloadLocalCopy(){
  const err = validateRequired();
  if(err){ alert(err); return; }
  const btn = document.getElementById('submitBtn');
  try{
    log('Membuat file Word...');
    const doc = await buildDocx();
    const blob = await Packer.toBlob(doc);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileBaseName()+'.docx';
    a.click();
    log('File tersimpan ke perangkat ✓');
  }catch(e){
    log('Gagal membuat file: '+e.message);
    alert('Gagal membuat file: '+e.message);
  }
}

async function submitReport(){
  const err = validateRequired();
  if(err){ alert(err); return; }
  const btn = document.getElementById('submitBtn');
  const parentId = getCfg(CFG_KEYS.folderId, '');
  if(!parentId){ alert('ID Folder Google Drive belum diatur. Buka Pengaturan.'); return; }
  if(!getCfg(CFG_KEYS.clientId,'')){ alert('Google Client ID belum diatur. Buka Pengaturan.'); return; }

  btn.disabled = true;
  const origText = btn.textContent;
  document.getElementById('statusLog').textContent = '';
  document.getElementById('statusLog').classList.add('show');

  try{
    btn.textContent = 'Login Google...';
    setProgress(10);
    log('Menghubungkan ke akun Google...');
    await ensureSignedIn();

    btn.textContent = 'Menyusun laporan...';
    setProgress(30);
    log('Menyusun dokumen Word...');
    const doc = await buildDocx();
    const blob = await Packer.toBlob(doc);

    btn.textContent = 'Menyiapkan folder Drive...';
    setProgress(55);
    const fName = folderName();
    log('Menyiapkan folder: '+fName);
    const targetFolderId = await ensureFolder(fName, parentId);

    btn.textContent = 'Mengupload...';
    setProgress(75);
    const finalName = fileBaseName()+'.docx';
    log('Mengupload file: '+finalName);
    const result = await driveUploadFile(blob, finalName, targetFolderId, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

    setProgress(100);
    log('Berhasil diupload ✓');
    if(result.webViewLink) log('Link: '+result.webViewLink);
    alert('Laporan berhasil diupload ke Google Drive!\nFolder: '+fName);
  }catch(e){
    log('ERROR: '+e.message);
    alert('Gagal upload: '+e.message+'\n\nCoba tombol "Simpan ke HP/Laptop" sebagai cadangan.');
  }finally{
    btn.disabled = false;
    btn.textContent = origText;
    setTimeout(()=>setProgress(0), 1500);
  }
}
