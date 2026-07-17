<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Laporan Kunjungan Teknisi</title>
<script src="https://accounts.google.com/gsi/client" async defer></script>
<script src="https://unpkg.com/docx@8.5.0/build/index.iife.js"></script>
<style>
  :root{
    --navy:#0f2942; --navy-2:#153a5b; --accent:#2f6fed; --accent-2:#1c9e6e;
    --bg:#f4f6f9; --card:#ffffff; --border:#e2e6ec; --text:#1b2430; --muted:#6b7684;
    --danger:#d64545; --warn:#c9821a; --radius:12px;
  }
  *{box-sizing:border-box;}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--text);}
  header.app{background:linear-gradient(135deg,var(--navy),var(--navy-2));color:#fff;padding:20px 16px 28px;}
  header.app h1{margin:0 0 4px;font-size:1.25rem;}
  header.app p{margin:0;font-size:.85rem;opacity:.85;}
  .wrap{max-width:820px;margin:-16px auto 60px;padding:0 12px;}
  .card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);padding:18px;margin-bottom:16px;box-shadow:0 1px 3px rgba(15,41,66,.06);}
  .card h2{font-size:1rem;margin:0 0 4px;color:var(--navy);}
  .card .sub{color:var(--muted);font-size:.8rem;margin:0 0 14px;}
  .section-num{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;background:var(--accent);color:#fff;font-size:.7rem;font-weight:700;margin-right:8px;}
  label{display:block;font-size:.82rem;font-weight:600;color:var(--navy);margin:10px 0 4px;}
  label .opt{color:var(--muted);font-weight:400;}
  input[type=text],input[type=date],input[type=time],input[type=number],textarea,select{
    width:100%;padding:9px 10px;border:1px solid var(--border);border-radius:8px;font-size:.9rem;background:#fbfcfd;color:var(--text);font-family:inherit;
  }
  textarea{min-height:70px;resize:vertical;}
  input:focus,textarea:focus,select:focus{outline:2px solid var(--accent);outline-offset:1px;background:#fff;}
  .row2{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
  .row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;}
  @media (max-width:560px){.row2,.row3{grid-template-columns:1fr;}}
  .hint{font-size:.72rem;color:var(--muted);margin-top:2px;}
  button{font-family:inherit;cursor:pointer;}
  .btn{border:none;border-radius:9px;padding:10px 16px;font-size:.88rem;font-weight:600;}
  .btn-primary{background:var(--accent);color:#fff;}
  .btn-primary:hover{background:#265ecb;}
  .btn-outline{background:#fff;border:1px solid var(--border);color:var(--navy);}
  .btn-outline:hover{background:#f0f3f8;}
  .btn-danger{background:#fdeceb;color:var(--danger);border:1px solid #f6c9c5;}
  .btn-sm{padding:6px 10px;font-size:.78rem;border-radius:7px;}
  .btn-block{width:100%;}
  .btn[disabled]{opacity:.5;cursor:not-allowed;}

  /* status compare rows */
  .compare-item{border:1px solid var(--border);border-radius:10px;padding:10px 12px;margin-bottom:8px;background:#fbfcfd;}
  .compare-item .ci-label{font-weight:600;font-size:.86rem;margin-bottom:8px;color:var(--text);}
  .ci-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
  @media (max-width:520px){.ci-grid{grid-template-columns:1fr;}}
  .ci-col{background:#fff;border:1px solid var(--border);border-radius:8px;padding:8px;}
  .ci-col .ci-tag{font-size:.68rem;font-weight:700;letter-spacing:.04em;color:var(--muted);text-transform:uppercase;margin-bottom:6px;display:block;}
  .toggle-group{display:flex;gap:6px;}
  .toggle-btn{flex:1;padding:7px 4px;border:1px solid var(--border);border-radius:7px;background:#fff;font-size:.8rem;font-weight:600;color:var(--muted);text-align:center;}
  .toggle-btn.active-baik{background:#e5f7ee;border-color:var(--accent-2);color:var(--accent-2);}
  .toggle-btn.active-buruk{background:#fdeceb;border-color:var(--danger);color:var(--danger);}
  .extra-field{margin-top:8px;}
  .extra-field input{padding:7px 8px;font-size:.83rem;}

  .custom-row-list .compare-item{position:relative;}
  .remove-x{position:absolute;top:8px;right:8px;background:none;border:none;color:var(--muted);font-size:1rem;line-height:1;padding:2px 6px;}
  .remove-x:hover{color:var(--danger);}
  .add-link{background:none;border:1px dashed var(--accent);color:var(--accent);border-radius:8px;padding:8px;width:100%;font-size:.82rem;font-weight:600;}

  /* photo grid */
  .photo-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
  @media (max-width:520px){.photo-grid{grid-template-columns:1fr;}}
  .photo-slot{border:1px solid var(--border);border-radius:10px;padding:10px;background:#fbfcfd;}
  .photo-slot .ps-label{font-size:.78rem;font-weight:600;margin-bottom:8px;min-height:32px;}
  .photo-preview{width:100%;height:130px;border-radius:8px;background:#eef1f5 center/cover no-repeat;display:flex;align-items:center;justify-content:center;color:var(--muted);font-size:.72rem;margin-bottom:8px;overflow:hidden;border:1px dashed var(--border);}
  .photo-preview img{width:100%;height:100%;object-fit:cover;}
  .photo-actions{display:flex;gap:6px;flex-wrap:wrap;}
  .photo-actions label.upl{flex:1 1 40%;min-width:90px;text-align:center;background:var(--accent);color:#fff;border-radius:7px;padding:7px 4px;font-size:.72rem;font-weight:600;}
  .photo-actions .clear-btn{background:#fff;border:1px solid var(--border);border-radius:7px;padding:7px 8px;font-size:.75rem;color:var(--muted);}
  .photo-slot input[type=file]{display:none;}
  .photo-slot.filled{border-color:var(--accent-2);}

  .badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:.68rem;font-weight:700;}
  .badge-ok{background:#e5f7ee;color:var(--accent-2);}
  .badge-off{background:#f1eef8;color:#7a5fc2;}

  #authBar{position:sticky;top:0;z-index:20;background:#fff;border-bottom:1px solid var(--border);padding:10px 14px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;}
  #authBar .status-text{font-size:.78rem;color:var(--muted);}
  #gsiBtnHolder{min-height:36px;}

  #statusLog{white-space:pre-wrap;font-size:.78rem;background:#0f1720;color:#c9d6e3;border-radius:8px;padding:10px;max-height:180px;overflow:auto;display:none;}
  #statusLog.show{display:block;}

  .config-toggle{font-size:.75rem;color:var(--accent);background:none;border:none;text-decoration:underline;padding:0;margin-top:4px;}
  #configPanel{display:none;margin-top:10px;border-top:1px dashed var(--border);padding-top:10px;}
  #configPanel.show{display:block;}

  .footer-actions{position:sticky;bottom:0;background:linear-gradient(to top,var(--bg) 60%,transparent);padding:14px 0 6px;}
  .footer-actions .inner{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:12px;display:flex;gap:8px;box-shadow:0 -2px 10px rgba(15,41,66,.08);}
  .progress-bar{height:6px;background:#eef1f5;border-radius:4px;overflow:hidden;margin:6px 0 0;}
  .progress-bar > div{height:100%;background:var(--accent);width:0%;transition:width .2s;}
</style>
</head>
<body>

<div id="authBar">
  <div id="gsiBtnHolder"></div>
  <span class="status-text" id="authStatusText">Belum login Google.</span>
  <button class="config-toggle" type="button" onclick="toggleConfig()">Pengaturan</button>
</div>

<div class="wrap" id="configWrap" style="margin-top:12px;">
  <div class="card" style="padding:12px 14px;">
    <button class="config-toggle" type="button" onclick="toggleConfig()">⚙ Pengaturan Aplikasi (Client ID / Folder Drive)</button>
    <div id="configPanel">
      <label>Google OAuth Client ID</label>
      <input type="text" id="cfgClientId" placeholder="xxxxxxxx.apps.googleusercontent.com">
      <p class="hint">Diminta sekali per perangkat. Lihat PANDUAN-SETUP.md untuk cara membuatnya di Google Cloud Console.</p>
      <label>ID Folder Google Drive Tujuan (folder induk semua laporan)</label>
      <input type="text" id="cfgFolderId" value="1LfDpdzDnpj2FuQ0bz3Q9xCMRBFPbDYVa">
      <p class="hint">Diambil otomatis dari link folder yang diberikan. Ganti jika perlu.</p>
      <button class="btn btn-outline btn-sm" type="button" onclick="saveConfig()">Simpan Pengaturan</button>
      <span id="cfgSavedMsg" class="hint" style="color:var(--accent-2);"></span>
    </div>
  </div>
</div>

<header class="app">
  <h1>📡 Laporan Kunjungan Teknisi</h1>
  <p>Isi detail kunjungan site, upload foto, dan laporan otomatis tersimpan ke Google Drive dalam format Word (.docx).</p>
</header>

<div class="wrap">
  <form id="mainForm" onsubmit="return false;">

  <!-- form sections injected by JS -->
  <div id="formSections"></div>

  </form>

  <div class="card">
    <div id="statusLog"></div>
    <div class="progress-bar" id="progressWrap" style="display:none;"><div id="progressBar"></div></div>
  </div>
</div>

<div class="footer-actions">
  <div class="wrap" style="margin:0 auto;">
    <div class="inner">
      <button class="btn btn-outline" type="button" style="flex:1;" onclick="downloadLocalCopy()">⬇ Simpan ke HP/Laptop</button>
      <button class="btn btn-primary btn-block" style="flex:2;" id="submitBtn" onclick="submitReport()">✅ Selesai &amp; Upload ke Drive</button>
    </div>
  </div>
</div>

<script src="app.js"></script>
</body>
</html>
