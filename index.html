<!DOCTYPE html>
<html lang="id">
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css">
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>

  <style>
    :root { --primary: #002D57; --accent: #00A8E8; --bg: #F4F7FE; --danger: #DC3545; --success: #198754; --warning: #FFC107; --orange: #FD7E14; }
    body { font-family: 'Plus Jakarta Sans', sans-serif; background-color: var(--bg); padding-bottom: 90px; color: #1A202C; overflow-x: hidden; }
    .header-box { background: var(--primary); color: white; padding: 18px 20px; border-radius: 0 0 25px 25px; }
    .card-mewah { background: white; border-radius: 20px; padding: 15px; margin-bottom: 12px; border: none; box-shadow: 0 4px 15px rgba(0,0,0,0.02); }
    .hero-stat { background: linear-gradient(135deg, #002D57 0%, #00509D 100%); color: white; padding: 15px !important; border-radius: 25px; }
    .jab-pct { background: var(--primary); color: white; border-radius: 12px; font-size: 22px !important; font-weight: 800; padding: 6px 0; display: block; margin-bottom: 10px; }
    .jab-detail { font-size: 10px; line-height: 1.4; text-align: left; border-top: 1px dashed #eee; padding-top: 8px; color: #64748B; }
    .jab-detail b { color: var(--primary); font-size: 11px; }
    .cair-row { background: #f8fafc; padding: 12px 16px; border-radius: 15px; margin-bottom: 8px; display: flex; justify-content: space-between; align-items: center; border-left: 4px solid var(--accent); }
    .total-row { background: #002D57 !important; color: #FFD700 !important; border-left: 4px solid #FFD700 !important; font-weight: 800; }
    .p-title { font-size: 8px; font-weight: 800; color: #94A3B8; text-transform: uppercase; }
    .p-val { font-size: 15px; font-weight: 800; color: var(--primary); display: block; }
    .funnel-step { background: #f1f5f9; padding: 10px 15px; border-radius: 12px; margin-bottom: 8px; border-left: 6px solid var(--primary); display: flex; justify-content: space-between; align-items: center; }
    .nav-fix { position: fixed; bottom: 12px; left: 12px; right: 12px; background: white; display: flex; justify-content: space-around; padding: 12px 0; border-radius: 22px; box-shadow: 0 10px 35px rgba(0,0,0,0.12); z-index: 1000; }
    .nav-item { text-align: center; color: #94A3B8; text-decoration: none; font-size: 10px; flex: 1; font-weight: 700; cursor: pointer; }
    .nav-item.active { color: var(--primary); }
    #loader-global { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: white; z-index: 9999; display: flex; flex-direction: column; justify-content: center; align-items: center; }
  </style>
</head>
<body onload="initApp()">

  <div id="loader-global"><div class="spinner-grow text-primary"></div><p class="mt-3 fw-bold">Sinkronisasi CC...</p></div>

  <div id="app">
    <div class="header-box d-flex justify-content-between align-items-center">
      <div><p class="small mb-0 opacity-75 fw-bold" style="font-size:9px;">BPR KS</p><h6 class="fw-bold mb-0" id="u-nama">MEMUAT...</h6></div>
      <button class="btn btn-sm btn-outline-light rounded-pill px-3" style="font-size:10px; font-weight:800;" onclick="refreshData()">REFRESH</button>
    </div>

    <div class="p-3">
      <div id="v-home">
        <div class="card-mewah mb-2 shadow-sm">
          <div class="row g-1">
            <div class="col-6"><input type="date" id="f-start" class="form-control form-control-sm" onchange="runFilter()"></div>
            <div class="col-6"><input type="date" id="f-end" class="form-control form-control-sm" onchange="runFilter()"></div>
            <div class="col-6 mt-1"><select id="f-c" class="form-select form-select-sm" onchange="runFilter(true)"></select></div>
            <div class="col-6 mt-1"><select id="f-n" class="form-select form-select-sm" onchange="runFilter()"></select></div>
          </div>
        </div>
        <div class="card-mewah hero-stat text-center shadow mb-2">
          <div class="d-flex justify-content-center mb-1"><div style="width:130px; position:relative;"><canvas id="gChart"></canvas><div style="position:absolute; top:65%; left:50%; transform:translate(-50%,-50%); text-align:center;"><h3 class="fw-bold mb-0" id="g-p" style="font-size:24px;">0%</h3></div></div></div>
          <p class="small mb-0 opacity-75" style="font-size:10px; font-weight: 700;">TOTAL REALISASI :</p>
          <h2 class="fw-bold mb-1" id="g-r" style="font-size: 28px;">Rp 0</h2>
          <div class="mt-2 py-2" style="background: rgba(255,255,255,0.1); border-radius: 15px;">
            <p class="small mb-0 opacity-75" style="font-size:10px; font-weight: 700;">TOTAL TARGET :</p>
            <h4 class="fw-bold mb-0" id="g-t" style="font-size: 19px;">Rp 0</h4>
          </div>
        </div>
        <div class="row g-2 mb-3">
          <div class="col-4"><div class="card-mewah p-2 text-center" style="height:100%"><small class="fw-bold text-muted">CS</small><span class="jab-pct" id="cs-p">0%</span><div class="jab-detail">TARGET:<br><b id="cs-tgt">0</b><br>REALISASI:<br><b id="cs-ach">0</b></div></div></div>
          <div class="col-4"><div class="card-mewah p-2 text-center" style="height:100%"><small class="fw-bold text-muted">SPV</small><span class="jab-pct" id="spv-p" style="background:#10B981">0%</span><div class="jab-detail">TARGET:<br><b id="spv-tgt">0</b><br>REALISASI:<br><b id="spv-ach">0</b></div></div></div>
          <div class="col-4"><div class="card-mewah p-2 text-center" style="height:100%"><small class="fw-bold text-muted">TELE</small><span class="jab-pct" id="tele-p" style="background:#F59E0B">0%</span><div class="jab-detail">TARGET:<br><b id="tele-tgt">0</b><br>REALISASI:<br><b id="tele-ach">0</b></div></div></div>
        </div>
        <div id="cair-list"></div>
      </div>

      <div id="v-pipe" style="display:none">
        <div class="card-mewah mb-2 shadow-sm" style="border-top: 4px solid var(--primary);">
          <div class="row g-1">
            <div class="col-6"><input type="date" id="p-start" class="form-control form-control-sm" onchange="filterPipe()"></div>
            <div class="col-6"><input type="date" id="p-end" class="form-control form-control-sm" onchange="filterPipe()"></div>
            <div class="col-4 mt-1"><select id="p-c" class="form-select form-select-sm" onchange="syncPipelineFilter('cab')"></select></div>
            <div class="col-4 mt-1"><select id="p-spv" class="form-select form-select-sm" onchange="syncPipelineFilter('spv')"></select></div>
            <div class="col-4 mt-1"><select id="p-s" class="form-select form-select-sm" onchange="syncPipelineFilter('sales')"></select></div>
          </div>
        </div>
        <div class="row g-2 mb-2 text-center">
          <div class="col-4"><div class="card-mewah p-2"><span class="p-title">APK</span><b class="p-val" id="kp-apk">0</b></div></div>
          <div class="col-8"><div class="card-mewah p-2"><span class="p-title">PLAFOND PIPELINE</span><b class="p-val" id="kp-pla">Rp 0</b></div></div>
        </div>
        <div class="card-mewah mb-3"><div class="funnel-step"><span>APK NAIK</span><b id="fn-naik">0</b></div><div class="funnel-step" style="border-left-color:var(--success)"><span>APPROVE</span><b id="fn-app">0</b></div><div class="funnel-step" style="border-left-color:var(--accent)"><span>CAIR</span><b id="fn-cair">0</b></div></div>
        <div class="card-mewah p-0 overflow-hidden shadow-sm">
          <div class="p-2 border-bottom bg-light d-flex justify-content-between align-items-center"><span class="fw-bold small">LIST DEBITUR</span><input type="text" id="p-search" class="form-control form-control-sm w-50" placeholder="Cari..." onkeyup="filterTablePipe()"></div>
          <div style="max-height:400px; overflow-y:auto"><table class="table table-sm table-hover mb-0" style="font-size:10px;"><thead class="bg-dark text-white"><tr><th>DEBITUR</th><th>CAB/SPV</th><th>STATUS</th><th>PLAFOND</th></tr></thead><tbody id="pipe-body"></tbody></table></div>
        </div>
      </div>

      <div id="v-admin" style="display:none">
        <div class="card-mewah shadow">
          <h6 class="fw-bold mb-4 text-primary">ADMIN UPLOAD</h6>
          <input type="date" id="u-t" class="form-control mb-3">
          <select id="u-j" class="form-select mb-3">
            <option value="1">Achv (Pencapaian) CS</option><option value="2">Achv (Pencapaian) SPV</option><option value="3">Achv (Pencapaian) TELE</option>
            <option value="4">CAIR - (Pencairan CS)</option><option value="5">CAIR - (Pencairan SPV)</option><option value="6">CAIR - (Pencairan TELE)</option>
            <option value="7">Raw_Pipeline (Excel)</option>
          </select>
          <input type="file" id="u-f" class="form-control mb-4"><button class="btn btn-primary w-100 fw-bold py-3 shadow" onclick="execU()">UNGGAH DATA</button>
        </div>
      </div>
    </div>
  </div>

  <div class="nav-fix">
    <a onclick="switchPage('home')" class="nav-item active" id="n-h"><i class="bi bi-speedometer2" style="font-size:20px; display:block;"></i>Home</a>
    <a onclick="switchPage('pipe')" class="nav-item" id="n-p"><i class="bi bi-funnel-fill" style="font-size:20px; display:block;"></i>Pipeline</a>
    <a onclick="switchPage('admin')" class="nav-item" id="n-a"><i class="bi bi-shield-lock-fill" style="font-size:20px; display:block;"></i>Admin</a>
  </div>

  <script>
    let db = { achv:[], cair:[], pipeline:[], listCHome:[], listCPipe:[], listS:[], listSPV:[] }, myChart = null;

    function initApp() {
      google.script.run.withSuccessHandler(res => {
        db = res; document.getElementById('u-nama').innerText = res.u.nama;
        const fill = (id, list, txt) => {
          const s = document.getElementById(id); s.innerHTML = `<option value="">${txt}</option>`;
          [...new Set(list)].sort().forEach(v => s.add(new Option(v, v)));
        };
        fill('f-c', res.listCHome, "Semua Cabang"); fill('p-c', res.listCPipe, "Semua Cabang");
        fill('p-spv', res.listSPV, "Semua SPV"); fill('p-s', res.listS, "Semua Sales");
        runFilter(true); filterPipe();
        document.getElementById('loader-global').style.display = 'none';
      }).getDashboardData();
    }

    function refreshData() {
      Swal.fire({ title: 'Menyinkronkan...', allowOutsideClick: false, didOpen: () => { Swal.showLoading(); } });
      google.script.run.withSuccessHandler(res => {
        db = res; runFilter(true); filterPipe();
        Swal.fire({ icon: 'success', title: 'Terupdate', timer: 1000, showConfirmButton: false });
      }).getDashboardData();
    }

    function switchPage(p) {
      ['v-home','v-pipe','v-admin'].forEach(v => document.getElementById(v).style.display = 'none');
      ['n-h','n-p','n-a'].forEach(n => document.getElementById(n).classList.remove('active'));
      document.getElementById('v-'+p).style.display = 'block';
      document.getElementById('n-'+p.charAt(0)).classList.add('active');
    }

    function runFilter(isBranchChange = false) {
      const s = document.getElementById('f-start').value || "1900-01-01", e = document.getElementById('f-end').value || "2100-12-31";
      const c = document.getElementById('f-c').value, n = document.getElementById('f-n').value;
      if (isBranchChange) {
        const staffSel = document.getElementById('f-n'); staffSel.innerHTML = '<option value="">Semua Staff</option>';
        const avNames = [...new Set(db.achv.filter(r => c === "" || r[1] === c).map(r => r[2]))].sort();
        avNames.forEach(st => staffSel.add(new Option(st, st)));
      }
      const getLatest = (src) => {
        let u = {}; src.forEach(r => { if(r[0] >= s && r[0] <= e && (c==="" || r[1]===c) && (n==="" || r[2]===n)) { let k = r[1] + r[2] + (r[3] || ""); if(!u[k] || r[0] > u[k][0]) u[k] = r; } });
        return Object.values(u);
      };
      const fA = getLatest(db.achv); const fC = getLatest(db.cair);
      let resA = { global:{r:0,t:0}, cs:{r:0,t:0}, spv:{r:0,t:0}, tele:{r:0,t:0} };
      fA.forEach(r => { resA.global.r += r[5]; resA.global.t += r[4]; if(resA[r[3]]) { resA[r[3]].r += r[5]; resA[r[3]].t += r[4]; } });
      document.getElementById('g-r').innerText = "Rp " + resA.global.r.toLocaleString('id-ID');
      document.getElementById('g-t').innerText = "Rp " + resA.global.t.toLocaleString('id-ID');
      let pG = resA.global.t > 0 ? (resA.global.r / resA.global.t * 100).toFixed(1) : 0;
      document.getElementById('g-p').innerText = pG + "%";
      ['cs','spv','tele'].forEach(k => {
        let p = resA[k].t > 0 ? (resA[k].r / resA[k].t * 100).toFixed(1) : 0;
        document.getElementById(k+'-p').innerText = p + "%";
        document.getElementById(k+'-tgt').innerText = "Rp " + resA[k].t.toLocaleString('id-ID');
        document.getElementById(k+'-ach').innerText = "Rp " + resA[k].r.toLocaleString('id-ID');
      });
      let resC = {}, gI = 0, gV = 0; fC.forEach(r => { if(!resC[r[3]]) resC[r[3]] = { q:0, v:0 }; resC[r[3]].q += r[4]; resC[r[3]].v += r[5]; gI += r[4]; gV += r[5]; });
      let sel = resA.global.r - gV; let h = ""; Object.keys(resC).sort().forEach(p => { h += `<div class="cair-row"><div><b>${p}</b><br><small>${resC[p].q} Item</small></div><div class="text-end"><span class="text-primary fw-bold">Rp ${resC[p].v.toLocaleString('id-ID')}</span></div></div>`; });
      if (sel > 1000) h += `<div class="cair-row" style="background:#f1f5f9; border-left-color:#cbd5e1;"><div><b class="text-muted">UNMAPPED</b></div><div class="text-end"><span class="text-muted fw-bold">Rp ${sel.toLocaleString('id-ID')}</span></div></div>`;
      h += `<div class="cair-row total-row"><div><b>GRAND TOTAL</b><br><small>${gI} Item</small></div><div class="text-end"><span>Rp ${resA.global.r.toLocaleString('id-ID')}</span></div></div>`;
      document.getElementById('cair-list').innerHTML = h || "<p class='text-center py-5 small text-muted'>Kosong</p>";
      const ctx = document.getElementById('gChart').getContext('2d');
      if(myChart) myChart.destroy();
      myChart = new Chart(ctx, { type:'doughnut', data:{ datasets:[{ data:[pG, 100-pG], backgroundColor:['#00A8E8','rgba(255,255,255,0.1)'], borderWidth:0, cutout:'80%', circumference:240, rotation:240 }] }, options:{ plugins:{ tooltip:{enabled:false} } } });
    }

    function syncPipelineFilter(trigger) {
      const cabEl = document.getElementById('p-c'), spvEl = document.getElementById('p-spv'), salesEl = document.getElementById('p-s');
      if(trigger === 'cab') {
        const v = cabEl.value; const avSPV = [...new Set(db.pipeline.filter(r => v === "" || r.cab === v).map(r => r.spv))].sort();
        const avSales = [...new Set(db.pipeline.filter(r => v === "" || r.cab === v).map(r => r.sal))].sort();
        updateSelect('p-spv', avSPV, "Semua SPV"); updateSelect('p-s', avSales, "Semua Sales");
      } else if(trigger === 'spv') {
        const v = spvEl.value; if(v !== "") { const row = db.pipeline.find(r => r.spv === v); if(row) cabEl.value = row.cab; }
        const avSales = [...new Set(db.pipeline.filter(r => v === "" || r.spv === v).map(r => r.sal))].sort();
        updateSelect('p-s', avSales, "Semua Sales");
      } else if(trigger === 'sales') {
        const v = salesEl.value; if(v !== "") { const row = db.pipeline.find(r => r.sal === v); if(row) { cabEl.value = row.cab; updateSelect('p-spv', [row.spv], "Semua SPV"); spvEl.value = row.spv; } }
      }
      filterPipe();
    }

    function updateSelect(id, list, txt) {
      const s = document.getElementById(id); const oldVal = s.value;
      s.innerHTML = `<option value="">${txt}</option>`;
      list.forEach(v => s.add(new Option(v, v)));
      if(list.includes(oldVal)) s.value = oldVal;
    }

    function filterPipe() {
      const start = document.getElementById('p-start').value || "1900-01-01", end = document.getElementById('p-end').value || "2100-12-31";
      const cab = document.getElementById('p-c').value, spv = document.getElementById('p-spv').value, sal = document.getElementById('p-s').value;
      const f = db.pipeline.filter(r => r.tgl >= start && r.tgl <= end && (cab==="" || r.cab===cab) && (sal==="" || r.sal===sal) && (spv==="" || r.spv===spv));
      
      let m = { apk:0, app:0, cair:0, pla:0 };
      f.forEach(r => {
        m.apk++; m.pla += r.pla;
        if(r.kep.toUpperCase().includes('APPROVE')) m.app++;
        if(r.st.toUpperCase() === 'CAIR') m.cair++;
      });
      document.getElementById('kp-apk').innerText = m.apk; document.getElementById('kp-pla').innerText = "Rp " + m.pla.toLocaleString('id-ID');
      document.getElementById('fn-naik').innerText = m.apk; document.getElementById('fn-app').innerText = m.app; document.getElementById('fn-cair').innerText = m.cair;
      
      let h = ""; f.slice(0,50).forEach(r => {
        h += `<tr><td><b>${r.deb}</b><br><small class="text-muted">${r.sal}</small></td><td><small>${r.cab}<br>${r.spv}</small></td><td>${r.st}</td><td class="fw-bold">Rp ${r.pla.toLocaleString('id-ID')}</td></tr>`;
      });
      document.getElementById('pipe-body').innerHTML = h || "<tr><td colspan='4' class='text-center'>Kosong</td></tr>";
    }

    function filterTablePipe() {
      const q = document.getElementById('p-search').value.toLowerCase();
      const f = db.pipeline.filter(r => r.deb.toLowerCase().includes(q) || r.sal.toLowerCase().includes(q));
      let h = ""; f.slice(0,50).forEach(r => {
        h += `<tr><td><b>${r.deb}</b><br><small class="text-muted">${r.sal}</small></td><td><small>${r.cab}<br>${r.spv}</small></td><td>${r.st}</td><td class="fw-bold">Rp ${r.pla.toLocaleString('id-ID')}</td></tr>`;
      });
      document.getElementById('pipe-body').innerHTML = h || "<tr><td colspan='4' class='text-center'>Kosong</td></tr>";
    }

    function execU() {
      const f = document.getElementById('u-f').files[0], t = document.getElementById('u-t').value, j = document.getElementById('u-j').value;
      if(!f || !t) return Swal.fire('Error', 'Lengkapi data!', 'error');
      Swal.fire({ title: 'Mengunggah...', allowOutsideClick: false, didOpen: () => { Swal.showLoading(); } });
      const r = new FileReader();
      r.onload = function(e) {
        const rows = XLSX.utils.sheet_to_json(XLSX.read(new Uint8Array(e.target.result), {type:'array'}).Sheets[XLSX.read(new Uint8Array(e.target.result), {type:'array'}).SheetNames[0]], {header:1});
        google.script.run.withSuccessHandler(res => { Swal.fire('Berhasil', res, 'success').then(() => refreshData()); }).processExcelData(rows, t, j);
      };
      r.readAsArrayBuffer(f);
    }
  </script>
</body>
</html>
