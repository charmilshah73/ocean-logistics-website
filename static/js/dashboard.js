/** Dashboard UI wiring — filters, charts, pages. */
const Dashboard = (() => {
  let all = [];
  let viewA = [];
  let viewP = [];
  let cols = [];
  let etaFilter = "";
  let performanceGroup = "Forwarder";
  let averageGroup = "Steamship Line";

  const $ = (id) => document.getElementById(id);
  const L = Logic;

  function showPage(id) {
    document.querySelectorAll(".page").forEach((x) => x.classList.toggle("active", x.id === id));
    scrollTo(0, 0);
    if (id === "performance") applyP();
    else if (id === "analysis" || id === "detail") applyA();
  }

  function showDetail() {
    showPage("detail");
  }

  function refreshWorkbook() {
    location.reload();
  }

  function filtersHTML(prefix) {
    return (
      L.FILTER_SPECS.map(
        ([lab, f]) =>
          `<div class="filter"><label>${lab}</label><select data-field="${f}" data-prefix="${prefix}"><option value="">All</option></select></div>`
      ).join("") + `<button class="reset-btn" data-reset="${prefix}">Reset All Filters</button>`
    );
  }

  function setupFilters(id, prefix) {
    $(id).innerHTML = filtersHTML(prefix);
    document.querySelectorAll(`#${id} [data-prefix="${prefix}"]`).forEach((s) => {
      s.onchange = prefix === "P" ? applyP : applyA;
    });
    document.querySelector(`#${id} [data-reset="${prefix}"]`).onclick = () => {
      document.querySelectorAll(`#filtersA [data-prefix="A"], #filtersD [data-prefix="A"], #filtersP [data-prefix="P"]`)
        .forEach((s) => {
          if (s.dataset.prefix === prefix) s.value = "";
        });
      if (prefix === "A") {
        etaFilter = "";
        document.querySelectorAll(".chip").forEach((b) => b.classList.remove("on"));
      }
      prefix === "P" ? applyP() : applyA();
    };
  }

  function windowRows(days) {
    return L.base(all.filter(L.isOceanMot), days);
  }

  function refreshFilterOptions(prefix, days) {
    const barId = prefix === "P" ? "filtersP" : "filtersA";
    const bar = $(barId);
    if (!bar) return;
    const selects = [...bar.querySelectorAll("select")];
    const windowed = windowRows(days);
    selects.forEach((s) => {
      const field = s.dataset.field;
      const rows = windowed.filter((r) =>
        selects.every((other) => other === s || L.fieldMatches(r, other.dataset.field, other.value))
      );
      let vals =
        field === "Class"
          ? [...new Set(rows.flatMap((r) => L.classTokens(r.Class)))].sort((a, b) =>
              a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" })
            )
          : [...new Set(rows.map((r) => r[field]).filter(Boolean))].sort();
      if (field === "Forwarder") vals = vals.filter((v) => v !== "Detailed Tracking");
      const keep = vals.includes(s.value) ? s.value : "";
      s.innerHTML =
        `<option value=""${keep === "" ? " selected" : ""}>All</option>` +
        vals
          .map((v) => `<option${v === keep ? " selected" : ""}>${L.esc(v)}</option>`)
          .join("");
      s.value = keep;
    });
  }

  function selected(prefix, r) {
    const barId = prefix === "P" ? "filtersP" : "filtersA";
    return [...$(barId).querySelectorAll("select")].every((s) =>
      L.fieldMatches(r, s.dataset.field, s.value)
    );
  }

  function bars(id, items) {
    const m = items[0]?.[1] || 1;
    $(id).innerHTML = items.length
      ? items
          .map(
            ([n, v]) =>
              `<div class="bar"><span class="name" title="${L.esc(n)}">${L.esc(n)}</span><div class="track"><div class="fill" style="width:${(100 * v) / m}%"></div></div><span class="val">${v}</span></div>`
          )
          .join("")
      : '<div class="empty">No data</div>';
  }

  function columns(id, items, line = false) {
    const m = Math.max(...items.map((x) => x[1]), 1);
    const total = items.reduce((a, x) => a + x[1], 0) || 1;
    const el = $(id);
    if (!items.length) {
      el.classList.remove("line-mode");
      el.innerHTML = '<div class="empty">No data</div>';
      return;
    }
    el.classList.toggle("line-mode", line);
    el.innerHTML = items
      .map(
        ([n, v, c]) =>
          `<div class="colwrap"><div class="col" style="height:${(100 * v) / m}%;background:${c || ""}">${
            line ? `<span class="col-count">${v}</span>` : `<em>${v}</em>`
          }<span title="${L.esc(n)}">${L.esc(n)}</span></div></div>`
      )
      .join("");
    if (!line) return;
    requestAnimationFrame(() => {
      el.querySelectorAll(".pct-overlay").forEach((node) => node.remove());
      const box = el.getBoundingClientRect();
      if (!box.width || !box.height) return;
      const points = [...el.querySelectorAll(".col")].map((col, i) => {
        const r = col.getBoundingClientRect();
        const exact = ((items[i][1] / total) * 100).toFixed(1);
        return {
          x: r.left + r.width / 2 - box.left,
          y: r.top - box.top,
          pct: Math.round((items[i][1] / total) * 100),
          exact: `${exact}%`,
        };
      });
      const w = Math.round(box.width);
      const h = Math.round(box.height);
      const overlay = document.createElement("div");
      overlay.className = "pct-overlay";
      overlay.innerHTML = `<svg width="${w}" height="${h}" xmlns="http://www.w3.org/2000/svg">
        <polyline points="${points.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(" ")}"/>
        ${points
          .map(
            (p) =>
              `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="3.5" fill="#fff" stroke="#e84aa5" stroke-width="2"/>`
          )
          .join("")}
      </svg>${points
        .map(
          (p) =>
            `<span class="pct-hit" title="${p.exact}" style="left:${p.x.toFixed(1)}px;top:${Math.max(2, p.y - 18).toFixed(1)}px">${p.pct}%</span>`
        )
        .join("")}`;
      el.appendChild(overlay);
    });
  }

  function perfEntries(rows, type) {
    const counts = {};
    rows.forEach((r) => {
      const label = L.perf(r, type);
      counts[label] = (counts[label] || 0) + 1;
    });
    const order = type === "dep" ? L.DEPARTURE_ORDER : L.ARRIVAL_PERF_ORDER;
    return L.sortByOrder(Object.entries(counts), order).map(([name, value]) => [
      name,
      value,
      L.perfColor(name),
    ]);
  }

  function setEtaFilter(value) {
    etaFilter = etaFilter === value ? "" : value;
    document.querySelectorAll("#etaButtons .chip,#etaButtonsD .chip").forEach((b) =>
      b.classList.toggle("on", b.dataset.value === etaFilter)
    );
    applyA();
  }

  function buildEtaButtons() {
    const buttons = L.ETA_CHIP_VALUES.map(
      (v) => `<button type="button" class="chip" data-value="${v}">${v}</button>`
    ).join("");
    $("etaButtons").innerHTML = buttons;
    $("etaButtonsD").innerHTML = buttons;
    document.querySelectorAll("#etaButtons .chip,#etaButtonsD .chip").forEach((btn) => {
      btn.onclick = () => setEtaFilter(btn.dataset.value);
    });
  }

  function applyA() {
    refreshFilterOptions("A", 14);
    const population = L.collapse(
      L.base(all.filter(L.isOceanMot), 14).filter((r) => selected("A", r))
    );
    viewA = population.filter((r) => L.matchesEtaFilter(r, etaFilter));
    renderA();
    renderTable();
    syncDetailFilters();
  }

  function renderA() {
    const e7 = new Date(L.today);
    e7.setDate(e7.getDate() + 7);
    const e14 = new Date(L.today);
    e14.setDate(e14.getDate() + 14);

    const e8 = new Date(L.today);
    e8.setDate(e8.getDate() + 8);
    const containerCount = L.uniq(viewA.map((r) => r["Container Number"]));
    $("kTotal").textContent = containerCount;
    $("k7").textContent = L.uniq(
      viewA
        .filter((r) => !L.isDelivered(r) && L.eta(r) >= L.iso(L.today) && L.eta(r) <= L.iso(e7))
        .map((r) => r["Container Number"])
    );
    $("k14").textContent = L.uniq(
      viewA
        .filter((r) => !L.isDelivered(r) && L.eta(r) >= L.iso(e8) && L.eta(r) <= L.iso(e14))
        .map((r) => r["Container Number"])
    );
    $("kDelay").textContent = L.uniq(
      viewA.filter((r) => L.etaPerformance(r).startsWith("Delayed")).map((r) => r["Container Number"])
    );
    $("kToday").textContent = L.uniq(
      viewA.filter((r) => !L.isDelivered(r) && L.eta(r) === L.iso(L.today)).map((r) => r["Container Number"])
    );
    $("kHot").textContent = L.uniq(
      viewA.filter((r) => (r.Hotlist || "").toLowerCase() === "yes").map((r) => r["Container Number"])
    );

    const parts = [
      ["Delayed", viewA.filter((r) => L.etaPerformance(r).startsWith("Delayed")).length, "#e95a5e"],
      ["Early", viewA.filter((r) => L.etaPerformance(r).startsWith("Early")).length, "#20b87a"],
      ["On Time", viewA.filter((r) => L.etaPerformance(r) === "On Time").length, "#2698e8"],
    ].filter((x) => x[1]);

    const total = parts.reduce((s, x) => s + x[1], 0) || 1;
    let a = 0;
    const st = [];
    parts.forEach((x) => {
      const b = a + (x[1] / total) * 360;
      st.push(`${x[2]} ${a}deg ${b}deg`);
      a = b;
    });
    $("etaDonut").style.background = `conic-gradient(${st.join(",") || "#ddd 0 360deg"})`;
    $("etaCenter").innerHTML = `${containerCount}<small>container #</small>`;
    $("etaLegend").innerHTML = parts
      .map((x) => {
        const exact = ((x[1] / total) * 100).toFixed(1);
        return `<div class="leg"><i class="dot" style="background:${x[2]}"></i><span>${x[0]}</span><b title="${exact}%">${Math.round((x[1] / total) * 100)}%</b></div>`;
      })
      .join("");

    bars(
      "plantBars",
      L.counts(
        viewA.map((r) => ({ ...r, "Plant Group": L.plantGroup(r["Delivery Location"]) })),
        "Plant Group"
      )
    );
    bars("vendorBars", L.counts(viewA, "Shipper Name", 10));

    const vc = {};
    viewA.forEach((r) => {
      const v = L.etaPerformance(r);
      if (v) vc[v] = (vc[v] || 0) + 1;
    });
    columns(
      "varianceCols",
      L.VARIANCE_ORDER.filter((label) => vc[label]).map((label) => [
        label,
        vc[label],
        label.startsWith("Delayed") ? "#e95a5e" : label.startsWith("Early") ? "#20b87a" : "#2698e8",
      ])
    );
  }

  function getDetailExportRows() {
    const q = ($("tableSearch")?.value || "").toLowerCase();
    return viewA
      .map((r) => L.withEtaColumns(r))
      .filter((r) => !q || Object.values(r).join(" ").toLowerCase().includes(q));
  }

  async function downloadDetailExcel() {
    const rows = getDetailExportRows();
    if (!rows.length) {
      alert("No rows to export. Clear filters or search and try again.");
      return;
    }
    const btn = $("downloadDetailBtn");
    if (btn) {
      btn.disabled = true;
      btn.textContent = "Preparing…";
    }
    try {
      const res = await fetch("/api/export/detail", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ records: rows, columns: cols }),
      });
      if (!res.ok) {
        let message = `Export failed (${res.status})`;
        if (res.status === 404) {
          message = "Export API not found. Restart the server: python run.py";
        } else {
          const err = await res.json().catch(() => ({}));
          if (err.error) message = err.error;
        }
        throw new Error(message);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const stamp = L.iso(new Date());
      a.href = url;
      a.download = `Ocean_Detailed_Report_${stamp}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      alert(String(e.message || e));
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.textContent = "↓ Download Excel";
      }
    }
  }

  function renderTable() {
    if (!$("tableHead")) return;
    const q = ($("tableSearch")?.value || "").toLowerCase();
    const filtered = getDetailExportRows();
    const rows = filtered.slice(0, 600);
    $("rowCount").textContent = `showing ${rows.length} of ${filtered.length} (filtered)`;
    $("tableHead").innerHTML = cols.map((c) => `<th>${L.esc(c)}</th>`).join("");
    $("tableBody").innerHTML =
      rows
        .map(
          (r) =>
            `<tr>${cols.map((c) => `<td><span class="${c === "Hotlist" && r[c] === "Yes" ? "pill hot" : "pill"}">${L.esc(r[c] || "—")}</span></td>`).join("")}</tr>`
        )
        .join("") || "<tr><td colspan='99'>No matching shipments.</td></tr>";
  }

  function setPerformanceGroup(field) {
    performanceGroup = field;
    $("byForwarder").classList.toggle("on", field === "Forwarder");
    $("byCarrier").classList.toggle("on", field === "Steamship Line");
    renderP();
  }

  function setAverageGroup(field) {
    averageGroup = field;
    $("avgByCarrier").classList.toggle("on", field === "Steamship Line");
    $("avgByForwarder").classList.toggle("on", field === "Forwarder");
    renderP();
  }

  function applyP() {
    refreshFilterOptions("P", 365);
    viewP = L.collapse(L.base(all.filter(L.isOceanMot), 365).filter((r) => selected("P", r)));
    renderP();
  }

  function renderP() {
    $("pTotal").textContent = L.uniq(viewP.map((r) => r["Container Number"]));
    $("pDelivered").textContent = viewP.filter((r) => L.state(r) === "Delivered").length;
    $("pTransit").textContent = viewP.filter((r) => L.state(r) === "In Transit").length;
    $("pDelayed").textContent = viewP.filter((r) => L.perf(r, "arr").startsWith("Delayed")).length;
    $("pOnTime").textContent = viewP.filter((r) => L.perf(r, "arr") === "On Time").length;

    const av = viewP.map(L.avgPort).filter((x) => x !== null && x >= 0 && x < 200);
    $("pAvg").textContent = av.length ? Math.round(av.reduce((a, b) => a + b, 0) / av.length) + " d" : "—";

    columns("typeCols", L.counts(viewP, "CTR SIZE / LCL", 10));

    columns("arrivalCols", perfEntries(viewP, "arr"), true);

    const years = {};
    viewP.forEach((r) => {
      const etaVal = L.eta(r);
      if (!etaVal || etaVal.length < 4) return;
      const y = etaVal.slice(0, 4);
      if (!/^\d{4}$/.test(y)) return;
      const m = (r["CTR SIZE / LCL"] || "").toUpperCase().includes("LCL") ? "LCL" : "FCL";
      years[y + " " + m] = (years[y + " " + m] || 0) + 1;
    });
    columns(
      "modeCols",
      Object.entries(years)
        .sort((a, b) => a[0].localeCompare(b[0]))
        .slice(-8)
    );

    const groupRows =
      performanceGroup === "Forwarder" ? viewP.filter((r) => r.Forwarder !== "Detailed Tracking") : viewP;
    const forwards = L.counts(groupRows, performanceGroup, 18).map((x) => x[0]);
    $("forwardStacks").innerHTML = forwards
      .map((f) => {
        const rows = groupRows.filter((r) => r[performanceGroup] === f);
        const p = [
          ["#e95a5e", rows.filter((r) => L.perf(r, "arr").startsWith("Delayed")).length],
          ["#20b87a", rows.filter((r) => L.perf(r, "arr").startsWith("Early")).length],
          ["#8178cb", rows.filter((r) => L.perf(r, "arr") === "On Time" || L.perf(r, "arr") === "In Transit").length],
        ];
        const t = p.reduce((s, x) => s + x[1], 0) || 1;
        return `<div class="stackbar">${p
          .map(([color, count]) => {
            if (!count) return "";
            const show = count / t >= 0.08;
            const exact = ((count / t) * 100).toFixed(1);
            return `<div class="seg" style="height:${(count / t) * 100}%;background:${color}">${
              show
                ? `<span class="seg-label">${count}<small title="${exact}%">${Math.round((count / t) * 100)}%</small></span>`
                : ""
            }</div>`;
          })
          .join("")}<label title="${L.esc(f)}">${L.esc(f)}</label></div>`;
      })
      .join("");

    columns("departureCols", perfEntries(viewP, "dep"), true);

    const carriers = {};
    viewP.forEach((r) => {
      const c = r[averageGroup];
      const d = L.avgPort(r);
      if (averageGroup === "Forwarder" && c === "Detailed Tracking") return;
      if (c && d !== null && d >= 0 && d < 200) (carriers[c] ??= []).push(d);
    });
    bars(
      "carrierBars",
      Object.entries(carriers)
        .map(([c, a]) => [c, Math.round(a.reduce((x, y) => x + y, 0) / a.length)])
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
    );
  }

  function syncDetailFilters() {
    const source = $("filtersA");
    const dest = $("filtersD");
    if (!source || !dest) return;
    const selected = [...source.querySelectorAll("select")].map((s) => [s.dataset.field, s.value]);
    dest.innerHTML = source.innerHTML;
    selected.forEach(([field, value]) => {
      const box = dest.querySelector(`select[data-field="${CSS.escape(field)}"]`);
      if (box) box.value = value;
    });
    dest.querySelectorAll('[data-prefix="A"]').forEach((s) => {
      s.onchange = () => {
        const src = source.querySelector(`select[data-field="${CSS.escape(s.dataset.field)}"]`);
        if (src) src.value = s.value;
        applyA();
      };
    });
    const reset = dest.querySelector('[data-reset="A"]');
    if (reset) {
      reset.onclick = () => {
        source.querySelectorAll("select").forEach((s) => (s.value = ""));
        etaFilter = "";
        document.querySelectorAll(".chip").forEach((b) => b.classList.remove("on"));
        applyA();
      };
    }
  }

  function updateStamps(modified, rowCount) {
    const stamp = `${rowCount.toLocaleString()} shipment rows · refreshed ${modified}`;
    $("landStamp").textContent = stamp;
    if ($("refreshStamp")) $("refreshStamp").textContent = modified;
    if ($("refreshRows")) $("refreshRows").textContent = rowCount.toLocaleString() + " shipment rows";
  }

  async function fetchJson(url) {
    try {
      const res = await fetch(url, { cache: "no-store" });
      if (!res.ok) return null;
      const data = await res.json();
      if (data && data.error) return null;
      return data;
    } catch {
      return null;
    }
  }

  async function init() {
    try {
      let payload = await fetchJson("/api/data");
      let status = (await fetchJson("/api/status")) || {};
      if (!payload) {
        payload = await fetchJson("data/dashboard-data.json");
        status = {};
      }
      if (!payload) throw new Error("Unable to load dashboard data");

      all = payload.records || [];
      cols = L.detailColumns(all.length ? Object.keys(all[0]) : []);

      updateStamps(payload.modified || status.modified || "—", payload.rowCount || all.length);
      if ($("dashboardDate")) {
        $("dashboardDate").textContent = L.iso(L.today).replace(/-/g, "/");
      }

      setupFilters("filtersA", "A");
      buildEtaButtons();
      syncDetailFilters();
      setupFilters("filtersP", "P");
      $("tableSearch").oninput = renderTable;
      applyA();
      applyP();
    } catch (e) {
      $("landStamp").textContent = "Unable to load data";
      document.querySelector("#landing .land-hero").innerHTML = `
        <div class="error-panel">
          <h2>Dashboard could not start</h2>
          <p>${L.esc(String(e))}</p>
          <p>Run the server from <code>logistics_website</code> or upload files at <a href="/admin">/admin</a>.</p>
        </div>`;
    }
  }

  return {
    showPage,
    showDetail,
    refreshWorkbook,
    setPerformanceGroup,
    setAverageGroup,
    downloadDetailExcel,
    init,
  };
})();

window.showPage = Dashboard.showPage;
window.showDetail = Dashboard.showDetail;
window.refreshWorkbook = Dashboard.refreshWorkbook;
window.setPerformanceGroup = Dashboard.setPerformanceGroup;
window.setAverageGroup = Dashboard.setAverageGroup;
window.downloadDetailExcel = Dashboard.downloadDetailExcel;

document.addEventListener("DOMContentLoaded", () => Dashboard.init());
