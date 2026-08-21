/** Power BI–equivalent business logic (client-side filtering). */
const Logic = (() => {
  const day = 86400000;
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  const iso = (d) => {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const dayNum = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${dayNum}`;
  };
  const esc = (v) =>
    String(v ?? "").replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
    );

  const dt = (v) => (v ? new Date(v + "T00:00:00") : null);
  const eta = (r) => r["ETA To Port of Discharge"] || "";
  const atd = (r) => r["Vessel Departed"] || "";

  const uniq = (a) => new Set(a.filter(Boolean)).size;

  function diff(a, b) {
    return a && b ? Math.round((dt(a) - dt(b)) / day) : null;
  }

  function collapse(rows) {
    const m = new Map();
    rows.forEach((r) => {
      const k = r["Container Number"];
      if (!k) return;
      if (!m.has(k)) {
        m.set(k, { ...r });
      } else {
        const merged = m.get(k);
        Object.entries(r).forEach(([f, v]) => {
          if (v && !merged[f]) merged[f] = v;
          else if (v && /Date|ETA|ETD|Arrived|Delivered|Departed/.test(f) && v > merged[f])
            merged[f] = v;
        });
      }
    });
    return [...m.values()];
  }

  function anchor(r) {
    return r["Vessel Arrived"] || eta(r) || r["Container Delivered"] || atd(r) || r["Estimated Time of Departure"] || "";
  }

  function base(all, days) {
    const start = new Date(today);
    start.setDate(start.getDate() - days);
    return all.filter((r) => !anchor(r) || dt(anchor(r)) >= start);
  }

  function isDelivered(r) {
    return !!(r["Vessel Arrived"] || r["Container Delivered"]);
  }

  function plantGroup(loc) {
    const s = (loc || "").toUpperCase();
    if (s === "VTCW" || s.startsWith("EL PASO")) return "VTCW / EL PASO";
    return loc || "";
  }

  function isOceanMot(r) {
    const mot = (r.MOT || "").toUpperCase().trim();
    if (!mot) return true;
    return mot !== "AIR";
  }

  function counts(data, field, n = 8) {
    const c = {};
    data.forEach((r) => {
      const v = r[field];
      if (v) c[v] = (c[v] || 0) + 1;
    });
    return Object.entries(c)
      .sort((a, b) => b[1] - a[1])
      .slice(0, n);
  }

  function state(r) {
    if (r["Vessel Arrived"] || r["Container Delivered"]) return "Delivered";
    const e = r["ETA To Port of Discharge"];
    if (e && dt(e) < today) return "Delayed";
    return "In Transit";
  }

  function bucketFromDays(days) {
    if (days === null) return "";
    if (days === 0) return "On Time";
    const weeks = Math.ceil(Math.abs(days) / 7);
    const bucket = weeks === 1 ? "1 week" : weeks === 2 ? "2 weeks" : "2+ weeks";
    return days < 0 ? "Early by " + bucket : "Delayed by " + bucket;
  }

  function detailFromDays(days) {
    if (days === null) return "";
    if (days === 0) return "On Time";
    const abs = Math.abs(days);
    const n = abs >= 21 ? "21+" : String(abs);
    const unit = n === "1" ? "day" : "days";
    return (days < 0 ? "Early by " : "Delayed by ") + n + " " + unit;
  }

  /**
   * Delivered: ETA To Port of Discharge vs Vessel Arrived
   *   (actual arrival minus planned ETA; positive = delayed)
   * Arriving: Dlv Date vs ETA To Port of Discharge
   *   If Dlv Date is missing: Booked ETA Port vs ETA To Port of Discharge
   *   (current ETA minus target; positive = delayed)
   */
  function etaDays(r) {
    const etaPort = r["ETA To Port of Discharge"];
    if (!etaPort) return null;
    if (r["Vessel Arrived"]) {
      return diff(r["Vessel Arrived"], etaPort);
    }
    const target = r["Dlv Date"] || r["Booked ETA Port"];
    if (!target) return null;
    return diff(etaPort, target);
  }

  function etaPerformance(r) {
    return bucketFromDays(etaDays(r));
  }

  function etaPerformanceDays(r) {
    return detailFromDays(etaDays(r));
  }

  const ETA_COL_WEEK = "ETA Performance";
  const ETA_COL_DAYS = "Detailed ETA Performance Days";

  function withEtaColumns(r) {
    return {
      ...r,
      [ETA_COL_WEEK]: etaPerformance(r),
      [ETA_COL_DAYS]: etaPerformanceDays(r),
    };
  }

  function detailColumns(keys) {
    const hide = new Set(["Current Status", ETA_COL_WEEK, ETA_COL_DAYS]);
    const extra = [ETA_COL_WEEK, ETA_COL_DAYS];
    const base = (keys || []).filter((k) => !hide.has(k));
    const i = base.indexOf("Vessel Arrived");
    if (i < 0) return [...base, ...extra];
    return [...base.slice(0, i + 1), ...extra, ...base.slice(i + 1)];
  }

  function perf(r, type) {
    if (type === "arr") return etaPerformance(r) || "In Transit";
    const plan = r["Estimated Time of Departure"];
    const actual = type === "arr" ? r["Vessel Arrived"] || r["Container Delivered"] : atd(r);
    const d = diff(actual, plan);
    if (d === null) return "In Transit";
    if (d > 10) return "Delayed 10+ days";
    if (d > 5) return "Delayed 5–10 days";
    if (d > 0) return "Delayed by 5 days";
    if (d < -10) return "Early 10+ days";
    if (d < -5) return "Early by 10 days";
    if (d < 0) return "Early by 5 days";
    return "On Time";
  }

  function avgPort(r) {
    return diff(r["Vessel Arrived"], atd(r));
  }

  const ETA_CHIP_VALUES = [
    "Arriving but Delayed by 1 week",
    "Arriving but Delayed by 2 weeks",
    "Arriving but Delayed by 2+ weeks",
    "Arriving but Early by 1 week",
    "Arriving but Early by 2 weeks",
    "Arriving but Early by 2+ weeks",
    "Arriving On Time",
    "Delivered but Delayed by 1 week",
    "Delivered but Delayed by 2 weeks",
    "Delivered but Delayed by 2+ weeks",
    "Delivered but Early by 1 week",
    "Delivered but Early by 2 weeks",
    "Delivered but Early by 2+ weeks",
    "Delivered On Time",
    "Arriving Today",
  ];

  function matchesEtaFilter(r, etaFilter) {
    if (!etaFilter) return true;
    if (etaFilter === "Arriving Today") {
      return !isDelivered(r) && eta(r) === iso(today);
    }
    const delivered = !!r["Vessel Arrived"];
    const prefix = delivered ? "Delivered" : "Arriving";
    const label = etaPerformance(r);
    return etaFilter === (label === "On Time" ? prefix + " On Time" : prefix + " but " + label);
  }

  const VARIANCE_ORDER = [
    "Delayed by 2+ weeks",
    "Delayed by 2 weeks",
    "Delayed by 1 week",
    "On Time",
    "Early by 1 week",
    "Early by 2 weeks",
    "Early by 2+ weeks",
  ];

  const ARRIVAL_PERF_ORDER = [...VARIANCE_ORDER, "In Transit"];

  const DEPARTURE_ORDER = [
    "Delayed 10+ days",
    "Delayed 5–10 days",
    "Delayed by 5 days",
    "On Time",
    "Early by 5 days",
    "Early by 10 days",
    "Early 10+ days",
    "In Transit",
  ];

  function sortByOrder(entries, order) {
    const rank = new Map(order.map((k, i) => [k, i]));
    return [...entries].sort((a, b) => {
      const ra = rank.has(a[0]) ? rank.get(a[0]) : order.length;
      const rb = rank.has(b[0]) ? rank.get(b[0]) : order.length;
      return ra - rb || b[1] - a[1];
    });
  }

  function perfColor(label) {
    if (label.startsWith("Delayed")) return "#e95a5e";
    if (label.startsWith("Early")) return "#20b87a";
    if (label === "On Time") return "#f07032";
    return "#2698e8";
  }

  const FILTER_SPECS = [
    ["Plant", "Delivery Location"],
    ["Carrier", "Steamship Line"],
    ["Port of Loading", "Port of Loading"],
    ["Port of Discharge", "Port Of Discharge"],
    ["Arrived/Delivered", "Arrived/Delivered"],
    ["Container", "Container Number"],
    ["PO #", "PO#"],
    ["Hotlist", "Hotlist"],
    ["Class", "Class"],
    ["Forwarder", "Forwarder"],
  ];

  function classTokens(value) {
    return [...new Set(String(value || "").split(/[,;]/).map((t) => t.trim()).filter(Boolean))];
  }

  function fieldMatches(row, field, selected) {
    if (!selected) return true;
    if (field === "Class") return classTokens(row.Class).includes(selected);
    return row[field] === selected;
  }

  return {
    day,
    today,
    iso,
    esc,
    dt,
    eta,
    atd,
    uniq,
    diff,
    collapse,
    anchor,
    base,
    isDelivered,
    plantGroup,
    isOceanMot,
    counts,
    state,
    etaDays,
    etaPerformance,
    etaPerformanceDays,
    withEtaColumns,
    detailColumns,
    perf,
    avgPort,
    ETA_CHIP_VALUES,
    VARIANCE_ORDER,
    ARRIVAL_PERF_ORDER,
    DEPARTURE_ORDER,
    sortByOrder,
    perfColor,
    matchesEtaFilter,
    FILTER_SPECS,
    classTokens,
    fieldMatches,
  };
})();
