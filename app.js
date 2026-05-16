"use strict";

const CONFIG = {
  apiKey: "AIzaSyCjceOafhRMvv2A5nAp0zzJrf8f40lpQrM",
  spreadsheetId: "1wrqKSbH21jvwmesMdCVymsFSQD-P1OHU4iINq1uYdWo",
  refreshInterval: 10 * 60 * 1000,
  sheets: { farmers: "Farmers", payments: "Farmers Payments", delivery: "Farmers Delivery" }
};

const AppState = {
  raw: { farmers: [], payments: [], delivery: [] },
  filtered: { farmers: [], payments: [], delivery: [] },
  filters: { dateFrom: "", dateTo: "", zone: "", village: "", vano: "", package: "", gender: "" },
  charts: new Map(),
  activeTab: "overview",
  lastRefresh: null,
  map: null,
  mapLayers: { clusters: null, heat: null, zones: null },
  mapMode: "clusters"
};

const $ = selector => document.querySelector(selector);
const $$ = selector => Array.from(document.querySelectorAll(selector));

const Utils = {
  num(value) {
    if (value === null || value === undefined || value === "") return 0;
    return parseFloat(String(value).replace(/[^0-9.-]/g, "")) || 0;
  },
  fmtNum(value, digits = 0) {
    if (value === null || value === undefined || Number.isNaN(Number(value))) return "-";
    return Number(value).toLocaleString("en-US", { maximumFractionDigits: digits, minimumFractionDigits: digits });
  },
  fmtPct(value, digits = 1) {
    if (value === null || Number.isNaN(Number(value))) return "-";
    return `${Number(value).toFixed(digits)}%`;
  },
  fmtCurrency(value, symbol = "TZS") {
    const n = Number(value) || 0;
    if (!n) return `${symbol} 0`;
    if (Math.abs(n) >= 1e9) return `${symbol} ${(n / 1e9).toFixed(1)}B`;
    if (Math.abs(n) >= 1e6) return `${symbol} ${(n / 1e6).toFixed(1)}M`;
    if (Math.abs(n) >= 1e3) return `${symbol} ${(n / 1e3).toFixed(1)}K`;
    return `${symbol} ${Utils.fmtNum(n)}`;
  },
  monthKey(value) {
    if (!value) return "Unknown";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "Unknown";
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  },
  monthLabel(key) {
    if (!key || key === "Unknown") return key;
    const [year, month] = key.split("-");
    return new Date(Number(year), Number(month) - 1, 1).toLocaleDateString("en-US", { month: "short", year: "numeric" });
  },
  dateLabel(value) {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? value : date.toLocaleDateString("en-US", { day: "2-digit", month: "short" });
  },
  unique(rows, key) {
    return [...new Set(rows.map(row => row[key]).filter(Boolean))];
  },
  countBy(rows, key) {
    return rows.reduce((acc, row) => {
      const k = row[key] || "Unknown";
      acc[k] = (acc[k] || 0) + 1;
      return acc;
    }, {});
  },
  groupBy(rows, key) {
    return rows.reduce((acc, row) => {
      const k = row[key] || "Unknown";
      (acc[k] ||= []).push(row);
      return acc;
    }, {});
  },
  sumBy(rows, key) {
    return rows.reduce((sum, row) => sum + Utils.num(row[key]), 0);
  },
  avgBy(rows, key) {
    return rows.length ? Utils.sumBy(rows, key) / rows.length : 0;
  },
  entriesSorted(object, limit = 999) {
    return Object.entries(object).filter(([key]) => key && key !== "Unknown").sort((a, b) => b[1] - a[1]).slice(0, limit);
  },
  esc(value) {
    return String(value ?? "").replace(/[&<>"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[char]));
  },
  cssVar(name) {
    return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  }
};

const DataEngine = {
  baseUrl: "https://sheets.googleapis.com/v4/spreadsheets",
  async fetchSheet(sheetName) {
    const url = `${this.baseUrl}/${CONFIG.spreadsheetId}/values/${encodeURIComponent(sheetName)}?key=${CONFIG.apiKey}`;
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Failed to fetch ${sheetName}: ${response.status} ${response.statusText}`);
    const data = await response.json();
    return this.parseSheet(data.values || []);
  },
  parseSheet(values) {
    if (!values.length) return [];
    const headers = values[0].map(header => String(header).trim());
    return values.slice(1).map(row => {
      const record = {};
      headers.forEach((header, index) => { record[header] = row[index] !== undefined ? String(row[index]).trim() : ""; });
      return record;
    }).filter(row => Object.values(row).some(Boolean));
  },
  async fetchAll(onProgress) {
    const entries = Object.entries(CONFIG.sheets);
    const output = {};
    for (let i = 0; i < entries.length; i++) {
      const [key, sheet] = entries[i];
      onProgress?.(15 + Math.round((i / entries.length) * 70), `Loading ${sheet}...`);
      output[key] = await this.fetchSheet(sheet);
    }
    return output;
  }
};

const FilterEngine = {
  populate() {
    const all = [...AppState.raw.farmers, ...AppState.raw.payments, ...AppState.raw.delivery];
    this.fill("f-zone", Utils.unique(all, "Zone").sort(), "All zones");
    this.fill("f-village", Utils.unique(all, "Village").sort(), "All villages");
    this.fill("f-vano", Utils.unique(AppState.raw.farmers, "VANO").sort(), "All VANOs");
    const packages = [...Utils.unique(AppState.raw.farmers, "PACKAGE"), ...Utils.unique(AppState.raw.payments, "Package Name"), ...Utils.unique(AppState.raw.delivery, "Package Delivered")];
    this.fill("f-package", [...new Set(packages.filter(Boolean))].sort(), "All packages");
    const months = [...new Set(AppState.raw.farmers.map(row => Utils.monthKey(row["DATE CREATED"])).filter(key => key !== "Unknown"))].sort();
    this.fill("f-date-from", months, "Any start", Utils.monthLabel);
    this.fill("f-date-to", months, "Any end", Utils.monthLabel);
  },
  fill(id, values, placeholder, labelFn = value => value) {
    const element = document.getElementById(id);
    if (!element) return;
    const current = element.value;
    element.innerHTML = `<option value="">${placeholder}</option>`;
    values.forEach(value => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = labelFn(value);
      option.selected = value === current;
      element.appendChild(option);
    });
  },
  read() {
    AppState.filters = {
      dateFrom: $("#f-date-from")?.value || "",
      dateTo: $("#f-date-to")?.value || "",
      zone: $("#f-zone")?.value || "",
      village: $("#f-village")?.value || "",
      vano: $("#f-vano")?.value || "",
      package: $("#f-package")?.value || "",
      gender: $("#f-gender")?.value || ""
    };
  },
  apply() {
    this.read();
    const f = AppState.filters;
    AppState.filtered.farmers = AppState.raw.farmers.filter(row => {
      if (f.zone && row.Zone !== f.zone) return false;
      if (f.village && row.Village !== f.village) return false;
      if (f.vano && row.VANO !== f.vano) return false;
      if (f.package && row.PACKAGE !== f.package) return false;
      if (f.gender && row.GENDER !== f.gender) return false;
      const month = Utils.monthKey(row["DATE CREATED"]);
      if (f.dateFrom && month < f.dateFrom) return false;
      if (f.dateTo && month > f.dateTo) return false;
      return true;
    });
    AppState.filtered.payments = AppState.raw.payments.filter(row => {
      if (f.zone && row.Zone !== f.zone) return false;
      if (f.village && row.Village !== f.village) return false;
      if (f.vano && row["Village Agent"] !== f.vano) return false;
      if (f.package && row["Package Name"] !== f.package) return false;
      if (f.gender && row.Gender !== f.gender) return false;
      return true;
    });
    AppState.filtered.delivery = AppState.raw.delivery.filter(row => {
      if (f.zone && row.Zone !== f.zone) return false;
      if (f.village && row.Village !== f.village) return false;
      if (f.package && row["Package Delivered"] !== f.package) return false;
      if (f.gender && row.Gender !== f.gender) return false;
      return true;
    });
  },
  clear() {
    ["f-date-from", "f-date-to", "f-zone", "f-village", "f-vano", "f-package", "f-gender"].forEach(id => {
      const element = document.getElementById(id);
      if (element) element.value = "";
    });
    this.apply();
  }
};

const ChartEngine = {
  palette() {
    return ["--green", "--green-2", "--gold", "--blue", "--danger", "--muted"].map(Utils.cssVar);
  },
  base() {
    return {
      animation: !window.matchMedia("(prefers-reduced-motion: reduce)").matches,
      animationDuration: 600,
      animationEasing: "cubicOut",
      textStyle: { fontFamily: "Inter, system-ui, sans-serif", color: Utils.cssVar("--text") },
      color: this.palette(),
      grid: { top: 24, right: 22, bottom: 42, left: 48, containLabel: true },
      tooltip: {
        trigger: "item",
        backgroundColor: Utils.cssVar("--card"),
        borderColor: Utils.cssVar("--border-strong"),
        borderWidth: 1,
        textStyle: { color: Utils.cssVar("--text"), fontSize: 12 },
        extraCssText: "box-shadow:0 18px 45px rgba(15,23,42,.14);border-radius:14px;padding:12px 14px;"
      }
    };
  },
  get(id) {
    const element = document.getElementById(id);
    if (!element || !window.echarts) return null;
    if (!AppState.charts.has(id)) AppState.charts.set(id, echarts.init(element, null, { renderer: "canvas" }));
    return AppState.charts.get(id);
  },
  set(id, option) {
    const chart = this.get(id);
    if (!chart) return;
    chart.setOption(option, true);
  },
  xAxis(data, vertical = false) {
    return { type: "category", data, axisTick: { show: false }, axisLine: { lineStyle: { color: Utils.cssVar("--border-strong") } }, axisLabel: { color: Utils.cssVar("--muted"), fontSize: 11, rotate: !vertical && data.length > 6 ? 24 : 0 } };
  },
  yAxis(formatter) {
    return { type: "value", axisLabel: { color: Utils.cssVar("--muted"), fontSize: 11, formatter }, splitLine: { lineStyle: { color: Utils.cssVar("--border") } } };
  },
  legend() {
    return { bottom: 0, textStyle: { color: Utils.cssVar("--muted"), fontSize: 11 } };
  },
  line(id, categories, series) {
    const palette = this.palette();
    this.set(id, { ...this.base(), tooltip: { ...this.base().tooltip, trigger: "axis" }, xAxis: this.xAxis(categories), yAxis: this.yAxis(), legend: series.length > 1 ? this.legend() : { show: false }, series: series.map((item, index) => ({ type: "line", name: item.name, data: item.data, smooth: true, symbol: "circle", symbolSize: 5, lineStyle: { width: 3, color: item.color || palette[index % palette.length] }, itemStyle: { color: item.color || palette[index % palette.length] }, areaStyle: index === 0 ? { opacity: 0.08 } : undefined })) });
  },
  bar(id, categories, series, options = {}) {
    const palette = this.palette();
    const horizontal = options.horizontal || false;
    this.set(id, { ...this.base(), tooltip: { ...this.base().tooltip, trigger: "axis" }, xAxis: horizontal ? this.yAxis(options.valueFormatter) : this.xAxis(categories), yAxis: horizontal ? this.xAxis(categories, true) : this.yAxis(options.valueFormatter), legend: series.length > 1 ? this.legend() : { show: false }, series: series.map((item, index) => ({ type: "bar", name: item.name, data: item.data, barMaxWidth: 36, stack: item.stack, itemStyle: { color: item.color || palette[index % palette.length], borderRadius: horizontal ? [0, 8, 8, 0] : [8, 8, 0, 0] } })) });
  },
  donut(id, data) {
    this.set(id, { ...this.base(), tooltip: { ...this.base().tooltip, formatter: "{b}<br>{c} ({d}%)" }, legend: { orient: "vertical", right: 0, top: "center", textStyle: { color: Utils.cssVar("--muted"), fontSize: 11 } }, series: [{ type: "pie", radius: ["52%", "74%"], center: ["38%", "50%"], label: { show: false }, labelLine: { show: false }, data }] });
  },
  funnel(id, data) {
    this.set(id, { ...this.base(), tooltip: { ...this.base().tooltip, formatter: "{b}<br>{c}" }, series: [{ type: "funnel", left: "10%", top: 18, bottom: 18, width: "80%", sort: "none", label: { color: "#fff", fontWeight: 800 }, itemStyle: { borderColor: "transparent" }, data }] });
  },
  radar(id, indicators, values) {
    this.set(id, { ...this.base(), radar: { indicator: indicators.map(name => ({ name, max: 100 })), axisName: { color: Utils.cssVar("--muted"), fontSize: 11 }, splitLine: { lineStyle: { color: Utils.cssVar("--border") } }, splitArea: { areaStyle: { color: ["transparent", "rgba(27,94,32,.04)"] } } }, series: [{ type: "radar", data: [{ value: values, areaStyle: { opacity: 0.16 }, lineStyle: { width: 3 }, itemStyle: { color: Utils.cssVar("--green") } }] }] });
  },
  scatter(id, data) {
    this.set(id, { ...this.base(), tooltip: { ...this.base().tooltip, formatter: point => `${point.name}<br>Farmers: ${point.value[0]}<br>Collection: ${point.value[1]}%` }, xAxis: { type: "value", name: "Farmers", axisLabel: { color: Utils.cssVar("--muted") }, splitLine: { lineStyle: { color: Utils.cssVar("--border") } } }, yAxis: { type: "value", name: "Collection %", axisLabel: { color: Utils.cssVar("--muted"), formatter: "{value}%" }, splitLine: { lineStyle: { color: Utils.cssVar("--border") } } }, series: [{ type: "scatter", data, symbolSize: value => Math.max(Math.sqrt(value[2]) * 4, 12), itemStyle: { opacity: 0.86 }, label: { show: true, formatter: point => point.name, position: "right", color: Utils.cssVar("--muted"), fontSize: 10 } }] });
  },
  resizeAll() {
    AppState.charts.forEach(chart => chart.resize());
  },
  rerenderTheme() {
    AppState.charts.forEach(chart => chart.dispose());
    AppState.charts.clear();
    renderActiveTab();
  }
};

const KPIEngine = {
  render(id, cards) {
    const element = document.getElementById(id);
    if (!element) return;
    element.innerHTML = cards.map(card => `
      <article class="metric-card">
        <div><p class="metric-label">${Utils.esc(card.label)}</p><div class="metric-value" data-counter="${card.raw ?? ""}">${Utils.esc(card.value)}</div></div>
        <div>${card.trend !== undefined ? `<span class="metric-trend ${card.trend > 0 ? "up" : card.trend < 0 ? "down" : "neutral"}">${card.trend > 0 ? "+" : ""}${card.trend.toFixed(1)}%</span>` : ""}<p class="metric-note">${Utils.esc(card.note || "")}</p></div>
      </article>`).join("");
  }
};

function metrics() {
  const f = AppState.filtered.farmers;
  const p = AppState.filtered.payments;
  const d = AppState.filtered.delivery;
  const enrolled = f.length;
  const paid = p.length;
  const delivered = d.length;
  const revenue = Utils.sumBy(p, "Actual Paid (TZS)");
  const cost = Utils.sumBy(p, "Package Cost (TZS)");
  const collectionRate = cost ? (revenue / cost) * 100 : 0;
  const female = f.filter(row => row.GENDER === "Female").length;
  const femaleRatio = enrolled ? (female / enrolled) * 100 : 0;
  const villages = Utils.unique(f, "Village").length;
  const zones = Utils.unique(f, "Zone").length;
  const householdReach = Utils.sumBy(f, "MEMBERS IN HOUSEHOLD");
  const avgDelivery = Utils.avgBy(d, "Enrollment to delivery Days");
  const avgAfterDelivery = Utils.avgBy(d, "Days After Delivery");
  const avgNSP = Utils.avgBy(d, "NSP Score");
  const endToEnd = enrolled ? (delivered / enrolled) * 100 : 0;
  const enrollToPay = enrolled ? (paid / enrolled) * 100 : 0;
  const payToDeliver = paid ? (delivered / paid) * 100 : 0;
  const byMonth = {};
  f.forEach(row => { const key = Utils.monthKey(row["DATE CREATED"]); if (key !== "Unknown") byMonth[key] = (byMonth[key] || 0) + 1; });
  const months = Object.keys(byMonth).sort();
  let growth = 0;
  if (months.length >= 2) {
    const current = byMonth[months.at(-1)];
    const previous = byMonth[months.at(-2)];
    growth = previous ? ((current - previous) / previous) * 100 : 0;
  }
  const health = Math.round((Math.min(100, Math.max(0, growth + 60)) + Math.min(100, collectionRate) + Math.min(100, endToEnd) + Math.min(100, femaleRatio * 2) + (avgNSP ? Math.min(100, avgNSP) : 55)) / 5);
  return { f, p, d, enrolled, paid, delivered, revenue, cost, collectionRate, femaleRatio, villages, zones, householdReach, avgDelivery, avgAfterDelivery, avgNSP, endToEnd, enrollToPay, payToDeliver, growth, health, byMonth, months };
}

const Insight = {
  card(title, text) { return `<article class="insight-card"><h3>${Utils.esc(title)}</h3><p>${Utils.esc(text)}</p></article>`; },
  finding(title, text, tone = "positive") { return `<div class="finding ${tone}"><h3>${Utils.esc(title)}</h3><p>${Utils.esc(text)}</p></div>`; },
  rec(title, text, tone = "low") { return `<div class="recommendation ${tone}"><h3>${Utils.esc(title)}</h3><p>${Utils.esc(text)}</p></div>`; }
};

function rateByMonth(rows, dateKey) {
  const counts = {};
  rows.forEach(row => { const key = Utils.monthKey(row[dateKey]); if (key !== "Unknown") counts[key] = (counts[key] || 0) + 1; });
  return counts;
}

function growthRates(counts, months) {
  return months.map((month, index) => {
    if (index === 0) return 0;
    const prev = counts[months[index - 1]] || 0;
    const cur = counts[month] || 0;
    return prev ? Number(((cur - prev) / prev * 100).toFixed(1)) : 0;
  });
}

const PageEngines = {
  overview() {
    const m = metrics();
    $("#overview-health-score").textContent = String(m.health);
    $("#overview-health-label").textContent = m.health >= 75 ? "Strong" : m.health >= 55 ? "Watch" : "Needs action";
    $("#overview-narrative").textContent = `The program has enrolled ${Utils.fmtNum(m.enrolled)} farmers across ${m.zones} zones and ${m.villages} villages. Collection is ${Utils.fmtPct(m.collectionRate)}, delivery completion is ${Utils.fmtPct(m.endToEnd)}, and female participation is ${Utils.fmtPct(m.femaleRatio)}.`;
    KPIEngine.render("overview-kpis", [
      { label: "Total Enrolled", value: Utils.fmtNum(m.enrolled), raw: m.enrolled, trend: m.growth, note: "Unique farmers registered." },
      { label: "Total Delivered", value: Utils.fmtNum(m.delivered), raw: m.delivered, note: "Farmers receiving packages." },
      { label: "Revenue Collected", value: Utils.fmtCurrency(m.revenue), raw: m.revenue, note: `${Utils.fmtPct(m.collectionRate)} collection rate.` },
      { label: "Collection Rate", value: Utils.fmtPct(m.collectionRate), raw: m.collectionRate, note: "Actual paid over package cost." },
      { label: "Avg Delivery Time", value: m.avgDelivery ? `${m.avgDelivery.toFixed(0)}d` : "-", raw: m.avgDelivery, note: "Enrollment to delivery." },
      { label: "Villages Reached", value: Utils.fmtNum(m.villages), raw: m.villages, note: "Rural footprint." },
      { label: "Female Participation", value: Utils.fmtPct(m.femaleRatio), raw: m.femaleRatio, note: "Gender inclusion." },
      { label: "Household Reach", value: Utils.fmtNum(m.householdReach), raw: m.householdReach, note: "Indirect reach." },
      { label: "Outstanding Balance", value: Utils.fmtCurrency(Math.max(0, m.cost - m.revenue)), raw: Math.max(0, m.cost - m.revenue), note: "Still to collect." },
      { label: "Avg NSP Score", value: m.avgNSP ? m.avgNSP.toFixed(1) : "-", raw: m.avgNSP, note: "Satisfaction signal." }
    ]);
    ChartEngine.line("ov-enrollment-trend", m.months.map(Utils.monthLabel), [{ name: "Enrollments", data: m.months.map(key => m.byMonth[key]) }]);
    ChartEngine.donut("ov-gender-chart", Object.entries(Utils.countBy(m.f, "GENDER")).map(([name, value]) => ({ name, value })));
    const zoneRevenue = {};
    m.p.forEach(row => { zoneRevenue[row.Zone || "Unknown"] = (zoneRevenue[row.Zone || "Unknown"] || 0) + Utils.num(row["Actual Paid (TZS)"]); });
    const revenueEntries = Utils.entriesSorted(zoneRevenue, 8);
    ChartEngine.bar("ov-revenue-zone", revenueEntries.map(([key]) => key), [{ name: "Revenue", data: revenueEntries.map(([, value]) => Math.round(value)) }]);
    const packageEntries = Utils.entriesSorted(Utils.countBy(m.f, "PACKAGE"), 8);
    ChartEngine.donut("ov-package-chart", packageEntries.map(([name, value]) => ({ name, value })));
    $("#overview-insights").innerHTML = [
      Insight.card("Primary action", m.collectionRate < 75 ? "Prioritize collection follow-up in low-performing zones before scaling enrollment." : "Collection health can support controlled expansion."),
      Insight.card("Operational signal", m.avgDelivery > 45 ? `Average delivery time is ${m.avgDelivery.toFixed(0)} days, suggesting a logistics bottleneck.` : "Delivery timelines are within a manageable band."),
      Insight.card("Inclusion signal", m.femaleRatio >= 50 ? "Female participation is at or above parity." : "Female participation is below parity and needs targeted outreach."),
      Insight.card("Scale signal", `${m.villages} villages are active, providing a base for adjacent expansion.`)
    ].join("");
  },
  enrollment() {
    const m = metrics();
    const groupRate = m.enrolled ? (m.f.filter(row => ["yes", "true", "1"].includes(String(row["BELONGS IN A GROUP?"]).toLowerCase())).length / m.enrolled) * 100 : 0;
    const activeGroups = Object.keys(Utils.countBy(m.f, "GROUP NAME")).filter(key => key && key !== "Unknown").length;
    const avgHH = m.enrolled ? m.householdReach / m.enrolled : 0;
    const under5 = Utils.sumBy(m.f, "CHILDREN UNDER 5 YEARS OLD");
    const breastfeeding = Utils.sumBy(m.f, "BREAST FEEDING CHILDREN");
    const vulnerable = m.f.filter(row => Utils.num(row["CHILDREN UNDER 5 YEARS OLD"]) > 0 || Utils.num(row["BREAST FEEDING CHILDREN"]) > 0).length;
    const vulnRatio = m.enrolled ? vulnerable / m.enrolled * 100 : 0;
    const zoneScore = Math.min(100, m.zones * 10);
    const inclusionScore = (Math.min(100, m.femaleRatio * 2) + groupRate + zoneScore) / 3;
    KPIEngine.render("enr-biz-kpis", [
      { label: "Total Enrollments", value: Utils.fmtNum(m.enrolled), raw: m.enrolled, trend: m.growth, note: "Core farmer intake." },
      { label: "Enrollment Growth Rate", value: Utils.fmtPct(m.growth), raw: Math.abs(m.growth), note: "Month-over-month change." },
      { label: "Female Ratio", value: Utils.fmtPct(m.femaleRatio), raw: m.femaleRatio, note: "Gender participation." },
      { label: "Group Membership Rate", value: Utils.fmtPct(groupRate), raw: groupRate, note: "Community structure." },
      { label: "Avg Household Reach", value: avgHH.toFixed(1), raw: avgHH, note: "People reached per farmer." },
      { label: "Zones Reached", value: Utils.fmtNum(m.zones), raw: m.zones, note: "Operational spread." },
      { label: "Villages Reached", value: Utils.fmtNum(m.villages), raw: m.villages, note: "Rural footprint." },
      { label: "Active Groups", value: Utils.fmtNum(activeGroups), raw: activeGroups, note: "Community groups." },
      { label: "Household Impact", value: Utils.fmtNum(m.householdReach), raw: m.householdReach, note: "Indirect impact." },
      { label: "Current Month Enrolled", value: Utils.fmtNum(m.byMonth[m.months.at(-1)] || 0), raw: m.byMonth[m.months.at(-1)] || 0, note: "Latest intake." }
    ]);
    KPIEngine.render("enr-impact-kpis", [
      { label: "Household Impact Reach", value: Utils.fmtNum(m.householdReach), raw: m.householdReach, note: "Total household members." },
      { label: "Children Under 5", value: Utils.fmtNum(under5), raw: under5, note: "Nutrition-sensitive reach." },
      { label: "Breastfeeding Coverage", value: Utils.fmtNum(breastfeeding), raw: breastfeeding, note: "Breastfeeding children." },
      { label: "Women Participation Rate", value: Utils.fmtPct(m.femaleRatio), raw: m.femaleRatio, note: "Female farmers." },
      { label: "Vulnerable Household Ratio", value: Utils.fmtPct(vulnRatio), raw: vulnRatio, note: "Priority households." },
      { label: "Social Cohesion Index", value: Utils.fmtPct(groupRate), raw: groupRate, note: "Group participation." },
      { label: "Nutrition Exposure Estimate", value: Utils.fmtNum(m.householdReach + under5), raw: m.householdReach + under5, note: "Households plus children U5." },
      { label: "Rural Outreach Strength", value: Utils.fmtNum(m.villages), raw: m.villages, note: "Villages reached." },
      { label: "Inclusion Equity Score", value: Utils.fmtPct(inclusionScore), raw: inclusionScore, note: "Composite inclusion." },
      { label: "Vulnerable Households", value: Utils.fmtNum(vulnerable), raw: vulnerable, note: "Households with U5/breastfeeding." }
    ]);
    ChartEngine.line("enr-velocity", m.months.map(Utils.monthLabel), [{ name: "Enrollments", data: m.months.map(key => m.byMonth[key]) }]);
    ChartEngine.donut("enr-gender", Object.entries(Utils.countBy(m.f, "GENDER")).map(([name, value]) => ({ name, value })));
    ChartEngine.bar("enr-vano", Utils.entriesSorted(Utils.countBy(m.f, "VANO"), 12).map(([key]) => key), [{ name: "Farmers", data: Utils.entriesSorted(Utils.countBy(m.f, "VANO"), 12).map(([, value]) => value) }], { horizontal: true });
    const packages = Utils.entriesSorted(Utils.countBy(m.f, "PACKAGE"), 10);
    ChartEngine.bar("enr-package", packages.map(([key]) => key), [{ name: "Farmers", data: packages.map(([, value]) => value) }]);
    ChartEngine.bar("enr-growth", m.months.map(Utils.monthLabel), [{ name: "Growth %", data: growthRates(m.byMonth, m.months), color: Utils.cssVar("--gold") }], { valueFormatter: "{value}%" });
    const groups = Utils.entriesSorted(Utils.countBy(m.f, "GROUP NAME"), 15);
    ChartEngine.bar("enr-groups", groups.map(([key]) => key), [{ name: "Members", data: groups.map(([, value]) => value) }], { horizontal: true });
    const zones = Utils.entriesSorted(Utils.countBy(m.f, "Zone"), 12);
    ChartEngine.bar("enr-zone", zones.map(([key]) => key), [{ name: "Enrollments", data: zones.map(([, value]) => value) }]);
    ChartEngine.donut("enr-positions", Utils.entriesSorted(Utils.countBy(m.f, "POSITION IN GROUP"), 10).map(([name, value]) => ({ name, value })));
    ChartEngine.funnel("enr-nutrition-funnel", [{ name: "Enrolled Farmers", value: m.enrolled }, { name: "Household Members", value: m.householdReach }, { name: "Children Under 5", value: under5 }, { name: "Breastfeeding", value: breastfeeding }]);
    ChartEngine.radar("enr-equity-radar", ["Gender", "Groups", "Zones", "Vulnerable", "Cohesion"], [Math.min(100, m.femaleRatio * 2), groupRate, zoneScore, Math.min(100, vulnRatio), groupRate]);
    ChartEngine.donut("enr-vulnerable", [{ name: "Vulnerable", value: vulnerable }, { name: "Other", value: Math.max(0, m.enrolled - vulnerable) }]);
    const villageEntries = Utils.entriesSorted(Utils.countBy(m.f, "Village"), 15);
    ChartEngine.bar("enr-villages", villageEntries.map(([key]) => key), [{ name: "Farmers", data: villageEntries.map(([, value]) => value) }], { horizontal: true });
    $("#enr-insights").innerHTML = [Insight.card("Growth", m.growth > 0 ? `Enrollment is growing at ${m.growth.toFixed(1)}% month over month.` : "Enrollment is flat or declining."), Insight.card("Field performance", Utils.entriesSorted(Utils.countBy(m.f, "VANO"), 1)[0] ? `${Utils.entriesSorted(Utils.countBy(m.f, "VANO"), 1)[0][0]} leads farmer intake.` : "No VANO signal available."), Insight.card("Demand", packages[0] ? `${packages[0][0]} is the strongest package signal.` : "Package data is limited."), Insight.card("Inclusion", m.femaleRatio >= 50 ? "Gender parity is healthy." : "Women-focused enrollment should be strengthened.")].join("");
  },
  payment() {
    const m = metrics();
    const outstanding = Math.max(0, m.cost - m.revenue);
    const actualUsd = Utils.sumBy(m.p, "Actual Paid (USD)");
    const avgInstallments = Utils.avgBy(m.p, "Payment Counts");
    const overpaid = m.p.filter(row => Utils.num(row["% of Package Cost Paid"]) > 100).length;
    const fullPaid = m.p.filter(row => Utils.num(row["% of Package Cost Paid"]) >= 100).length;
    const partial = m.p.filter(row => { const v = Utils.num(row["% of Package Cost Paid"]); return v > 0 && v < 100; }).length;
    const zero = m.p.filter(row => Utils.num(row["% of Package Cost Paid"]) === 0).length;
    const delivered = m.p.filter(row => String(row["Delivered Status"]).toLowerCase() === "delivered").length;
    const deliveryRate = m.p.length ? delivered / m.p.length * 100 : 0;
    const installmentPayers = m.p.filter(row => Utils.num(row["Payment Counts"]) > 1).length;
    const financialInclusion = m.p.length ? installmentPayers / m.p.length * 100 : 0;
    const villages = Utils.unique(m.p, "Village").length;
    const avgRevenue = m.p.length ? m.revenue / m.p.length : 0;
    const reliability = Math.min(100, (Utils.avgBy(m.p, "% of Package Cost Paid") + avgInstallments * 10) / 2);
    KPIEngine.render("pay-biz-kpis", [
      { label: "Revenue Collected (TZS)", value: Utils.fmtCurrency(m.revenue), raw: m.revenue, note: "Actual paid." },
      { label: "Revenue (USD)", value: `$${Utils.fmtNum(actualUsd)}`, raw: actualUsd, note: "USD equivalent." },
      { label: "Collection Rate", value: Utils.fmtPct(m.collectionRate), raw: m.collectionRate, note: "Paid over expected." },
      { label: "Outstanding Balance", value: Utils.fmtCurrency(outstanding), raw: outstanding, note: "Remaining balance." },
      { label: "Delivery Rate", value: Utils.fmtPct(deliveryRate), raw: deliveryRate, note: "Paid records delivered." },
      { label: "Avg Payment Installments", value: avgInstallments.toFixed(1), raw: avgInstallments, note: "Payment frequency." },
      { label: "Packages Expected", value: Utils.fmtNum(m.p.length), raw: m.p.length, note: "Payment records." },
      { label: "Overpayment Rate", value: Utils.fmtPct(m.p.length ? overpaid / m.p.length * 100 : 0), raw: overpaid, note: "Above 100% paid." },
      { label: "Avg Revenue / Farmer", value: Utils.fmtCurrency(avgRevenue), raw: avgRevenue, note: "Revenue intensity." },
      { label: "Fully Paid Farmers", value: Utils.fmtNum(fullPaid), raw: fullPaid, note: "Paid 100% or more." }
    ]);
    KPIEngine.render("pay-impact-kpis", [
      { label: "Farmers Reached", value: Utils.fmtNum(m.p.length), raw: m.p.length, note: "Payment participants." },
      { label: "Delivery Impact Rate", value: Utils.fmtPct(deliveryRate), raw: deliveryRate, note: "Payment-to-delivery." },
      { label: "Financial Inclusion Rate", value: Utils.fmtPct(financialInclusion), raw: financialInclusion, note: "Installment users." },
      { label: "Rural Villages Active", value: Utils.fmtNum(villages), raw: villages, note: "Payment villages." },
      { label: "Payment Reliability Score", value: Utils.fmtPct(reliability), raw: reliability, note: "Composite payment quality." },
      { label: "Partial Payers", value: Utils.fmtNum(partial), raw: partial, note: "Collection opportunity." },
      { label: "Economic Value (TZS)", value: Utils.fmtCurrency(m.revenue), raw: m.revenue, note: "Community participation." },
      { label: "High Commitment Farmers", value: Utils.fmtNum(fullPaid), raw: fullPaid, note: "Program champions." },
      { label: "Zero-Payment Farmers", value: Utils.fmtNum(zero), raw: zero, note: "Urgent follow-up." }
    ]);
    const zoneRevenue = {};
    m.p.forEach(row => { zoneRevenue[row.Zone || "Unknown"] = (zoneRevenue[row.Zone || "Unknown"] || 0) + Utils.num(row["Actual Paid (TZS)"]); });
    const zr = Utils.entriesSorted(zoneRevenue, 10);
    ChartEngine.bar("pay-zone-revenue", zr.map(([key]) => key), [{ name: "Revenue", data: zr.map(([, value]) => Math.round(value)) }]);
    const buckets = { "0%": zero, "1-25%": 0, "26-50%": 0, "51-75%": 0, "76-99%": 0, "100%+": 0 };
    m.p.forEach(row => { const v = Utils.num(row["% of Package Cost Paid"]); if (v > 0 && v <= 25) buckets["1-25%"]++; else if (v <= 50 && v > 25) buckets["26-50%"]++; else if (v <= 75 && v > 50) buckets["51-75%"]++; else if (v < 100 && v > 75) buckets["76-99%"]++; else if (v >= 100) buckets["100%+"]++; });
    ChartEngine.donut("pay-collection-dist", Object.entries(buckets).map(([name, value]) => ({ name, value })));
    const agents = {};
    m.p.forEach(row => { agents[row["Village Agent"] || "Unknown"] = (agents[row["Village Agent"] || "Unknown"] || 0) + Utils.num(row["Actual Paid (TZS)"]); });
    const agentEntries = Utils.entriesSorted(agents, 12);
    ChartEngine.bar("pay-agent-leaderboard", agentEntries.map(([key]) => key), [{ name: "Revenue", data: agentEntries.map(([, value]) => Math.round(value)) }], { horizontal: true });
    const pkg = Utils.entriesSorted(Utils.countBy(m.p, "Package Name"), 12);
    ChartEngine.bar("pay-package-demand", pkg.map(([key]) => key), [{ name: "Farmers", data: pkg.map(([, value]) => value) }], { horizontal: true });
    ChartEngine.funnel("pay-delivery-funnel", [{ name: "Total Farmers", value: m.p.length }, { name: "Fully Paid", value: fullPaid }, { name: "Delivered", value: delivered }, { name: "Partial Payers", value: partial }]);
    ChartEngine.donut("pay-payment-segments", [{ name: "Overpaid", value: overpaid }, { name: "Fully Paid", value: Math.max(0, fullPaid - overpaid) }, { name: "Partial", value: partial }, { name: "Unpaid", value: zero }]);
    const adoption = rateByMonth(m.p, "Date of Finishing");
    const months = Object.keys(adoption).sort();
    ChartEngine.line("pay-adoption-growth", months.map(Utils.monthLabel), [{ name: "New Farmers", data: months.map(key => adoption[key]) }]);
    let running = 0;
    ChartEngine.line("pay-cumulative", months.map(Utils.monthLabel), [{ name: "Cumulative Revenue", data: months.map(key => Math.round(running += m.p.filter(row => Utils.monthKey(row["Date of Finishing"]) === key).reduce((sum, row) => sum + Utils.num(row["Actual Paid (TZS)"]), 0))) }]);
    const villageEntries = Utils.entriesSorted(Utils.countBy(m.p, "Village"), 15);
    ChartEngine.bar("pay-village-reach", villageEntries.map(([key]) => key), [{ name: "Farmers", data: villageEntries.map(([, value]) => value) }], { horizontal: true });
    const highByVillage = Utils.entriesSorted(Utils.countBy(m.p.filter(row => Utils.num(row["% of Package Cost Paid"]) >= 100), "Village"), 10);
    ChartEngine.bar("pay-high-commitment", highByVillage.map(([key]) => key), [{ name: "High Commitment", data: highByVillage.map(([, value]) => value) }]);
    $("#pay-insights").innerHTML = [Insight.card("Collection priority", outstanding > 0 ? `${Utils.fmtCurrency(outstanding)} remains outstanding.` : "Expected package value has been collected."), Insight.card("Agent accountability", agentEntries[0] ? `${agentEntries[0][0]} leads collections with ${Utils.fmtCurrency(agentEntries[0][1])}.` : "No agent signal available."), Insight.card("Payment quality", `${fullPaid} farmers have fully paid or overpaid.`), Insight.card("Revenue health", m.collectionRate >= 80 ? "Collection health is strong." : "Collection rate is below target.")].join("");
  },
  delivery() {
    const m = metrics();
    const totalInvest = Utils.sumBy(m.d, "Total Investment");
    const avgInvest = m.d.length ? totalInvest / m.d.length : 0;
    const villages = Utils.unique(m.d, "Village").length;
    const zones = Utils.unique(m.d, "Zone").length;
    const female = m.d.filter(row => row.Gender === "Female").length;
    const femaleRatio = m.d.length ? female / m.d.length * 100 : 0;
    const deliveryByMonth = rateByMonth(m.d, "Delivery Date");
    const months = Object.keys(deliveryByMonth).sort();
    const deliveryGrowth = months.length >= 2 && deliveryByMonth[months.at(-2)] ? (deliveryByMonth[months.at(-1)] - deliveryByMonth[months.at(-2)]) / deliveryByMonth[months.at(-2)] * 100 : 0;
    const accessibility = m.avgDelivery ? Math.max(0, 100 - m.avgDelivery / 90 * 100) : 0;
    KPIEngine.render("del-biz-kpis", [
      { label: "Total Farmers Delivered", value: Utils.fmtNum(m.d.length), raw: m.d.length, trend: deliveryGrowth, note: "Delivery completions." },
      { label: "Total Investment (TZS)", value: Utils.fmtCurrency(totalInvest), raw: totalInvest, note: "Distributed value." },
      { label: "Avg Delivery Time (days)", value: m.avgDelivery ? `${m.avgDelivery.toFixed(0)}d` : "-", raw: m.avgDelivery, note: "Operational speed." },
      { label: "Delivery Growth Rate", value: Utils.fmtPct(deliveryGrowth), raw: deliveryGrowth, note: "Month-over-month." },
      { label: "Avg Revenue / Farmer", value: Utils.fmtCurrency(avgInvest), raw: avgInvest, note: "Average investment." },
      { label: "Avg NSP Score", value: m.avgNSP ? m.avgNSP.toFixed(1) : "-", raw: m.avgNSP, note: "Satisfaction." },
      { label: "Villages Reached", value: Utils.fmtNum(villages), raw: villages, note: "Delivery footprint." },
      { label: "Female Farmer Ratio", value: Utils.fmtPct(femaleRatio), raw: femaleRatio, note: "Delivery inclusion." },
      { label: "Avg Days Since Delivery", value: m.avgAfterDelivery ? `${m.avgAfterDelivery.toFixed(0)}d` : "-", raw: m.avgAfterDelivery, note: "Post-delivery window." },
      { label: "Program Accessibility Rate", value: Utils.fmtPct(accessibility), raw: accessibility, note: "Speed score." }
    ]);
    KPIEngine.render("del-impact-kpis", [
      { label: "Farmers Impacted", value: Utils.fmtNum(m.d.length), raw: m.d.length, note: "Delivery impact." },
      { label: "Women Empowerment Score", value: Utils.fmtPct(femaleRatio), raw: femaleRatio, note: "Women delivered." },
      { label: "Farmer Satisfaction (NSP)", value: m.avgNSP ? m.avgNSP.toFixed(1) : "-", raw: m.avgNSP, note: "Satisfaction score." },
      { label: "Avg Days Since Delivery", value: m.avgAfterDelivery ? `${m.avgAfterDelivery.toFixed(0)}d` : "-", raw: m.avgAfterDelivery, note: "Follow-up maturity." },
      { label: "Rural Reach Index", value: Utils.fmtNum(villages), raw: villages, note: "Villages served." },
      { label: "Livelihood Investment", value: Utils.fmtCurrency(totalInvest), raw: totalInvest, note: "Economic value." },
      { label: "Program Accessibility", value: Utils.fmtPct(accessibility), raw: accessibility, note: "Access speed." },
      { label: "Zones Covered", value: Utils.fmtNum(zones), raw: zones, note: "Regional coverage." },
      { label: "Delivery Growth Rate", value: Utils.fmtPct(deliveryGrowth), raw: deliveryGrowth, note: "Scale momentum." },
      { label: "Geographic Equity Score", value: Utils.fmtPct(zones ? Math.min(villages / zones * 10, 100) : 0), raw: zones ? Math.min(villages / zones * 10, 100) : 0, note: "Villages per zone." }
    ]);
    ChartEngine.bar("del-growth", months.map(Utils.monthLabel), [{ name: "Growth %", data: growthRates(deliveryByMonth, months), color: Utils.cssVar("--gold") }], { valueFormatter: "{value}%" });
    ChartEngine.donut("del-package-dist", Utils.entriesSorted(Utils.countBy(m.d, "Package Delivered"), 10).map(([name, value]) => ({ name, value })));
    const zoneEntries = Utils.entriesSorted(Utils.countBy(m.d, "Zone"), 12);
    ChartEngine.bar("del-zone-coverage", zoneEntries.map(([key]) => key), [{ name: "Deliveries", data: zoneEntries.map(([, value]) => value) }]);
    const nspZones = Utils.unique(m.d, "Zone").filter(Boolean);
    ChartEngine.bar("del-nsp-dist", nspZones, [{ name: "NSP", data: nspZones.map(zone => Number(Utils.avgBy(m.d.filter(row => row.Zone === zone), "NSP Score").toFixed(1))) }]);
    const byDate = {};
    m.d.forEach(row => { if (row["Delivery Date"]) byDate[row["Delivery Date"]] = (byDate[row["Delivery Date"]] || 0) + 1; });
    const dates = Object.keys(byDate).sort().slice(-30);
    ChartEngine.line("del-daily", dates.map(Utils.dateLabel), [{ name: "Deliveries", data: dates.map(key => byDate[key]) }]);
    const buckets = { "0-7d": 0, "8-14d": 0, "15-30d": 0, "31-60d": 0, "61-90d": 0, "90d+": 0 };
    m.d.forEach(row => { const v = Utils.num(row["Enrollment to delivery Days"]); if (v <= 7) buckets["0-7d"]++; else if (v <= 14) buckets["8-14d"]++; else if (v <= 30) buckets["15-30d"]++; else if (v <= 60) buckets["31-60d"]++; else if (v <= 90) buckets["61-90d"]++; else buckets["90d+"]++; });
    ChartEngine.bar("del-timeline", Object.keys(buckets), [{ name: "Farmers", data: Object.values(buckets) }]);
    const investByZone = {};
    m.d.forEach(row => { investByZone[row.Zone || "Unknown"] = (investByZone[row.Zone || "Unknown"] || 0) + Utils.num(row["Total Investment"]); });
    ChartEngine.donut("del-zone-equity", Utils.entriesSorted(investByZone, 12).map(([name, value]) => ({ name, value: Math.round(value) })));
    ChartEngine.line("del-engagement", months.map(Utils.monthLabel), [{ name: "Deliveries", data: months.map(key => deliveryByMonth[key]) }]);
    ChartEngine.donut("del-gender", Object.entries(Utils.countBy(m.d, "Gender")).map(([name, value]) => ({ name, value })));
    const villagesEntries = Utils.entriesSorted(Utils.countBy(m.d, "Village"), 15);
    ChartEngine.bar("del-village-reach", villagesEntries.map(([key]) => key), [{ name: "Deliveries", data: villagesEntries.map(([, value]) => value) }], { horizontal: true });
    $("#del-insights").innerHTML = [Insight.card("Fulfillment", `${Utils.fmtNum(m.d.length)} farmers have received packages.`), Insight.card("Speed", m.avgDelivery > 45 ? "Delivery speed is above target and requires logistics review." : "Delivery speed is within a manageable range."), Insight.card("Experience", m.avgNSP ? `Average NSP score is ${m.avgNSP.toFixed(1)}.` : "NSP satisfaction data is limited."), Insight.card("Equity", `${Utils.fmtPct(femaleRatio)} of delivered farmers are women.`)].join("");
  },
  geo() {
    const m = metrics();
    const mapped = m.f.filter(GeoEngine.validCoord);
    const vanoCount = Object.keys(Utils.countBy(m.f, "VANO")).filter(key => key && key !== "Unknown").length;
    KPIEngine.render("geo-kpis", [
      { label: "Farmers Mapped", value: Utils.fmtNum(mapped.length), raw: mapped.length, note: "Records with coordinates." },
      { label: "Zones Covered", value: Utils.fmtNum(m.zones), raw: m.zones, note: "Operational zones." },
      { label: "Villages Reached", value: Utils.fmtNum(m.villages), raw: m.villages, note: "Rural footprint." },
      { label: "Top Zone", value: Utils.entriesSorted(Utils.countBy(m.f, "Zone"), 1)[0]?.[0] || "-", note: "Highest enrollment zone." },
      { label: "Active VANOs", value: Utils.fmtNum(vanoCount), raw: vanoCount, note: "Field officers." },
      { label: "Coverage Rate", value: Utils.fmtPct(m.enrolled ? mapped.length / m.enrolled * 100 : 0), raw: mapped.length, note: "GPS completeness." }
    ]);
    GeoEngine.renderMap(m.f, mapped);
    const zones = Utils.entriesSorted(Utils.countBy(m.f, "Zone"), 12);
    ChartEngine.bar("geo-zone-penetration", zones.map(([key]) => key), [{ name: "Farmers", data: zones.map(([, value]) => value) }]);
    const villages = Utils.entriesSorted(Utils.countBy(m.f, "Village"), 15);
    ChartEngine.bar("geo-village-density", villages.map(([key]) => key), [{ name: "Farmers", data: villages.map(([, value]) => value) }], { horizontal: true });
    const vanoTerritory = Object.entries(Utils.groupBy(m.f, "VANO")).map(([key, rows]) => [key, Utils.unique(rows, "Village").length]).filter(([key]) => key !== "Unknown").sort((a,b) => b[1] - a[1]).slice(0,12);
    ChartEngine.bar("geo-vano-coverage", vanoTerritory.map(([key]) => key), [{ name: "Villages", data: vanoTerritory.map(([, value]) => value) }], { horizontal: true });
    const zoneNames = Utils.unique(m.f, "Zone").filter(Boolean);
    ChartEngine.bar("geo-gender-zone", zoneNames, [{ name: "Female", data: zoneNames.map(zone => m.f.filter(row => row.Zone === zone && row.GENDER === "Female").length), stack: "gender" }, { name: "Male", data: zoneNames.map(zone => m.f.filter(row => row.Zone === zone && row.GENDER === "Male").length), stack: "gender", color: Utils.cssVar("--gold") }]);
    const pkgs = Utils.unique(m.f, "PACKAGE").filter(Boolean).slice(0, 5);
    ChartEngine.bar("geo-package-zone", zoneNames, pkgs.map((pkg, index) => ({ name: pkg, data: zoneNames.map(zone => m.f.filter(row => row.Zone === zone && row.PACKAGE === pkg).length), stack: "pkg", color: ChartEngine.palette()[index % ChartEngine.palette().length] })));
    ChartEngine.bar("geo-hh-zone", zoneNames, [{ name: "Avg Household", data: zoneNames.map(zone => Number(Utils.avgBy(m.f.filter(row => row.Zone === zone), "MEMBERS IN HOUSEHOLD").toFixed(1))) }]);
    $("#geo-insights").innerHTML = [Insight.card("Territory signal", `${m.zones} zones and ${m.villages} villages are active.`), Insight.card("Data quality", `${Utils.fmtPct(m.enrolled ? mapped.length / m.enrolled * 100 : 0)} of farmers have usable GPS coordinates.`), Insight.card("Expansion", villages[0] ? `${villages[0][0]} is the highest-density village.` : "Village density data is limited."), Insight.card("Resource allocation", zones[0] ? `${zones[0][0]} is the strongest enrollment territory.` : "Zone signal is limited.")].join("");
  },
  pipeline() {
    const m = metrics();
    const dropEnrollPay = Math.max(0, m.enrolled - m.paid);
    const dropPayDeliver = Math.max(0, m.paid - m.delivered);
    KPIEngine.render("pipe-kpis", [
      { label: "Enrolled", value: Utils.fmtNum(m.enrolled), raw: m.enrolled, note: "Pipeline start." },
      { label: "Converted to Payment", value: Utils.fmtNum(m.paid), raw: m.paid, note: "Payment records." },
      { label: "Delivered", value: Utils.fmtNum(m.delivered), raw: m.delivered, note: "Pipeline completion." },
      { label: "Enroll to Pay Rate", value: Utils.fmtPct(m.enrollToPay), raw: m.enrollToPay, note: "Payment conversion." },
      { label: "Pay to Deliver Rate", value: Utils.fmtPct(m.payToDeliver), raw: m.payToDeliver, note: "Fulfillment conversion." },
      { label: "End-to-End Rate", value: Utils.fmtPct(m.endToEnd), raw: m.endToEnd, note: "Full journey." },
      { label: "Enroll Drop-offs", value: Utils.fmtNum(dropEnrollPay), raw: dropEnrollPay, note: "Not yet paid." },
      { label: "Payment Drop-offs", value: Utils.fmtNum(dropPayDeliver), raw: dropPayDeliver, note: "Paid, not delivered." },
      { label: "Avg Delivery Time", value: m.avgDelivery ? `${m.avgDelivery.toFixed(0)}d` : "-", raw: m.avgDelivery, note: "Speed to delivery." }
    ]);
    $("#pipe-funnel-stages").innerHTML = [{ title: "Enrolled Farmers", count: m.enrolled, pct: 100, note: "Farmers registered in the program." }, { title: "Made Payment", count: m.paid, pct: m.enrollToPay, note: `${dropEnrollPay} farmers have not converted to payment.` }, { title: "Package Delivered", count: m.delivered, pct: m.endToEnd, note: `${dropPayDeliver} paying farmers still need fulfillment.` }].map(stage => `<div class="pipeline-stage"><div><h3>${stage.title}</h3><p>${stage.note}</p></div><strong>${Utils.fmtNum(stage.count)}</strong><div class="pipeline-bar"><span style="--w:${Math.max(0, Math.min(stage.pct, 100))}%"></span></div></div>`).join("");
    const zones = Utils.unique(m.f, "Zone").filter(Boolean);
    ChartEngine.bar("pipe-zone-conversion", zones, [{ name: "Enroll to Pay", data: zones.map(zone => { const enrolled = m.f.filter(row => row.Zone === zone).length; const paid = m.p.filter(row => row.Zone === zone).length; return enrolled ? Number((paid / enrolled * 100).toFixed(1)) : 0; }) }], { valueFormatter: "{value}%" });
    ChartEngine.bar("pipe-delivery-conversion", zones, [{ name: "Pay to Deliver", data: zones.map(zone => { const paid = m.p.filter(row => row.Zone === zone).length; const delivered = m.d.filter(row => row.Zone === zone).length; return paid ? Number((delivered / paid * 100).toFixed(1)) : 0; }) }], { valueFormatter: "{value}%" });
    const fM = rateByMonth(m.f, "DATE CREATED"), pM = rateByMonth(m.p, "Date of Finishing"), dM = rateByMonth(m.d, "Delivery Date");
    const months = [...new Set([...Object.keys(fM), ...Object.keys(pM), ...Object.keys(dM)])].sort();
    ChartEngine.line("pipe-velocity", months.map(Utils.monthLabel), [{ name: "Enrolled", data: months.map(key => fM[key] || 0) }, { name: "Payments", data: months.map(key => pM[key] || 0) }, { name: "Delivered", data: months.map(key => dM[key] || 0) }]);
    const pkgs = [...new Set([...Utils.unique(m.f, "PACKAGE"), ...Utils.unique(m.d, "Package Delivered")].filter(Boolean))].slice(0,10);
    ChartEngine.bar("pipe-package-conversion", pkgs, [{ name: "Enrolled", data: pkgs.map(pkg => m.f.filter(row => row.PACKAGE === pkg).length) }, { name: "Delivered", data: pkgs.map(pkg => m.d.filter(row => row["Package Delivered"] === pkg).length), color: Utils.cssVar("--gold") }]);
    $("#pipe-insights").innerHTML = [Insight.card("Conversion", `${Utils.fmtPct(m.endToEnd)} of enrolled farmers have completed the journey.`), Insight.card("Payment gap", `${Utils.fmtNum(dropEnrollPay)} enrolled farmers need payment conversion.`), Insight.card("Delivery gap", `${Utils.fmtNum(dropPayDeliver)} paying farmers need fulfillment follow-up.`), Insight.card("Executive focus", m.enrollToPay < m.payToDeliver ? "Payment conversion is the larger bottleneck." : "Fulfillment is the larger bottleneck.")].join("");
  },
  strategy() {
    const m = metrics();
    const outstanding = Math.max(0, m.cost - m.revenue);
    $("#strategy-exec-text").textContent = `HERVeg.05 currently shows a program health score of ${m.health}/100. Enrollment stands at ${Utils.fmtNum(m.enrolled)}, revenue collection is ${Utils.fmtPct(m.collectionRate)}, delivery completion is ${Utils.fmtPct(m.endToEnd)}, and female participation is ${Utils.fmtPct(m.femaleRatio)}. Leadership should focus on the largest conversion constraint before expanding field activity.`;
    $("#strategy-trends").innerHTML = [Insight.finding("Enrollment momentum", m.growth > 0 ? `Enrollment is growing at ${m.growth.toFixed(1)}% month over month.` : "Enrollment is flat or declining.", m.growth > 0 ? "positive" : "warning"), Insight.finding("Collection performance", `${Utils.fmtPct(m.collectionRate)} of expected package value has been collected.`, m.collectionRate >= 80 ? "positive" : "warning"), Insight.finding("Delivery conversion", `${Utils.fmtPct(m.endToEnd)} of enrolled farmers have received delivery.`, m.endToEnd >= 70 ? "positive" : "warning")].join("");
    $("#strategy-risks").innerHTML = [outstanding > 0 ? Insight.finding("Outstanding balance", `${Utils.fmtCurrency(outstanding)} remains unpaid.`, outstanding > m.revenue * 0.25 ? "critical" : "warning") : Insight.finding("No outstanding balance", "Expected package value has been collected.", "positive"), m.femaleRatio < 50 ? Insight.finding("Gender parity gap", `Female participation is ${Utils.fmtPct(m.femaleRatio)}.`, "warning") : Insight.finding("Gender parity", "Female participation is at or above parity.", "positive"), m.avgDelivery > 45 ? Insight.finding("Delivery delay", `Average delivery time is ${m.avgDelivery.toFixed(0)} days.`, "critical") : Insight.finding("Delivery timing", "Delivery time is within the operating band.", "positive")].join("");
    const healthValues = [Math.min(100, Math.max(0, m.growth + 55)), Math.min(100, m.collectionRate), Math.min(100, m.endToEnd), Math.min(100, m.femaleRatio * 2), m.avgNSP ? Math.min(100, m.avgNSP) : 55];
    ChartEngine.radar("strategy-health-radar", ["Growth", "Collection", "Delivery", "Gender", "Satisfaction"], healthValues);
    const targets = { "Collection": 80, "Female": 50, "Conversion": 70, "NSP": 70 };
    const actuals = { "Collection": m.collectionRate, "Female": m.femaleRatio, "Conversion": m.endToEnd, "NSP": m.avgNSP || 0 };
    ChartEngine.bar("strategy-anomaly", Object.keys(targets), [{ name: "Gap to target", data: Object.keys(targets).map(key => Number((actuals[key] - targets[key]).toFixed(1))), color: Utils.cssVar("--gold") }], { valueFormatter: "{value}%" });
    const matrix = Utils.unique(m.f, "Zone").filter(Boolean).map(zone => { const farmers = m.f.filter(row => row.Zone === zone).length; const paidRows = m.p.filter(row => row.Zone === zone); const revenue = Utils.sumBy(paidRows, "Actual Paid (TZS)"); const cost = Utils.sumBy(paidRows, "Package Cost (TZS)"); const rate = cost ? Number((revenue / cost * 100).toFixed(1)) : 0; return { name: zone, value: [farmers, rate, farmers] }; });
    ChartEngine.scatter("strategy-opportunity", matrix);
    const fM = rateByMonth(m.f, "DATE CREATED"), pM = rateByMonth(m.p, "Date of Finishing"), dM = rateByMonth(m.d, "Delivery Date");
    const months = [...new Set([...Object.keys(fM), ...Object.keys(pM), ...Object.keys(dM)])].sort();
    ChartEngine.line("strategy-trend-comparison", months.map(Utils.monthLabel), [{ name: "Enrolled", data: months.map(key => fM[key] || 0) }, { name: "Payments", data: months.map(key => pM[key] || 0) }, { name: "Delivered", data: months.map(key => dM[key] || 0) }]);
    $("#strategy-recommendations").innerHTML = [m.collectionRate < 80 ? Insight.rec("Accelerate collections", `Recover ${Utils.fmtCurrency(outstanding)} through VANO follow-up in low-collection zones.`, "high") : Insight.rec("Maintain collection discipline", "Collection rate is healthy; keep monitoring partial payers.", "low"), m.endToEnd < 70 ? Insight.rec("Close fulfillment gap", "Prioritize paid farmers who have not yet received delivery.", "high") : Insight.rec("Scale operating cadence", "Fulfillment is strong enough to support controlled expansion.", "low"), m.femaleRatio < 50 ? Insight.rec("Strengthen women outreach", "Launch targeted enrollment through women farmer groups and village leaders.", "medium") : Insight.rec("Protect inclusion gains", "Document women outreach practices for replication.", "low")].join("");
    $("#strategy-opportunities").innerHTML = [Insight.rec("Replicate strongest villages", "Use high-density villages as referral hubs for neighboring communities.", "low"), Insight.rec("Coach underperforming agents", "Benchmark top VANO behavior and convert it into field coaching routines.", "medium"), Insight.rec("Improve data quality", "Standardize village, zone, and GPS capture to improve geographic targeting.", "medium")].join("");
  }
};

const GeoEngine = {
  validCoord(row) {
    const lat = parseFloat(row.LATITUDE || row.Latitude || "");
    const lng = parseFloat(row.LONGITUDE || row.Longitude || "");
    return !Number.isNaN(lat) && !Number.isNaN(lng) && lat !== 0 && lng !== 0 && Math.abs(lat) <= 90 && Math.abs(lng) <= 180;
  },
  coord(row) { return [parseFloat(row.LATITUDE || row.Latitude), parseFloat(row.LONGITUDE || row.Longitude)]; },
  renderMap(allRows, mappedRows) {
    if (!window.L) return;
    const element = document.getElementById("geo-map");
    if (!element) return;
    if (AppState.map) { AppState.map.remove(); AppState.map = null; }
    const center = mappedRows.length ? this.centroid(mappedRows.map(row => this.coord(row))) : [-6.369, 34.889];
    AppState.map = L.map(element, { zoomControl: true, attributionControl: false }).setView(center, mappedRows.length > 50 ? 8 : 10);
    L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", { subdomains: "abcd", maxZoom: 19, attribution: "OpenStreetMap, CartoDB" }).addTo(AppState.map);
    this.buildLayers(allRows, mappedRows);
    this.showMode(AppState.mapMode);
    setTimeout(() => AppState.map?.invalidateSize(), 100);
  },
  buildLayers(allRows, mappedRows) {
    const colors = ChartEngine.palette();
    const zones = Utils.unique(allRows, "Zone");
    const zoneColors = {};
    zones.forEach((zone, index) => { zoneColors[zone] = colors[index % colors.length]; });
    const cluster = L.markerClusterGroup({ chunkedLoading: true, maxClusterRadius: 52 });
    mappedRows.forEach(row => {
      const [lat, lng] = this.coord(row);
      const marker = L.circleMarker([lat, lng], { radius: 6, fillColor: zoneColors[row.Zone] || colors[0], fillOpacity: 0.82, color: "#fff", weight: 1 });
      marker.bindPopup(`<strong>${Utils.esc(row.NAME || "Farmer")}</strong><br>${Utils.esc(row.Village || "Unknown village")} - ${Utils.esc(row.Zone || "Unknown zone")}<br>${Utils.esc(row.PACKAGE || "No package recorded")}`);
      cluster.addLayer(marker);
    });
    const heat = L.heatLayer(mappedRows.map(row => [...this.coord(row), 0.5]), { radius: 25, blur: 18, maxZoom: 14, gradient: { 0.25: colors[1], 0.55: colors[2], 1: colors[4] } });
    const zoneLayer = L.layerGroup();
    zones.forEach(zone => {
      const rows = mappedRows.filter(row => row.Zone === zone);
      if (!rows.length) return;
      const center = this.centroid(rows.map(row => this.coord(row)));
      const radius = Math.min(Math.max(Math.sqrt(rows.length) * 1200, 2200), 22000);
      L.circle(center, { radius, color: zoneColors[zone], fillColor: zoneColors[zone], fillOpacity: 0.14, weight: 2 }).bindTooltip(`${zone}: ${rows.length} farmers`).addTo(zoneLayer);
    });
    AppState.mapLayers = { clusters: cluster, heat, zones: zoneLayer };
    $("#map-legend").innerHTML = Object.entries(zoneColors).slice(0, 10).map(([zone, color]) => `<span class="legend-item"><span class="legend-dot" style="background:${color}"></span>${Utils.esc(zone)}</span>`).join("");
  },
  showMode(mode) {
    if (!AppState.map) return;
    AppState.mapMode = mode;
    Object.values(AppState.mapLayers).forEach(layer => { if (layer && AppState.map.hasLayer(layer)) AppState.map.removeLayer(layer); });
    AppState.mapLayers[mode]?.addTo(AppState.map);
    $$("[data-map-mode]").forEach(button => button.classList.toggle("active", button.dataset.mapMode === mode));
  },
  centroid(coords) {
    return [coords.reduce((sum, item) => sum + item[0], 0) / coords.length, coords.reduce((sum, item) => sum + item[1], 0) / coords.length];
  }
};

function switchTab(tab) {
  AppState.activeTab = tab;
  $$(".tab-page").forEach(page => page.classList.toggle("active", page.id === `tab-${tab}`));
  $$("[data-tab]").forEach(button => {
    const active = button.dataset.tab === tab;
    button.classList.toggle("active", active);
    button.setAttribute("aria-current", active ? "page" : "false");
  });
  renderActiveTab();
  $("#main")?.focus({ preventScroll: true });
  document.body.classList.remove("sidebar-open");
}
function renderActiveTab() {
  PageEngines[AppState.activeTab]?.();
  requestAnimationFrame(() => ChartEngine.resizeAll());
}
function setLoader(percent, message) {
  $("#loader-fill").style.width = `${percent}%`;
  $("#loader-status").textContent = message;
}
function setSyncState(state) {
  $("#sync-dot").className = `status-dot ${state === "syncing" ? "syncing" : state === "error" ? "error" : ""}`;
}
function updateTimestamp() {
  $("#last-updated-txt").textContent = AppState.lastRefresh ? `Updated ${AppState.lastRefresh.toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit" })}` : "Not updated";
}
async function refreshData(showLoader = false) {
  setSyncState("syncing");
  if (showLoader) setLoader(10, "Connecting to Google Sheets...");
  try {
    const raw = await DataEngine.fetchAll((percent, message) => showLoader && setLoader(percent, message));
    AppState.raw.farmers = raw.farmers || [];
    AppState.raw.payments = raw.payments || [];
    AppState.raw.delivery = raw.delivery || [];
    if (showLoader) setLoader(88, "Preparing filters...");
    FilterEngine.populate();
    FilterEngine.apply();
    if (showLoader) setLoader(96, "Rendering intelligence workspace...");
    renderActiveTab();
    AppState.lastRefresh = new Date();
    updateTimestamp();
    setSyncState("live");
    if (showLoader) { setLoader(100, "Ready"); setTimeout(() => $("#loading-overlay")?.classList.add("hidden"), 250); }
  } catch (error) {
    console.error(error);
    setSyncState("error");
    if (showLoader) { setLoader(100, "Connection failed. Check API key and spreadsheet access."); setTimeout(() => $("#loading-overlay")?.classList.add("hidden"), 1600); }
  }
}
function exportCSV() {
  const rows = AppState.filtered.farmers;
  if (!rows.length) return alert("No farmer data to export.");
  const headers = Object.keys(rows[0]);
  const csv = [headers, ...rows.map(row => headers.map(header => `"${String(row[header] || "").replace(/"/g, '""')}"`))].map(row => row.join(",")).join("\n");
  download("herveg05-farmers.csv", "text/csv", csv);
}
function exportExcel() {
  if (!window.XLSX) return alert("Excel library is still loading. Try again in a moment.");
  const workbook = XLSX.utils.book_new();
  Object.entries(AppState.filtered).forEach(([name, rows]) => { if (rows.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), name); });
  XLSX.writeFile(workbook, "herveg05-intelligence-data.xlsx");
}
async function exportPNG() {
  if (!window.html2canvas) await import("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");
  const canvas = await window.html2canvas(document.getElementById("main"), { backgroundColor: Utils.cssVar("--bg") });
  download("herveg05-dashboard.png", "image/png", canvas.toDataURL("image/png"));
}
function download(name, type, content) {
  const anchor = document.createElement("a");
  anchor.href = type === "image/png" ? content : URL.createObjectURL(new Blob([content], { type }));
  anchor.download = name;
  anchor.click();
  if (type !== "image/png") URL.revokeObjectURL(anchor.href);
}
function debounce(fn, delay) {
  let timer;
  return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), delay); };
}
function bindUI() {
  $$("[data-tab]").forEach(button => button.addEventListener("click", () => switchTab(button.dataset.tab)));
  $$(".filter-grid select").forEach(select => select.addEventListener("change", () => { FilterEngine.apply(); renderActiveTab(); }));
  $("#clear-filters-btn").addEventListener("click", () => { FilterEngine.clear(); renderActiveTab(); });
  $("#refresh-btn").addEventListener("click", () => refreshData(false));
  $("#mobile-menu").addEventListener("click", () => { const open = !document.body.classList.contains("sidebar-open"); document.body.classList.toggle("sidebar-open", open); $("#mobile-menu").setAttribute("aria-expanded", String(open)); });
  $("#filter-toggle").addEventListener("click", () => { const panel = $("#filter-panel"); const open = !panel.classList.contains("open"); panel.classList.toggle("open", open); document.body.classList.toggle("filters-collapsed", !open); $("#filter-toggle").setAttribute("aria-expanded", String(open)); setTimeout(() => { ChartEngine.resizeAll(); AppState.map?.invalidateSize(); }, 220); });
  $("#theme-btn").addEventListener("click", () => { const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark"; document.documentElement.dataset.theme = next; localStorage.setItem("hvg-theme", next); $("#theme-btn").textContent = next === "dark" ? "Light" : "Dark"; ChartEngine.rerenderTheme(); });
  $("#export-btn").addEventListener("click", event => { event.stopPropagation(); const menu = $("#export-menu"); const willOpen = menu.hasAttribute("hidden"); menu.toggleAttribute("hidden", !willOpen); $("#export-btn").setAttribute("aria-expanded", String(willOpen)); });
  $("#export-menu").addEventListener("click", event => { const action = event.target?.dataset?.export; if (!action) return; $("#export-menu").setAttribute("hidden", ""); $("#export-btn").setAttribute("aria-expanded", "false"); if (action === "csv") exportCSV(); if (action === "excel") exportExcel(); if (action === "png") exportPNG(); });
  document.addEventListener("click", event => { if (!$("#export-menu").contains(event.target) && event.target !== $("#export-btn")) { $("#export-menu").setAttribute("hidden", ""); $("#export-btn").setAttribute("aria-expanded", "false"); } });
  $$("[data-map-mode]").forEach(button => button.addEventListener("click", () => GeoEngine.showMode(button.dataset.mapMode)));
  $("#global-search").addEventListener("input", debounce(event => {
    const query = event.target.value.trim().toLowerCase();
    if (!query) { FilterEngine.apply(); renderActiveTab(); return; }
    FilterEngine.apply();
    AppState.filtered.farmers = AppState.filtered.farmers.filter(row => JSON.stringify(row).toLowerCase().includes(query));
    AppState.filtered.payments = AppState.filtered.payments.filter(row => JSON.stringify(row).toLowerCase().includes(query));
    AppState.filtered.delivery = AppState.filtered.delivery.filter(row => JSON.stringify(row).toLowerCase().includes(query));
    renderActiveTab();
  }, 180));
  window.addEventListener("resize", debounce(() => { ChartEngine.resizeAll(); AppState.map?.invalidateSize(); }, 180));
  document.addEventListener("keydown", event => {
    if (event.target.matches("input, textarea, select")) return;
    const map = { "1": "overview", "2": "enrollment", "3": "payment", "4": "delivery", "5": "geo", "6": "pipeline", "7": "strategy" };
    if (map[event.key]) switchTab(map[event.key]);
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "r") { event.preventDefault(); refreshData(false); }
  });
}
function initTheme() {
  const saved = localStorage.getItem("hvg-theme");
  const prefersDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
  const theme = saved || (prefersDark ? "dark" : "light");
  document.documentElement.dataset.theme = theme;
  $("#theme-btn").textContent = theme === "dark" ? "Light" : "Dark";
}
document.addEventListener("DOMContentLoaded", async () => {
  initTheme();
  bindUI();
  await refreshData(true);
  setInterval(() => refreshData(false), CONFIG.refreshInterval);
});
