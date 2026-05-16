"use strict";

const CONFIG = {
  apiKey: "AIzaSyCjceOafhRMvv2A5nAp0zzJrf8f40lpQrM",
  spreadsheetId: "1wrqKSbH21jvwmesMdCVymsFSQD-P1OHU4iINq1uYdWo",
  refreshInterval: 10 * 60 * 1000,
  sheets: {
    farmers: "Farmers",
    payments: "Farmers Payments",
    delivery: "Farmers Delivery"
  }
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
      headers.forEach((header, index) => {
        record[header] = row[index] !== undefined ? String(row[index]).trim() : "";
      });
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
    const packages = [
      ...Utils.unique(AppState.raw.farmers, "PACKAGE"),
      ...Utils.unique(AppState.raw.payments, "Package Name"),
      ...Utils.unique(AppState.raw.delivery, "Package Delivered")
    ];
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
    return ["--chart-green", "--chart-green-2", "--chart-amber", "--chart-blue", "--chart-red", "--chart-gray"].map(Utils.cssVar);
  },
  base() {
    return {
      animation: !window.matchMedia("(prefers-reduced-motion: reduce)").matches,
      animationDuration: 350,
      textStyle: { fontFamily: "Inter, system-ui, sans-serif", color: Utils.cssVar("--text") },
      color: this.palette(),
      grid: { top: 24, right: 20, bottom: 36, left: 46, containLabel: true },
      tooltip: {
        trigger: "item",
        backgroundColor: Utils.cssVar("--surface"),
        borderColor: Utils.cssVar("--border"),
        borderWidth: 1,
        textStyle: { color: Utils.cssVar("--text"), fontSize: 12 },
        extraCssText: "box-shadow:0 10px 28px rgba(0,0,0,.12);border-radius:8px;padding:10px 12px;"
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
    return {
      type: "category",
      data,
      axisTick: { show: false },
      axisLine: { lineStyle: { color: Utils.cssVar("--border") } },
      axisLabel: { color: Utils.cssVar("--text-3"), fontSize: 11, rotate: !vertical && data.length > 6 ? 24 : 0 }
    };
  },
  yAxis() {
    return {
      type: "value",
      axisLabel: { color: Utils.cssVar("--text-3"), fontSize: 11 },
      splitLine: { lineStyle: { color: Utils.cssVar("--chart-grid") } }
    };
  },
  legend() {
    return { bottom: 0, textStyle: { color: Utils.cssVar("--text-2"), fontSize: 11 } };
  },
  line(id, categories, series) {
    const palette = this.palette();
    this.set(id, {
      ...this.base(),
      tooltip: { ...this.base().tooltip, trigger: "axis" },
      xAxis: this.xAxis(categories),
      yAxis: this.yAxis(),
      legend: series.length > 1 ? this.legend() : { show: false },
      series: series.map((item, index) => ({
        type: "line",
        name: item.name,
        data: item.data,
        smooth: true,
        symbol: "circle",
        symbolSize: 5,
        lineStyle: { width: 2, color: palette[index % palette.length] },
        itemStyle: { color: palette[index % palette.length] },
        areaStyle: index === 0 ? { opacity: 0.08 } : undefined
      }))
    });
  },
  bar(id, categories, series, horizontal = false) {
    const palette = this.palette();
    this.set(id, {
      ...this.base(),
      tooltip: { ...this.base().tooltip, trigger: "axis" },
      xAxis: horizontal ? this.yAxis() : this.xAxis(categories),
      yAxis: horizontal ? this.xAxis(categories, true) : this.yAxis(),
      legend: series.length > 1 ? this.legend() : { show: false },
      series: series.map((item, index) => ({
        type: "bar",
        name: item.name,
        data: item.data,
        barMaxWidth: 34,
        itemStyle: { color: item.color || palette[index % palette.length], borderRadius: horizontal ? [0, 4, 4, 0] : [4, 4, 0, 0] }
      }))
    });
  },
  donut(id, data) {
    this.set(id, {
      ...this.base(),
      tooltip: { ...this.base().tooltip, formatter: "{b}<br>{c} ({d}%)" },
      legend: { orient: "vertical", right: 0, top: "center", textStyle: { color: Utils.cssVar("--text-2"), fontSize: 11 } },
      series: [{ type: "pie", radius: ["52%", "74%"], center: ["38%", "50%"], label: { show: false }, labelLine: { show: false }, data }]
    });
  },
  radar(id, indicators, values) {
    this.set(id, {
      ...this.base(),
      radar: {
        indicator: indicators.map(name => ({ name, max: 100 })),
        axisName: { color: Utils.cssVar("--text-2"), fontSize: 11 },
        splitLine: { lineStyle: { color: Utils.cssVar("--chart-grid") } },
        splitArea: { areaStyle: { color: ["transparent", "rgba(40,112,71,.04)"] } }
      },
      series: [{ type: "radar", data: [{ value: values, areaStyle: { opacity: 0.14 }, lineStyle: { width: 2 }, itemStyle: { color: Utils.cssVar("--chart-green") } }] }]
    });
  },
  scatter(id, data) {
    this.set(id, {
      ...this.base(),
      tooltip: { ...this.base().tooltip, formatter: point => `${point.name}<br>Farmers: ${point.value[0]}<br>Collection: ${point.value[1]}%` },
      xAxis: { type: "value", name: "Farmers", axisLabel: { color: Utils.cssVar("--text-3") }, splitLine: { lineStyle: { color: Utils.cssVar("--chart-grid") } } },
      yAxis: { type: "value", name: "Collection %", axisLabel: { color: Utils.cssVar("--text-3"), formatter: "{value}%" }, splitLine: { lineStyle: { color: Utils.cssVar("--chart-grid") } } },
      series: [{ type: "scatter", data, symbolSize: value => Math.max(Math.sqrt(value[2]) * 4, 12), itemStyle: { opacity: 0.82 }, label: { show: true, formatter: point => point.name, position: "right", color: Utils.cssVar("--text-2"), fontSize: 10 } }]
    });
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
        <div>
          <p class="metric-label">${Utils.esc(card.label)}</p>
          <div class="metric-value">${Utils.esc(card.value)}</div>
        </div>
        <div>
          ${card.trend !== undefined ? `<span class="metric-trend ${card.trend > 0 ? "up" : card.trend < 0 ? "down" : "neutral"}">${card.trend > 0 ? "+" : ""}${card.trend.toFixed(1)}%</span>` : ""}
          <p class="metric-note">${Utils.esc(card.note || "")}</p>
        </div>
      </article>
    `).join("");
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
  const avgNSP = Utils.avgBy(d, "NSP Score");
  const endToEnd = enrolled ? (delivered / enrolled) * 100 : 0;
  const enrollToPay = enrolled ? (paid / enrolled) * 100 : 0;
  const payToDeliver = paid ? (delivered / paid) * 100 : 0;
  const byMonth = {};
  f.forEach(row => {
    const key = Utils.monthKey(row["DATE CREATED"]);
    if (key !== "Unknown") byMonth[key] = (byMonth[key] || 0) + 1;
  });
  const months = Object.keys(byMonth).sort();
  let growth = 0;
  if (months.length >= 2) {
    const current = byMonth[months.at(-1)];
    const previous = byMonth[months.at(-2)];
    growth = previous ? ((current - previous) / previous) * 100 : 0;
  }
  const health = Math.round((
    Math.min(100, Math.max(0, growth + 60)) +
    Math.min(100, collectionRate) +
    Math.min(100, endToEnd) +
    Math.min(100, femaleRatio * 2) +
    (avgNSP ? Math.min(100, avgNSP) : 55)
  ) / 5);
  return { f, p, d, enrolled, paid, delivered, revenue, cost, collectionRate, femaleRatio, villages, zones, householdReach, avgDelivery, avgNSP, endToEnd, enrollToPay, payToDeliver, growth, health, byMonth, months };
}

const Insight = {
  card(title, text) {
    return `<article class="insight-card"><h3>${Utils.esc(title)}</h3><p>${Utils.esc(text)}</p></article>`;
  },
  finding(title, text, tone = "positive") {
    return `<div class="finding ${tone}"><h3>${Utils.esc(title)}</h3><p>${Utils.esc(text)}</p></div>`;
  },
  recommendation(title, text, tone = "low") {
    return `<div class="recommendation ${tone}"><h3>${Utils.esc(title)}</h3><p>${Utils.esc(text)}</p></div>`;
  }
};

const PageEngines = {
  overview() {
    const m = metrics();
    $("#overview-health-score").textContent = String(m.health);
    $("#overview-health-label").textContent = m.health >= 75 ? "Strong" : m.health >= 55 ? "Watch" : "Needs action";
    $("#overview-narrative").textContent = `The program has enrolled ${Utils.fmtNum(m.enrolled)} farmers across ${m.zones} zones and ${m.villages} villages. Collection is ${Utils.fmtPct(m.collectionRate)}, delivery completion is ${Utils.fmtPct(m.endToEnd)}, and female participation is ${Utils.fmtPct(m.femaleRatio)}.`;
    KPIEngine.render("overview-kpis", [
      { label: "Farmers Enrolled", value: Utils.fmtNum(m.enrolled), trend: m.growth, note: "Registration momentum across active villages." },
      { label: "Revenue Collected", value: Utils.fmtCurrency(m.revenue), note: `${Utils.fmtPct(m.collectionRate)} of expected package value.` },
      { label: "Delivered Farmers", value: Utils.fmtNum(m.delivered), note: `${Utils.fmtPct(m.endToEnd)} end-to-end completion.` },
      { label: "Household Reach", value: Utils.fmtNum(m.householdReach), note: "Estimated indirect reach through farmer households." }
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
      Insight.card("Primary action", m.collectionRate < 75 ? "Prioritize collection follow-up in low-performing zones before scaling enrollment." : "Collection health is stable enough to support continued field expansion."),
      Insight.card("Operational signal", m.avgDelivery > 45 ? `Average delivery time is ${m.avgDelivery.toFixed(0)} days, suggesting a logistics bottleneck.` : "Delivery timelines are within a manageable operating band."),
      Insight.card("Inclusion signal", m.femaleRatio >= 50 ? "Female participation is at or above parity, supporting donor and impact reporting." : "Female participation is below parity and needs targeted outreach."),
      Insight.card("Scale signal", `${m.villages} villages are active. Top villages can be used as referral hubs for adjacent expansion.`)
    ].join("");
  },
  enrollment() {
    const m = metrics();
    const groupRate = m.enrolled ? (m.f.filter(row => ["yes", "true", "1"].includes(String(row["BELONGS IN A GROUP?"]).toLowerCase())).length / m.enrolled) * 100 : 0;
    const under5 = Utils.sumBy(m.f, "CHILDREN UNDER 5 YEARS OLD");
    KPIEngine.render("enr-kpis", [
      { label: "Total Enrolled", value: Utils.fmtNum(m.enrolled), trend: m.growth, note: "Program intake volume." },
      { label: "Female Participation", value: Utils.fmtPct(m.femaleRatio), note: "Gender inclusion signal." },
      { label: "Group Participation", value: Utils.fmtPct(groupRate), note: "Community network strength." },
      { label: "Children Under 5", value: Utils.fmtNum(under5), note: "Nutrition-sensitive reach." }
    ]);
    ChartEngine.line("enr-velocity", m.months.map(Utils.monthLabel), [{ name: "Enrollments", data: m.months.map(key => m.byMonth[key]) }]);
    const vano = Utils.entriesSorted(Utils.countBy(m.f, "VANO"), 12);
    ChartEngine.bar("enr-vano", vano.map(([key]) => key), [{ name: "Farmers", data: vano.map(([, value]) => value) }], true);
    const packages = Utils.entriesSorted(Utils.countBy(m.f, "PACKAGE"), 10);
    ChartEngine.bar("enr-package", packages.map(([key]) => key), [{ name: "Farmers", data: packages.map(([, value]) => value) }]);
    ChartEngine.radar("enr-equity-radar", ["Gender", "Groups", "Zones", "Villages", "Growth"], [Math.min(100, m.femaleRatio * 2), groupRate, Math.min(100, m.zones * 10), Math.min(100, m.villages * 2), Math.min(100, Math.max(0, m.growth + 50))]);
    $("#enr-insights").innerHTML = [
      Insight.card("Growth", m.growth > 0 ? `Enrollment is growing at ${m.growth.toFixed(1)}% month over month.` : "Enrollment is flat or declining; review VANO outreach targets."),
      Insight.card("Field performance", vano[0] ? `${vano[0][0]} leads field intake with ${vano[0][1]} farmers.` : "No VANO performance data is available."),
      Insight.card("Package demand", packages[0] ? `${packages[0][0]} is the strongest package signal.` : "Package demand data is limited."),
      Insight.card("Inclusion", m.femaleRatio >= 50 ? "Gender parity is healthy." : "Women-focused enrollment should be strengthened.")
    ].join("");
  },
  payment() {
    const m = metrics();
    const outstanding = Math.max(0, m.cost - m.revenue);
    const fullPaid = m.p.filter(row => Utils.num(row["% of Package Cost Paid"]) >= 100).length;
    KPIEngine.render("pay-kpis", [
      { label: "Revenue Collected", value: Utils.fmtCurrency(m.revenue), note: "Actual paid across all payment records." },
      { label: "Collection Rate", value: Utils.fmtPct(m.collectionRate), note: "Actual paid divided by package cost." },
      { label: "Outstanding", value: Utils.fmtCurrency(outstanding), note: "Remaining balance to collect." },
      { label: "Fully Paid Farmers", value: Utils.fmtNum(fullPaid), note: "Farmers at or above 100% paid." }
    ]);
    const zoneRevenue = {};
    m.p.forEach(row => { zoneRevenue[row.Zone || "Unknown"] = (zoneRevenue[row.Zone || "Unknown"] || 0) + Utils.num(row["Actual Paid (TZS)"]); });
    const revenueEntries = Utils.entriesSorted(zoneRevenue, 10);
    ChartEngine.bar("pay-zone-revenue", revenueEntries.map(([key]) => key), [{ name: "Revenue", data: revenueEntries.map(([, value]) => Math.round(value)) }]);
    const buckets = { "0%": 0, "1-50%": 0, "51-99%": 0, "100%+": 0 };
    m.p.forEach(row => {
      const value = Utils.num(row["% of Package Cost Paid"]);
      if (value === 0) buckets["0%"]++;
      else if (value <= 50) buckets["1-50%"]++;
      else if (value < 100) buckets["51-99%"]++;
      else buckets["100%+"]++;
    });
    ChartEngine.donut("pay-collection-dist", Object.entries(buckets).map(([name, value]) => ({ name, value })));
    const agents = {};
    m.p.forEach(row => { agents[row["Village Agent"] || "Unknown"] = (agents[row["Village Agent"] || "Unknown"] || 0) + Utils.num(row["Actual Paid (TZS)"]); });
    const agentEntries = Utils.entriesSorted(agents, 12);
    ChartEngine.bar("pay-agent-leaderboard", agentEntries.map(([key]) => key), [{ name: "Revenue", data: agentEntries.map(([, value]) => Math.round(value)) }], true);
    const byMonth = {};
    m.p.forEach(row => {
      const key = Utils.monthKey(row["Date of Finishing"]);
      if (key !== "Unknown") byMonth[key] = (byMonth[key] || 0) + Utils.num(row["Actual Paid (TZS)"]);
    });
    const months = Object.keys(byMonth).sort();
    let running = 0;
    ChartEngine.line("pay-cumulative", months.map(Utils.monthLabel), [{ name: "Revenue", data: months.map(key => Math.round(running += byMonth[key])) }]);
    $("#pay-insights").innerHTML = [
      Insight.card("Collection priority", outstanding > 0 ? `${Utils.fmtCurrency(outstanding)} remains outstanding.` : "Expected package value has been collected."),
      Insight.card("Agent accountability", agentEntries[0] ? `${agentEntries[0][0]} leads collections with ${Utils.fmtCurrency(agentEntries[0][1])}.` : "No agent collection signal available."),
      Insight.card("Payment quality", `${fullPaid} farmers have fully paid or overpaid.`),
      Insight.card("Revenue health", m.collectionRate >= 80 ? "Collection health is strong." : "Collection rate is below executive target.")
    ].join("");
  },
  delivery() {
    const m = metrics();
    const femaleDelivered = m.d.filter(row => row.Gender === "Female").length;
    const femaleDeliveryRate = m.delivered ? (femaleDelivered / m.delivered) * 100 : 0;
    KPIEngine.render("del-kpis", [
      { label: "Farmers Delivered", value: Utils.fmtNum(m.delivered), note: "Completed fulfillment count." },
      { label: "Avg Delivery Time", value: m.avgDelivery ? `${m.avgDelivery.toFixed(0)}d` : "-", note: "Enrollment to delivery." },
      { label: "Avg NSP Score", value: m.avgNSP ? m.avgNSP.toFixed(1) : "-", note: "Farmer satisfaction quality." },
      { label: "Women Delivered", value: Utils.fmtPct(femaleDeliveryRate), note: "Gender inclusion in fulfillment." }
    ]);
    const byMonth = {};
    m.d.forEach(row => {
      const key = Utils.monthKey(row["Delivery Date"]);
      if (key !== "Unknown") byMonth[key] = (byMonth[key] || 0) + 1;
    });
    const months = Object.keys(byMonth).sort();
    ChartEngine.line("del-growth", months.map(Utils.monthLabel), [{ name: "Deliveries", data: months.map(key => byMonth[key]) }]);
    const zones = Utils.entriesSorted(Utils.countBy(m.d, "Zone"), 10);
    ChartEngine.bar("del-zone-coverage", zones.map(([key]) => key), [{ name: "Deliveries", data: zones.map(([, value]) => value) }]);
    const buckets = { "0-14d": 0, "15-30d": 0, "31-60d": 0, "60d+": 0 };
    m.d.forEach(row => {
      const value = Utils.num(row["Enrollment to delivery Days"]);
      if (value <= 14) buckets["0-14d"]++;
      else if (value <= 30) buckets["15-30d"]++;
      else if (value <= 60) buckets["31-60d"]++;
      else buckets["60d+"]++;
    });
    ChartEngine.bar("del-timeline", Object.keys(buckets), [{ name: "Farmers", data: Object.values(buckets) }]);
    const nspZones = Utils.unique(m.d, "Zone").filter(Boolean);
    ChartEngine.bar("del-nsp-dist", nspZones, [{ name: "NSP", data: nspZones.map(zone => Number(Utils.avgBy(m.d.filter(row => row.Zone === zone), "NSP Score").toFixed(1))) }]);
    $("#del-insights").innerHTML = [
      Insight.card("Fulfillment", `${Utils.fmtNum(m.delivered)} farmers have received packages.`),
      Insight.card("Speed", m.avgDelivery > 45 ? "Delivery speed is above target and requires logistics review." : "Delivery speed is within a manageable range."),
      Insight.card("Experience", m.avgNSP ? `Average NSP score is ${m.avgNSP.toFixed(1)}.` : "NSP satisfaction data is limited."),
      Insight.card("Equity", `${Utils.fmtPct(femaleDeliveryRate)} of delivered farmers are women.`)
    ].join("");
  },
  geo() {
    const m = metrics();
    const mapped = m.f.filter(GeoEngine.validCoord);
    KPIEngine.render("geo-kpis", [
      { label: "Farmers Mapped", value: Utils.fmtNum(mapped.length), note: "Records with valid coordinates." },
      { label: "GPS Coverage", value: Utils.fmtPct(m.enrolled ? mapped.length / m.enrolled * 100 : 0), note: "Mapping data quality." },
      { label: "Zones Covered", value: Utils.fmtNum(m.zones), note: "Operational territories." },
      { label: "Villages Reached", value: Utils.fmtNum(m.villages), note: "Rural footprint." }
    ]);
    GeoEngine.renderMap(m.f, mapped);
    const zones = Utils.entriesSorted(Utils.countBy(m.f, "Zone"), 10);
    ChartEngine.bar("geo-zone-penetration", zones.map(([key]) => key), [{ name: "Farmers", data: zones.map(([, value]) => value) }]);
    const villages = Utils.entriesSorted(Utils.countBy(m.f, "Village"), 15);
    ChartEngine.bar("geo-village-density", villages.map(([key]) => key), [{ name: "Farmers", data: villages.map(([, value]) => value) }], true);
    $("#geo-insights").innerHTML = [
      Insight.card("Territory signal", `${m.zones} zones and ${m.villages} villages are active.`),
      Insight.card("Data quality", `${Utils.fmtPct(m.enrolled ? mapped.length / m.enrolled * 100 : 0)} of farmers have usable GPS coordinates.`),
      Insight.card("Expansion", villages[0] ? `${villages[0][0]} is the highest-density village and can anchor adjacent outreach.` : "Village density data is limited."),
      Insight.card("Resource allocation", zones[0] ? `${zones[0][0]} is the strongest enrollment territory.` : "Zone signal is limited.")
    ].join("");
  },
  pipeline() {
    const m = metrics();
    const dropEnrollPay = Math.max(0, m.enrolled - m.paid);
    const dropPayDeliver = Math.max(0, m.paid - m.delivered);
    KPIEngine.render("pipe-kpis", [
      { label: "Enroll to Pay", value: Utils.fmtPct(m.enrollToPay), note: "Payment conversion." },
      { label: "Pay to Deliver", value: Utils.fmtPct(m.payToDeliver), note: "Fulfillment conversion." },
      { label: "End-to-End", value: Utils.fmtPct(m.endToEnd), note: "Completed journey." },
      { label: "Pipeline Leakage", value: Utils.fmtNum(dropEnrollPay + dropPayDeliver), note: "Farmers needing follow-up." }
    ]);
    $("#pipe-funnel-stages").innerHTML = [
      { title: "Enrolled", count: m.enrolled, pct: 100, note: "Farmers registered in the program." },
      { title: "Made Payment", count: m.paid, pct: m.enrollToPay, note: `${dropEnrollPay} farmers have not converted to payment.` },
      { title: "Delivered", count: m.delivered, pct: m.endToEnd, note: `${dropPayDeliver} paying farmers still need fulfillment.` }
    ].map(stage => `
      <div class="pipeline-stage">
        <div><h3>${stage.title}</h3><p>${stage.note}</p></div>
        <strong>${Utils.fmtNum(stage.count)}</strong>
        <div class="pipeline-bar"><span style="--w:${Math.max(0, Math.min(stage.pct, 100))}%"></span></div>
      </div>
    `).join("");
    const zones = Utils.unique(m.f, "Zone").filter(Boolean);
    ChartEngine.bar("pipe-zone-conversion", zones, [{ name: "Enroll to Pay", data: zones.map(zone => {
      const enrolled = m.f.filter(row => row.Zone === zone).length;
      const paid = m.p.filter(row => row.Zone === zone).length;
      return enrolled ? Number((paid / enrolled * 100).toFixed(1)) : 0;
    }) }]);
    ChartEngine.bar("pipe-delivery-conversion", zones, [{ name: "Pay to Deliver", data: zones.map(zone => {
      const paid = m.p.filter(row => row.Zone === zone).length;
      const delivered = m.d.filter(row => row.Zone === zone).length;
      return paid ? Number((delivered / paid * 100).toFixed(1)) : 0;
    }) }]);
    $("#pipe-insights").innerHTML = [
      Insight.card("Conversion", `${Utils.fmtPct(m.endToEnd)} of enrolled farmers have completed the journey.`),
      Insight.card("Payment gap", `${Utils.fmtNum(dropEnrollPay)} enrolled farmers need payment conversion.`),
      Insight.card("Delivery gap", `${Utils.fmtNum(dropPayDeliver)} paying farmers need fulfillment follow-up.`),
      Insight.card("Executive focus", m.enrollToPay < m.payToDeliver ? "Payment conversion is the larger bottleneck." : "Fulfillment is the larger bottleneck.")
    ].join("");
  },
  strategy() {
    const m = metrics();
    const outstanding = Math.max(0, m.cost - m.revenue);
    $("#strategy-exec-text").textContent = `HERVeg.05 currently shows a program health score of ${m.health}/100. Enrollment stands at ${Utils.fmtNum(m.enrolled)}, revenue collection is ${Utils.fmtPct(m.collectionRate)}, delivery completion is ${Utils.fmtPct(m.endToEnd)}, and female participation is ${Utils.fmtPct(m.femaleRatio)}. Leadership should focus on the largest conversion constraint before expanding field activity.`;
    $("#strategy-trends").innerHTML = [
      Insight.finding("Enrollment momentum", m.growth > 0 ? `Enrollment is growing at ${m.growth.toFixed(1)}% month over month.` : "Enrollment is flat or declining.", m.growth > 0 ? "positive" : "warning"),
      Insight.finding("Collection performance", `${Utils.fmtPct(m.collectionRate)} of expected package value has been collected.`, m.collectionRate >= 80 ? "positive" : "warning"),
      Insight.finding("Delivery conversion", `${Utils.fmtPct(m.endToEnd)} of enrolled farmers have received delivery.`, m.endToEnd >= 70 ? "positive" : "warning")
    ].join("");
    $("#strategy-risks").innerHTML = [
      outstanding > 0 ? Insight.finding("Outstanding balance", `${Utils.fmtCurrency(outstanding)} remains unpaid.`, outstanding > m.revenue * 0.25 ? "critical" : "warning") : Insight.finding("No outstanding balance", "Expected package value has been collected.", "positive"),
      m.femaleRatio < 50 ? Insight.finding("Gender parity gap", `Female participation is ${Utils.fmtPct(m.femaleRatio)}.`, "warning") : Insight.finding("Gender parity", "Female participation is at or above parity.", "positive"),
      m.avgDelivery > 45 ? Insight.finding("Delivery delay", `Average delivery time is ${m.avgDelivery.toFixed(0)} days.`, "critical") : Insight.finding("Delivery timing", "Delivery time is within the operating band.", "positive")
    ].join("");
    ChartEngine.radar("strategy-health-radar", ["Growth", "Collection", "Delivery", "Gender", "Satisfaction"], [Math.min(100, Math.max(0, m.growth + 55)), Math.min(100, m.collectionRate), Math.min(100, m.endToEnd), Math.min(100, m.femaleRatio * 2), m.avgNSP ? Math.min(100, m.avgNSP) : 55]);
    const matrix = Utils.unique(m.f, "Zone").filter(Boolean).map(zone => {
      const farmers = m.f.filter(row => row.Zone === zone).length;
      const paidRows = m.p.filter(row => row.Zone === zone);
      const revenue = Utils.sumBy(paidRows, "Actual Paid (TZS)");
      const cost = Utils.sumBy(paidRows, "Package Cost (TZS)");
      const rate = cost ? Number((revenue / cost * 100).toFixed(1)) : 0;
      return { name: zone, value: [farmers, rate, farmers] };
    });
    ChartEngine.scatter("strategy-opportunity", matrix);
    $("#strategy-recommendations").innerHTML = [
      m.collectionRate < 80 ? Insight.recommendation("Accelerate collections", `Recover ${Utils.fmtCurrency(outstanding)} through VANO follow-up in low-collection zones.`, "high") : Insight.recommendation("Maintain collection discipline", "Collection rate is healthy; keep monitoring partial payers.", "low"),
      m.endToEnd < 70 ? Insight.recommendation("Close fulfillment gap", "Prioritize paid farmers who have not yet received delivery.", "high") : Insight.recommendation("Scale operating cadence", "Fulfillment is strong enough to support controlled expansion.", "low"),
      m.femaleRatio < 50 ? Insight.recommendation("Strengthen women outreach", "Launch targeted enrollment through women farmer groups and village leaders.", "medium") : Insight.recommendation("Protect inclusion gains", "Document women outreach practices for replication.", "low")
    ].join("");
    $("#strategy-opportunities").innerHTML = [
      Insight.recommendation("Replicate strongest villages", "Use high-density villages as referral hubs for neighboring communities.", "low"),
      Insight.recommendation("Coach underperforming agents", "Benchmark top VANO behavior and convert it into field coaching routines.", "medium"),
      Insight.recommendation("Improve data quality", "Standardize village, zone, and GPS capture to improve geographic targeting.", "medium")
    ].join("");
  }
};

const GeoEngine = {
  validCoord(row) {
    const lat = parseFloat(row.LATITUDE || row.Latitude || "");
    const lng = parseFloat(row.LONGITUDE || row.Longitude || "");
    return !Number.isNaN(lat) && !Number.isNaN(lng) && lat !== 0 && lng !== 0 && Math.abs(lat) <= 90 && Math.abs(lng) <= 180;
  },
  coord(row) {
    return [parseFloat(row.LATITUDE || row.Latitude), parseFloat(row.LONGITUDE || row.Longitude)];
  },
  renderMap(allRows, mappedRows) {
    if (!window.L) return;
    const element = document.getElementById("geo-map");
    if (!element) return;
    if (AppState.map) {
      AppState.map.remove();
      AppState.map = null;
    }
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
    Object.values(AppState.mapLayers).forEach(layer => {
      if (layer && AppState.map.hasLayer(layer)) AppState.map.removeLayer(layer);
    });
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
}

function renderActiveTab() {
  const engine = PageEngines[AppState.activeTab];
  if (engine) engine();
  requestAnimationFrame(() => ChartEngine.resizeAll());
}

function setLoader(percent, message) {
  $("#loader-fill").style.width = `${percent}%`;
  $("#loader-status").textContent = message;
}

function setSyncState(state) {
  const dot = $("#sync-dot");
  dot.className = `status-dot ${state === "syncing" ? "syncing" : state === "error" ? "error" : ""}`;
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
    if (showLoader) {
      setLoader(100, "Ready");
      setTimeout(() => $("#loading-overlay")?.classList.add("hidden"), 250);
    }
  } catch (error) {
    console.error(error);
    setSyncState("error");
    if (showLoader) {
      setLoader(100, "Connection failed. Check API key and spreadsheet access.");
      setTimeout(() => $("#loading-overlay")?.classList.add("hidden"), 1600);
    }
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
  Object.entries(AppState.filtered).forEach(([name, rows]) => {
    if (rows.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), name);
  });
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
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
}

function bindUI() {
  $$("[data-tab]").forEach(button => button.addEventListener("click", () => switchTab(button.dataset.tab)));
  $$(".filter-grid select").forEach(select => select.addEventListener("change", () => {
    FilterEngine.apply();
    renderActiveTab();
  }));
  $("#clear-filters-btn").addEventListener("click", () => {
    FilterEngine.clear();
    renderActiveTab();
  });
  $("#refresh-btn").addEventListener("click", () => refreshData(false));
  $("#filter-toggle").addEventListener("click", () => {
    const panel = $("#filter-panel");
    const open = !panel.classList.contains("open");
    panel.classList.toggle("open", open);
    document.body.classList.toggle("filters-collapsed", !open);
    $("#filter-toggle").setAttribute("aria-expanded", String(open));
    setTimeout(() => {
      ChartEngine.resizeAll();
      AppState.map?.invalidateSize();
    }, 220);
  });
  $("#theme-btn").addEventListener("click", () => {
    const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
    document.documentElement.dataset.theme = next;
    localStorage.setItem("hvg-theme", next);
    $("#theme-btn").textContent = next === "dark" ? "Light" : "Dark";
    ChartEngine.rerenderTheme();
  });
  $("#export-btn").addEventListener("click", event => {
    event.stopPropagation();
    const menu = $("#export-menu");
    const willOpen = menu.hasAttribute("hidden");
    menu.toggleAttribute("hidden", !willOpen);
    $("#export-btn").setAttribute("aria-expanded", String(willOpen));
  });
  $("#export-menu").addEventListener("click", event => {
    const action = event.target?.dataset?.export;
    if (!action) return;
    $("#export-menu").setAttribute("hidden", "");
    $("#export-btn").setAttribute("aria-expanded", "false");
    if (action === "csv") exportCSV();
    if (action === "excel") exportExcel();
    if (action === "png") exportPNG();
  });
  document.addEventListener("click", event => {
    if (!$("#export-menu").contains(event.target) && event.target !== $("#export-btn")) {
      $("#export-menu").setAttribute("hidden", "");
      $("#export-btn").setAttribute("aria-expanded", "false");
    }
  });
  $$("[data-map-mode]").forEach(button => button.addEventListener("click", () => GeoEngine.showMode(button.dataset.mapMode)));
  window.addEventListener("resize", debounce(() => {
    ChartEngine.resizeAll();
    AppState.map?.invalidateSize();
  }, 180));
  document.addEventListener("keydown", event => {
    if (event.target.matches("input, textarea, select")) return;
    const map = { "1": "overview", "2": "enrollment", "3": "payment", "4": "delivery", "5": "geo", "6": "pipeline", "7": "strategy" };
    if (map[event.key]) switchTab(map[event.key]);
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "r") {
      event.preventDefault();
      refreshData(false);
    }
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
