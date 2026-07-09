const DEFAULT_DATA = {
  markets: {
    us: {
      label: "美国市场",
      marketCode: "US MARKET",
      defaultManager: "BlackRock/iShares",
      managers: [
        { rank: 1, scaleRank: 1, ecosystemRank: 2, name: "BlackRock/iShares", slug: "blackrock-ishares", etfScale: 4559912130000.00, marketShare: "28.78%", scores: { "产品广度": 100.0, "费率优势": 92.6, "交易活跃": 72.3, "多元能力": 85.4, "旗舰产品": 68.7 }, ecosystemScore: 83.8 },
        { rank: 2, scaleRank: 2, ecosystemRank: 6, name: "Vanguard", slug: "vanguard", etfScale: 4490894470000.00, marketShare: "28.34%", scores: { "产品广度": 67.2, "费率优势": 100.0, "交易活跃": 75.5, "多元能力": 61.7, "旗舰产品": 75.6 }, ecosystemScore: 76.0 },
        { rank: 3, scaleRank: 3, ecosystemRank: 3, name: "State Street/SPDR", slug: "state-street-spdr", etfScale: 1997886650000.00, marketShare: "12.61%", scores: { "产品广度": 73.9, "费率优势": 96.3, "交易活跃": 77.1, "多元能力": 80.0, "旗舰产品": 88.4 }, ecosystemScore: 83.1 },
        { rank: 4, scaleRank: 4, ecosystemRank: 1, name: "Invesco", slug: "invesco", etfScale: 984693614520.00, marketShare: "6.21%", scores: { "产品广度": 77.9, "费率优势": 88.2, "交易活跃": 75.1, "多元能力": 99.2, "旗舰产品": 100.0 }, ecosystemScore: 88.1 },
        { rank: 5, scaleRank: 5, ecosystemRank: 5, name: "Charles Schwab", slug: "charles-schwab", etfScale: 583807300000.00, marketShare: "3.68%", scores: { "产品广度": 60.0, "费率优势": 98.2, "交易活跃": 81.0, "多元能力": 80.6, "旗舰产品": 72.3 }, ecosystemScore: 78.4 },
        { rank: 6, scaleRank: 6, ecosystemRank: 7, name: "JPMorgan AM", slug: "jpmorgan-am", etfScale: 324763370000.00, marketShare: "2.05%", scores: { "产品广度": 63.4, "费率优势": 84.9, "交易活跃": 67.9, "多元能力": 84.2, "旗舰产品": 70.3 }, ecosystemScore: 74.1 },
        { rank: 7, scaleRank: 7, ecosystemRank: 8, name: "Dimensional", slug: "dimensional", etfScale: 299310000000.00, marketShare: "1.89%", scores: { "产品广度": 60.7, "费率优势": 88.6, "交易活跃": 72.0, "多元能力": 80.0, "旗舰产品": 68.0 }, ecosystemScore: 73.9 },
        { rank: 8, scaleRank: 8, ecosystemRank: 10, name: "First Trust", slug: "first-trust", etfScale: 279797124680.00, marketShare: "1.77%", scores: { "产品广度": 81.5, "费率优势": 60.0, "交易活跃": 61.3, "多元能力": 100.0, "旗舰产品": 60.0 }, ecosystemScore: 72.6 },
        { rank: 9, scaleRank: 9, ecosystemRank: 9, name: "Fidelity", slug: "fidelity", etfScale: 172005260000.00, marketShare: "1.09%", scores: { "产品广度": 65.5, "费率优势": 87.8, "交易活跃": 64.2, "多元能力": 80.0, "旗舰产品": 69.4 }, ecosystemScore: 73.4 },
        { rank: 10, scaleRank: 10, ecosystemRank: 4, name: "VanEck", slug: "vaneck", etfScale: 160409092720.00, marketShare: "1.01%", scores: { "产品广度": 66.5, "费率优势": 71.5, "交易活跃": 83.9, "多元能力": 89.8, "旗舰产品": 96.5 }, ecosystemScore: 81.6 }
      ]
    },
    china: {
      label: "中国市场",
      marketCode: "CHINA MARKET",
      defaultManager: "富国基金",
      managers: [
        { rank: 1, scaleRank: 1, ecosystemRank: 6, name: "华夏基金", slug: "chinaamc", etfScale: 5965.56, marketShare: "12.52%", scores: { "产品广度": 99.8, "费率优势": 86.2, "交易活跃": 73.3, "多元能力": 86.6, "旗舰产品": 60.0 }, ecosystemScore: 81.2 },
        { rank: 2, scaleRank: 2, ecosystemRank: 5, name: "易方达基金", slug: "e-fund", etfScale: 5821.65, marketShare: "12.19%", scores: { "产品广度": 99.2, "费率优势": 100.0, "交易活跃": 62.5, "多元能力": 83.2, "旗舰产品": 62.5 }, ecosystemScore: 81.5 },
        { rank: 3, scaleRank: 3, ecosystemRank: 8, name: "国泰基金", slug: "guotai-fund", etfScale: 3715.02, marketShare: "7.89%", scores: { "产品广度": 73.7, "费率优势": 66.4, "交易活跃": 83.9, "多元能力": 85.1, "旗舰产品": 67.5 }, ecosystemScore: 75.3 },
        { rank: 4, scaleRank: 4, ecosystemRank: 1, name: "华泰柏瑞基金", slug: "huatai-pinebridge", etfScale: 2999.29, marketShare: "6.17%", scores: { "产品广度": 66.0, "费率优势": 93.6, "交易活跃": 100.0, "多元能力": 80.0, "旗舰产品": 97.5 }, ecosystemScore: 87.4 },
        { rank: 5, scaleRank: 5, ecosystemRank: 10, name: "广发基金", slug: "gf-fund", etfScale: 2593.97, marketShare: "5.53%", scores: { "产品广度": 74.7, "费率优势": 60.0, "交易活跃": 76.6, "多元能力": 89.6, "旗舰产品": 61.8 }, ecosystemScore: 72.5 },
        { rank: 6, scaleRank: 6, ecosystemRank: 3, name: "南方基金", slug: "china-southern", etfScale: 2455.50, marketShare: "5.31%", scores: { "产品广度": 75.9, "费率优势": 97.9, "交易活跃": 90.5, "多元能力": 83.9, "旗舰产品": 68.0 }, ecosystemScore: 83.2 },
        { rank: 7, scaleRank: 7, ecosystemRank: 2, name: "博时基金", slug: "bosera", etfScale: 2292.66, marketShare: "4.88%", scores: { "产品广度": 66.7, "费率优势": 89.5, "交易活跃": 82.3, "多元能力": 100.0, "旗舰产品": 83.5 }, ecosystemScore: 84.4 },
        { rank: 8, scaleRank: 8, ecosystemRank: 7, name: "嘉实基金", slug: "harvest-fund", etfScale: 2247.00, marketShare: "4.37%", scores: { "产品广度": 66.7, "费率优势": 85.4, "交易活跃": 77.7, "多元能力": 82.4, "旗舰产品": 81.4 }, ecosystemScore: 78.7 },
        { rank: 9, scaleRank: 9, ecosystemRank: 9, name: "华宝基金", slug: "huabao", etfScale: 2092.66, marketShare: "4.40%", scores: { "产品广度": 60.0, "费率优势": 65.8, "交易活跃": 76.8, "多元能力": 74.8, "旗舰产品": 98.9 }, ecosystemScore: 75.3 },
        { rank: 10, scaleRank: 10, ecosystemRank: 4, name: "富国基金", slug: "fullgoal", etfScale: 1976.67, marketShare: "4.26%", scores: { "产品广度": 81.1, "费率优势": 96.4, "交易活跃": 67.2, "多元能力": 94.0, "旗舰产品": 76.1 }, ecosystemScore: 83.0 }
      ]
    }
  }
};

const dimensions = ["产品广度", "费率优势", "交易活跃", "多元能力", "旗舰产品"];
const rankPositionClasses = ["product-breadth", "fee-advantage", "trading-activity", "multi-capability", "flagship-product"];
const logoPaths = {
  china: {
    chinaamc: "assets/logos/svg/CN_01_华夏基金.svg",
    "e-fund": "assets/logos/svg/CN_02_易方达基金.svg",
    "guotai-fund": "assets/logos/svg/CN_03_国泰基金.svg",
    "huatai-pinebridge": "assets/logos/svg/CN_04_华泰柏瑞基金.svg",
    "gf-fund": "assets/logos/svg/CN_05_广发基金.svg",
    "china-southern": "assets/logos/svg/CN_06_南方基金.svg",
    bosera: "assets/logos/svg/CN_07_博时基金.svg",
    "harvest-fund": "assets/logos/svg/CN_08_嘉实基金.svg",
    huabao: "assets/logos/svg/CN_09_华宝基金.svg",
    fullgoal: "assets/logos/svg/CN_10_富国基金.svg"
  },
  us: {
    "blackrock-ishares": "assets/logos/svg/US_02_BlackRock_iShares.svg",
    vanguard: "assets/logos/svg/US_06_Vanguard.svg",
    "state-street-spdr": "assets/logos/svg/US_03_StateStreet_SPDR.svg",
    invesco: "assets/logos/svg/US_01_Invesco.svg",
    "charles-schwab": "assets/logos/svg/US_05_CharlesSchwab.svg",
    "jpmorgan-am": "assets/logos/svg/US_07_JPMorganAM.svg",
    dimensional: "assets/logos/svg/US_08_Dimensional.svg",
    "first-trust": "assets/logos/svg/US_10_FirstTrust.svg",
    fidelity: "assets/logos/svg/US_09_Fidelity.svg",
    vaneck: "assets/logos/svg/US_04_VanEck.svg"
  }
};
let state = {
  data: DEFAULT_DATA,
  market: "china",
  selectedName: "富国基金",
  mode: "profile",
  compareNames: {
    china: ["富国基金", "华夏基金"],
    us: ["BlackRock/iShares", "Invesco"]
  }
};

const els = {
  managerList: document.getElementById("managerList"),
  marketLabel: document.getElementById("marketLabel"),
  compareLegend: document.getElementById("compareLegend"),
  selectedName: document.getElementById("selectedName"),
  brandMarket: document.getElementById("brandMarket"),
  brandRank: document.getElementById("brandRank"),
  brandLogo: document.getElementById("brandLogo"),
  logoFallback: document.getElementById("logoFallback"),
  logoFrame: document.querySelector(".logo-frame"),
  profileBrandView: document.getElementById("profileBrandView"),
  compareBrandView: document.getElementById("compareBrandView"),
  brandName: document.getElementById("brandName"),
  brandSub: document.getElementById("brandSub"),
  ecosystemScore: document.getElementById("ecosystemScore"),
  metricReadout: document.getElementById("metricReadout"),
  dimensionRankLayer: document.getElementById("dimensionRankLayer"),
  insightText: document.getElementById("insightText"),
  radarGrid: document.getElementById("radarGrid"),
  radarLabels: document.getElementById("radarLabels"),
  radarShape: document.getElementById("radarShape"),
  radarStage: document.querySelector(".radar-stage")
};

function rankLabel(rank) {
  return `RANK ${String(rank).padStart(2, "0")}`;
}

function managerScaleRank(manager) {
  return manager.scaleRank ?? manager.rank;
}

function logoPathFor(manager) {
  return manager.logo || logoPaths[state.market]?.[manager.slug] || "";
}

function logoClassFor(manager) {
  const wideNames = ["BlackRock/iShares", "State Street/SPDR", "Charles Schwab", "JPMorgan AM", "Dimensional", "First Trust"];
  const classes = ["brand-logo", `brand-logo-${state.market}`];
  if (wideNames.includes(manager.name)) classes.push("brand-logo-wide");
  return classes.join(" ");
}

function formatRankLabel(rank) {
  if (rank === 1) return "BEST";
  if (rank === 2) return "2nd";
  if (rank === 3) return "3rd";
  return `${rank}th`;
}

function rankClass(rank) {
  if (rank === 1) return "rank-best";
  if (rank === 2) return "rank-second";
  if (rank === 3) return "rank-third";
  return "rank-normal";
}

function currentMarket() {
  return state.data.markets[state.market];
}

function currentManager() {
  const market = currentMarket();
  return market.managers.find((manager) => manager.name === state.selectedName) || market.managers[0];
}

function managerByName(name, market = currentMarket()) {
  return market.managers.find((manager) => manager.name === name);
}

function getDefaultComparePair(marketKey = state.market) {
  const market = state.data.markets[marketKey];
  const preferred = marketKey === "china"
    ? ["富国基金", "华夏基金"]
    : ["BlackRock/iShares", "Invesco"];
  const valid = preferred.filter((name) => market.managers.some((manager) => manager.name === name));
  if (valid.length === 2) return valid;
  return market.managers.slice(0, 2).map((manager) => manager.name);
}

function resetComparePair(marketKey = state.market) {
  state.compareNames[marketKey] = getDefaultComparePair(marketKey);
}

function currentCompareManagers() {
  const market = currentMarket();
  const names = state.compareNames[state.market] || [];
  return names.map((name) => managerByName(name, market)).filter(Boolean);
}

function toggleCompareManager(manager) {
  const names = [...(state.compareNames[state.market] || [])];
  const existingIndex = names.indexOf(manager.name);

  if (existingIndex >= 0) {
    names.splice(existingIndex, 1);
  } else if (names.length < 2) {
    names.push(manager.name);
  } else {
    names[1] = manager.name;
  }

  state.compareNames[state.market] = names;
}

function polarPoint(index, value, radius = 198) {
  const angle = -Math.PI / 2 + index * (Math.PI * 2 / dimensions.length);
  const scaledRadius = radius * (value / 100);
  return {
    x: 260 + Math.cos(angle) * scaledRadius,
    y: 260 + Math.sin(angle) * scaledRadius
  };
}

function pointsToString(points) {
  return points.map((point) => `${point.x.toFixed(2)},${point.y.toFixed(2)}`).join(" ");
}

function getDimensionRanks(marketData) {
  const ranksByManager = {};

  marketData.managers.forEach((manager) => {
    ranksByManager[manager.name] = {};
  });

  dimensions.forEach((dimension) => {
    const scores = marketData.managers.map((manager) => manager.scores[dimension]);

    marketData.managers.forEach((manager) => {
      const score = manager.scores[dimension];
      const rank = scores.filter((otherScore) => otherScore > score).length + 1;
      ranksByManager[manager.name][dimension] = rank;
    });
  });

  return ranksByManager;
}

function drawGrid() {
  const rings = [20, 40, 60, 80, 100];
  els.radarGrid.innerHTML = "";
  els.radarLabels.innerHTML = "";

  rings.forEach((ring) => {
    const points = dimensions.map((_, index) => polarPoint(index, ring));
    const polygon = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
    polygon.setAttribute("points", pointsToString(points));
    polygon.setAttribute("class", `radar-ring${ring === 100 ? " major" : ""}`);
    els.radarGrid.appendChild(polygon);
  });

  dimensions.forEach((dimension, index) => {
    const end = polarPoint(index, 100);
    const axis = document.createElementNS("http://www.w3.org/2000/svg", "line");
    axis.setAttribute("x1", "260");
    axis.setAttribute("y1", "260");
    axis.setAttribute("x2", end.x.toFixed(2));
    axis.setAttribute("y2", end.y.toFixed(2));
    axis.setAttribute("class", "radar-axis");
    els.radarGrid.appendChild(axis);
  });
}

function drawRadar(manager) {
  const values = dimensions.map((dimension) => manager.scores[dimension]);
  const points = values.map((value, index) => polarPoint(index, value));
  els.radarShape.innerHTML = "";

  const polygon = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
  polygon.setAttribute("points", pointsToString(points));
  polygon.setAttribute("class", "radar-polygon");
  els.radarShape.appendChild(polygon);

  points.forEach((point, index) => {
    const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    circle.setAttribute("cx", point.x.toFixed(2));
    circle.setAttribute("cy", point.y.toFixed(2));
    circle.setAttribute("r", "5");
    circle.setAttribute("class", "radar-point");
    els.radarShape.appendChild(circle);

    const valuePoint = polarPoint(index, Math.min(values[index] + 10, 112));
    const valueLabel = document.createElementNS("http://www.w3.org/2000/svg", "text");
    valueLabel.setAttribute("x", valuePoint.x.toFixed(2));
    valueLabel.setAttribute("y", valuePoint.y.toFixed(2));
    valueLabel.setAttribute("class", "radar-value-label");
    valueLabel.setAttribute("text-anchor", "middle");
    valueLabel.setAttribute("dominant-baseline", "middle");
    valueLabel.textContent = values[index].toFixed(1);
    els.radarShape.appendChild(valueLabel);
  });

  els.radarStage.classList.remove("switching");
  void els.radarStage.offsetWidth;
  els.radarStage.classList.add("switching");
}

function drawCompareRadar(managerA, managerB) {
  els.radarShape.innerHTML = "";

  [
    { manager: managerA, className: "compare-a", pointClass: "compare-point-a" },
    { manager: managerB, className: "compare-b", pointClass: "compare-point-b" }
  ].filter((series) => series.manager).forEach((series) => {
    const values = dimensions.map((dimension) => series.manager.scores[dimension]);
    const points = values.map((value, index) => polarPoint(index, value));
    const polygon = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
    polygon.setAttribute("points", pointsToString(points));
    polygon.setAttribute("class", `radar-polygon compare-radar-polygon ${series.className}`);
    els.radarShape.appendChild(polygon);

    points.forEach((point) => {
      const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      circle.setAttribute("cx", point.x.toFixed(2));
      circle.setAttribute("cy", point.y.toFixed(2));
      circle.setAttribute("r", "4");
      circle.setAttribute("class", `radar-point compare-radar-point ${series.pointClass}`);
      els.radarShape.appendChild(circle);
    });
  });

  els.radarStage.classList.remove("switching");
  void els.radarStage.offsetWidth;
  els.radarStage.classList.add("switching");
}

function renderManagerList() {
  const market = currentMarket();
  const managers = [...market.managers].sort((a, b) => managerScaleRank(a) - managerScaleRank(b));
  const compareNames = state.compareNames[state.market] || [];
  els.managerList.innerHTML = managers.map((manager) => `
    <button class="manager-item${manager.name === state.selectedName && state.mode === "profile" ? " active" : ""}${compareNames[0] === manager.name && state.mode === "compare" ? " compare-a-active" : ""}${compareNames[1] === manager.name && state.mode === "compare" ? " compare-b-active" : ""}" data-name="${manager.name}" type="button">
      <span class="rank">${rankLabel(managerScaleRank(manager))}</span>
      <span class="manager-title" title="${manager.name}">${manager.name}</span>
      <span class="manager-score">${manager.marketShare || "--"}</span>
      ${state.mode === "compare" && compareNames.includes(manager.name) ? `<span class="compare-slot ${compareNames[0] === manager.name ? "slot-a" : "slot-b"}">${compareNames[0] === manager.name ? "A" : "B"}</span>` : ""}
    </button>
  `).join("");
}

function renderMetrics(manager) {
  els.metricReadout.innerHTML = dimensions.map((dimension) => `
    <div class="metric-chip">
      <span>${dimension}</span>
      <strong>${manager.scores[dimension].toFixed(1)}</strong>
    </div>
  `).join("");
}

function renderCompareMetrics(managerA, managerB) {
  els.metricReadout.innerHTML = dimensions.map((dimension) => {
    if (!managerA || !managerB) {
      return `
        <div class="metric-chip compare-metric-chip">
          <span>${dimension}</span>
          <strong>请选择两家</strong>
          <em>等待对比对象</em>
        </div>
      `;
    }

    const valueA = managerA.scores[dimension];
    const valueB = managerB.scores[dimension];
    const diff = +(valueA - valueB).toFixed(1);
    const leader = diff >= 0 ? "A" : "B";
    const diffText = Math.abs(diff) < 0.05 ? "EVEN" : `${leader} +${Math.abs(diff).toFixed(1)}`;
    const diffClass = Math.abs(diff) < 0.05 ? "diff-even" : leader === "A" ? "diff-a" : "diff-b";

    return `
      <div class="metric-chip compare-metric-chip">
        <span>${dimension}</span>
        <strong><b class="metric-a">A</b> ${valueA.toFixed(1)} <i>|</i> <b class="metric-b">B</b> ${valueB.toFixed(1)}</strong>
        <em class="${diffClass}">${diffText}</em>
      </div>
    `;
  }).join("");
}

function renderDimensionRanks(manager, market) {
  const ranks = getDimensionRanks(market)[manager.name];
  els.dimensionRankLayer.querySelectorAll(".dimension-rank").forEach((badge) => {
    const dimension = badge.dataset.dimension;
    const rank = ranks[dimension];
    const positionClass = rankPositionClasses.find((className) => badge.classList.contains(className));
    badge.className = `dimension-rank ${positionClass} rank-badge ${rankClass(rank)}`;
    badge.innerHTML = `
      <strong>${formatRankLabel(rank)}</strong>
      <span>${dimension}</span>
    `;
  });
}

function renderCompareDimensionRanks(managerA, managerB, market) {
  const allRanks = getDimensionRanks(market);
  els.dimensionRankLayer.querySelectorAll(".dimension-rank").forEach((badge) => {
    const dimension = badge.dataset.dimension;
    const positionClass = rankPositionClasses.find((className) => badge.classList.contains(className));
    const rankA = managerA ? allRanks[managerA.name][dimension] : null;
    const rankB = managerB ? allRanks[managerB.name][dimension] : null;
    badge.className = `dimension-rank ${positionClass} compare-rank-badge`;
    badge.innerHTML = `
      <span class="compare-rank-line rank-line-a ${rankA === 1 ? "is-best" : ""}">A ${rankA ? formatRankLabel(rankA) : "--"}</span>
      <span class="compare-rank-line rank-line-b ${rankB === 1 ? "is-best" : ""}">B ${rankB ? formatRankLabel(rankB) : "--"}</span>
      <small>${dimension}</small>
    `;
  });
}

function insightFor(manager) {
  const scores = manager.scores;
  if (manager.name === "富国基金") {
    return "富国的优势不在单点爆发，而在产品广度、费率和多元能力；下一步关键是补强入口产品心智与交易沉淀。";
  }
  if (scores["产品广度"] >= 80 && scores["交易活跃"] < 75) {
    return "产品矩阵较完整，但交易入口和流动性沉淀仍是提升重点。";
  }
  if (scores["旗舰产品"] >= 85) {
    return "旗舰产品心智突出，是生态入口的重要抓手。";
  }
  if (scores["多元能力"] >= 88) {
    return "多元能力较强，适合从产品线扩张走向生态化运营。";
  }
  if (scores["费率优势"] >= 90) {
    return "成本端竞争力突出，但仍需通过入口产品强化用户留存。";
  }
  return "生态画像较为均衡，后续可围绕优势维度继续强化差异化表达。";
}

function topDimensionNames(items) {
  return items.slice(0, 2).map((item) => item.dimension).join("和");
}

function generateCompareInsight(managerA, managerB) {
  if (!managerA || !managerB) {
    return "请选择两家同市场管理人，系统将基于五维差异、规模排名、市场份额和生态评分生成对比洞察。";
  }

  const diffs = dimensions.map((dimension) => ({
    dimension,
    diff: +(managerA.scores[dimension] - managerB.scores[dimension]).toFixed(1)
  }));
  const aLead = diffs.filter((item) => item.diff > 0).sort((a, b) => b.diff - a.diff);
  const bLead = diffs.filter((item) => item.diff < 0).sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
  const aText = aLead.length ? topDimensionNames(aLead) : "结构均衡性";
  const bText = bLead.length ? topDimensionNames(bLead) : "结构均衡性";
  const scoreDiff = +(managerA.ecosystemScore - managerB.ecosystemScore).toFixed(1);
  const scoreText = Math.abs(scoreDiff) < 1 ? "生态评分接近" : scoreDiff > 0 ? `${managerA.name}生态评分略高` : `${managerB.name}生态评分略高`;

  return `${managerA.name}在${aText}上更突出，${managerB.name}在${bText}和规模入口上更强（${rankLabel(managerScaleRank(managerB))}，份额${managerB.marketShare}）；${scoreText}，该对比仅用于观察同市场生态结构差异。`;
}

function updateLogo(manager) {
  const logoPath = logoPathFor(manager);
  els.logoFrame.classList.remove("logo-loaded");
  els.brandLogo.hidden = true;
  els.brandLogo.removeAttribute("src");
  els.logoFallback.hidden = false;
  els.logoFallback.textContent = manager.name;

  if (!logoPath) return;

  const probe = new Image();
  probe.onload = () => {
    els.brandLogo.src = logoPath;
    els.brandLogo.alt = `${manager.name} logo`;
    els.brandLogo.className = logoClassFor(manager);
    els.brandLogo.hidden = false;
    els.logoFallback.hidden = true;
    els.logoFrame.classList.add("logo-loaded");
  };
  probe.onerror = () => {
    els.brandLogo.hidden = true;
    els.logoFallback.hidden = false;
    els.logoFrame.classList.remove("logo-loaded");
  };
  probe.src = logoPath;
}

function compareLogoMarkup(manager, slot) {
  if (!manager) {
    return `
      <div class="compare-logo-card slot-${slot.toLowerCase()}">
        <div class="compare-logo-fallback">待选择</div>
      </div>
    `;
  }

  const logoPath = logoPathFor(manager);
  return `
    <div class="compare-logo-card slot-${slot.toLowerCase()}">
      ${logoPath ? `<img class="compare-logo ${logoClassFor(manager)}" src="${logoPath}" alt="${manager.name} logo">` : ""}
      <div class="compare-logo-fallback" ${logoPath ? "hidden" : ""}>${manager.name}</div>
    </div>
  `;
}

function hydrateCompareLogos() {
  els.compareBrandView.querySelectorAll(".compare-logo").forEach((image) => {
    const fallback = image.nextElementSibling;
    image.addEventListener("error", () => {
      image.hidden = true;
      if (fallback) fallback.hidden = false;
    }, { once: true });
  });
}

function renderCompareBrandCard(managerA, managerB, market) {
  els.profileBrandView.hidden = true;
  els.compareBrandView.hidden = false;
  els.compareBrandView.innerHTML = `
    <div class="brand-topline compare-topline">
      <span>${market.marketCode}</span>
      <strong>A vs B</strong>
    </div>
    <div class="vs-arena">
      ${compareLogoMarkup(managerA, "A")}
      <div class="vs-mark">VS</div>
      ${compareLogoMarkup(managerB, "B")}
    </div>
    <div class="compare-profile-lines">
      ${[["A", managerA], ["B", managerB]].map(([slot, manager]) => manager ? `
        <div class="compare-profile-line slot-${slot.toLowerCase()}">
          <strong><span>${slot}</span>${manager.name}</strong>
          <p>${rankLabel(managerScaleRank(manager))} | 份额 ${manager.marketShare || "--"} | 生态评分 ${manager.ecosystemScore.toFixed(1)}</p>
        </div>
      ` : `
        <div class="compare-profile-line slot-${slot.toLowerCase()} muted-line">
          <strong><span>${slot}</span>待选择</strong>
          <p>请在左侧选择对比对象</p>
        </div>
      `).join("")}
    </div>
  `;
  hydrateCompareLogos();
}

function renderProfileMode(manager, market) {
  els.selectedName.textContent = manager.name;
  els.compareLegend.hidden = true;
  els.profileBrandView.hidden = false;
  els.compareBrandView.hidden = true;

  els.brandMarket.textContent = market.marketCode;
  els.brandRank.textContent = rankLabel(managerScaleRank(manager));
  els.brandName.textContent = manager.name;
  els.brandSub.textContent = `${market.label} ETF Manager Ecosystem Profile`;
  els.ecosystemScore.textContent = manager.ecosystemScore.toFixed(1);
  els.insightText.textContent = insightFor(manager);

  renderMetrics(manager);
  renderDimensionRanks(manager, market);
  drawRadar(manager);
  updateLogo(manager);
}

function renderCompareLegend(managerA, managerB) {
  els.compareLegend.hidden = false;
  els.compareLegend.innerHTML = `
    <span class="legend-item legend-a"><i></i>A ${managerA ? managerA.name : "待选择"}</span>
    <span class="legend-item legend-b"><i></i>B ${managerB ? managerB.name : "待选择"}</span>
  `;
}

function renderCompareMode(market) {
  const [managerA, managerB] = currentCompareManagers();
  els.selectedName.textContent = `${managerA ? managerA.name : "A"} VS ${managerB ? managerB.name : "B"}`;
  renderCompareLegend(managerA, managerB);
  renderCompareMetrics(managerA, managerB);
  renderCompareDimensionRanks(managerA, managerB, market);
  drawCompareRadar(managerA, managerB);
  renderCompareBrandCard(managerA, managerB, market);
  els.insightText.textContent = generateCompareInsight(managerA, managerB);
}

function render() {
  const market = currentMarket();
  const manager = currentManager();

  document.querySelectorAll(".market-btn").forEach((button) => {
    button.classList.toggle("active", button.dataset.market === state.market);
  });

  document.querySelectorAll(".mode-btn").forEach((button) => {
    button.classList.toggle("active", button.dataset.mode === state.mode);
  });

  els.marketLabel.textContent = state.mode === "compare" ? `${market.label} A/B 对比` : `${market.label}占比`;
  renderManagerList();

  if (state.mode === "compare") {
    renderCompareMode(market);
  } else {
    renderProfileMode(manager, market);
  }
}

function bindEvents() {
  document.querySelector(".mode-switch").addEventListener("click", (event) => {
    const button = event.target.closest(".mode-btn");
    if (!button) return;
    state.mode = button.dataset.mode;
    if (state.mode === "compare") {
      resetComparePair(state.market);
    }
    render();
  });

  document.querySelector(".market-switch").addEventListener("click", (event) => {
    const button = event.target.closest(".market-btn");
    if (!button) return;
    state.market = button.dataset.market;
    state.selectedName = currentMarket().defaultManager;
    resetComparePair(state.market);
    render();
  });

  els.managerList.addEventListener("click", (event) => {
    const button = event.target.closest(".manager-item");
    if (!button) return;
    const manager = managerByName(button.dataset.name);
    if (!manager) return;
    if (state.mode === "compare") {
      toggleCompareManager(manager);
    } else {
      state.selectedName = manager.name;
    }
    render();
  });
}

async function init() {
  bindEvents();
  drawGrid();

  try {
    const response = await fetch("data/etf_scores.json", { cache: "no-store" });
    if (response.ok) {
      state.data = await response.json();
      state.selectedName = state.data.markets.china.defaultManager;
      resetComparePair("china");
      resetComparePair("us");
    }
  } catch (error) {
    // file:// often blocks fetch; DEFAULT_DATA keeps the dashboard usable by double-click.
  }

  render();
}

init();
