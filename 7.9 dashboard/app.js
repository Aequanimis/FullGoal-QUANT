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
let state = {
  data: DEFAULT_DATA,
  market: "china",
  selectedName: "富国基金"
};

const els = {
  managerList: document.getElementById("managerList"),
  marketLabel: document.getElementById("marketLabel"),
  selectedName: document.getElementById("selectedName"),
  brandMarket: document.getElementById("brandMarket"),
  brandRank: document.getElementById("brandRank"),
  brandLogo: document.getElementById("brandLogo"),
  logoFallback: document.getElementById("logoFallback"),
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

function renderManagerList() {
  const market = currentMarket();
  const managers = [...market.managers].sort((a, b) => managerScaleRank(a) - managerScaleRank(b));
  els.managerList.innerHTML = managers.map((manager) => `
    <button class="manager-item${manager.name === state.selectedName ? " active" : ""}" data-name="${manager.name}" type="button">
      <span class="rank">${rankLabel(managerScaleRank(manager))}</span>
      <span class="manager-title" title="${manager.name}">${manager.name}</span>
      <span class="manager-score">${manager.marketShare || "--"}</span>
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

function updateLogo(manager) {
  const logoPath = `assets/logos/${manager.slug}.png`;
  els.brandLogo.hidden = true;
  els.brandLogo.removeAttribute("src");
  els.logoFallback.hidden = false;
  els.logoFallback.textContent = manager.name;

  const probe = new Image();
  probe.onload = () => {
    els.brandLogo.src = logoPath;
    els.brandLogo.alt = `${manager.name} logo`;
    els.brandLogo.hidden = false;
    els.logoFallback.hidden = true;
  };
  probe.onerror = () => {
    els.brandLogo.hidden = true;
    els.logoFallback.hidden = false;
  };
  probe.src = logoPath;
}

function render() {
  const market = currentMarket();
  const manager = currentManager();

  document.querySelectorAll(".market-btn").forEach((button) => {
    button.classList.toggle("active", button.dataset.market === state.market);
  });

  els.marketLabel.textContent = `${market.label}占比`;
  els.selectedName.textContent = manager.name;
  els.brandMarket.textContent = market.marketCode;
  els.brandRank.textContent = rankLabel(managerScaleRank(manager));
  els.brandName.textContent = manager.name;
  els.brandSub.textContent = `${market.label} ETF Manager Ecosystem Profile`;
  els.ecosystemScore.textContent = manager.ecosystemScore.toFixed(1);
  els.insightText.textContent = insightFor(manager);

  renderManagerList();
  renderMetrics(manager);
  renderDimensionRanks(manager, market);
  drawRadar(manager);
  updateLogo(manager);
}

function bindEvents() {
  document.querySelector(".market-switch").addEventListener("click", (event) => {
    const button = event.target.closest(".market-btn");
    if (!button) return;
    state.market = button.dataset.market;
    state.selectedName = currentMarket().defaultManager;
    render();
  });

  els.managerList.addEventListener("click", (event) => {
    const button = event.target.closest(".manager-item");
    if (!button) return;
    state.selectedName = button.dataset.name;
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
    }
  } catch (error) {
    // file:// often blocks fetch; DEFAULT_DATA keeps the dashboard usable by double-click.
  }

  render();
}

init();
