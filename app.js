const state = { sourceName: "", rows: [], results: [], meta: null };

const elements = {
  fileInput: document.getElementById("file-input"),
  dropZone: document.getElementById("drop-zone"),
  status: document.getElementById("status"),
  sign: document.getElementById("sign-convention"),
  unit: document.getElementById("frequency-unit"),
  mode: document.getElementById("calculation-mode"),
  thicknessField: document.getElementById("thickness-field"),
  thickness: document.getElementById("thickness-mm"),
  chartTitle: document.getElementById("chart-title"),
  formulaPrimary: document.getElementById("formula-primary"),
  modelNote: document.getElementById("model-note"),
  summary: document.getElementById("summary"),
  pointCount: document.getElementById("point-count"),
  frequencyRange: document.getElementById("frequency-range"),
  ignoredRows: document.getElementById("ignored-rows"),
  emptyState: document.getElementById("empty-state"),
  chart: document.getElementById("smith-chart"),
  chartWrap: document.getElementById("chart-wrap"),
  tooltip: document.getElementById("tooltip"),
  impedanceRange: document.getElementById("impedance-range"),
  exportCsv: document.getElementById("export-csv"),
  exportTxt: document.getElementById("export-txt"),
  exportSvg: document.getElementById("export-svg"),
  exportPng: document.getElementById("export-png"),
};

function normalizeHeader(value) {
  return String(value ?? "").trim().toLowerCase()
    .replaceAll("μ", "u").replaceAll("µ", "u").replaceAll("ε", "e")
    .replaceAll("频率", "frequency").replaceAll("介电常数", "e").replaceAll("磁导率", "u")
    .replaceAll("imaginary", "''").replaceAll("imag", "''").replaceAll("real", "'")
    .replace(/[\s_]/g, "");
}

function classifyHeader(value) {
  const text = normalizeHeader(value);
  if (text.includes("frequency") || ["freq", "f", "ghz"].includes(text)) return "frequency";
  if (["e'", "e1", "eprime", "er'", "epsilon'"].includes(text)) return "epsilon_real";
  if (["e''", "e2", "edoubleprime", "er''", "epsilon''", "eloss"].includes(text)) return "epsilon_loss";
  if (["u'", "u1", "uprime", "ur'", "mu'"].includes(text)) return "mu_real";
  if (["u''", "u2", "udoubleprime", "ur''", "mu''", "uloss"].includes(text)) return "mu_loss";
  return null;
}

function detectDelimiter(lines) {
  const candidates = ["\t", ",", ";"];
  return candidates
    .map(delimiter => ({ delimiter, score: lines.slice(0, 30).reduce((sum, line) => sum + line.split(delimiter).length - 1, 0) }))
    .sort((a, b) => b.score - a.score)[0].delimiter;
}

function splitDelimitedLine(line, delimiter) {
  const fields = [];
  let field = "";
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];
    if (character === '"') {
      if (quoted && line[index + 1] === '"') { field += '"'; index += 1; }
      else quoted = !quoted;
    } else if (character === delimiter && !quoted) {
      fields.push(field.trim()); field = "";
    } else field += character;
  }
  fields.push(field.trim());
  return fields;
}

function parseDelimited(text) {
  const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/).filter(line => line.trim());
  const delimiter = detectDelimiter(lines);
  const rows = lines.map(line => splitDelimitedLine(line, delimiter));
  const required = ["frequency", "epsilon_real", "epsilon_loss", "mu_real", "mu_loss"];
  let headerIndex = -1;
  let mapping = {};
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const candidate = {};
    rows[rowIndex].forEach((value, columnIndex) => {
      const field = classifyHeader(value);
      if (field && !(field in candidate)) candidate[field] = columnIndex;
    });
    if (required.every(field => field in candidate)) { headerIndex = rowIndex; mapping = candidate; break; }
  }
  if (headerIndex < 0) throw new Error("未找到 frequency、e'、e''、u'、u'' 五个数据列。");

  const data = [];
  let skipped = 0;
  rows.slice(headerIndex + 1).forEach(row => {
    const values = required.map(field => Number(row[mapping[field]]));
    if (values.every(Number.isFinite)) data.push(values);
    else skipped += 1;
  });
  if (!data.length) throw new Error("表头下方没有可用的数值数据。");
  return { rows: data, headerRow: headerIndex + 1, ignoredRows: headerIndex + skipped, sheet: "文本文件" };
}

function parseTabularRows(rows, sheet = "文本文件") {
  const required = ["frequency", "epsilon_real", "epsilon_loss", "mu_real", "mu_loss"];
  let headerIndex = -1;
  let mapping = {};
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const candidate = {};
    rows[rowIndex].forEach((value, columnIndex) => {
      const field = classifyHeader(value);
      if (field && !(field in candidate)) candidate[field] = columnIndex;
    });
    if (required.every(field => field in candidate)) { headerIndex = rowIndex; mapping = candidate; break; }
  }
  if (headerIndex < 0) throw new Error("未找到 frequency、e'、e''、u'、u'' 五个数据列。");

  const data = [];
  let skipped = 0;
  rows.slice(headerIndex + 1).forEach(row => {
    const values = required.map(field => Number(row[mapping[field]]));
    if (values.every(Number.isFinite)) data.push(values);
    else if (row.some(value => value !== null && value !== undefined && String(value).trim())) skipped += 1;
  });
  if (!data.length) throw new Error("表头下方没有可用的数值数据。");
  return { rows: data, headerRow: headerIndex + 1, ignoredRows: headerIndex + skipped, sheet };
}

async function parseExcel(file) {
  if (typeof XLSX === "undefined") throw new Error("Excel 解析组件未加载。请确认 xlsx.full.min.js 与 index.html 位于同一文件夹。");
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
  const errors = [];
  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
    try { return parseTabularRows(rows, sheetName); }
    catch (error) { errors.push(`${sheetName}: ${error.message}`); }
  }
  throw new Error(`工作簿中没有可识别的数据表。${errors.join(" ")}`);
}

function complexDivide(a, b) {
  const denominator = b.re * b.re + b.im * b.im;
  return { re: (a.re * b.re + a.im * b.im) / denominator, im: (a.im * b.re - a.re * b.im) / denominator };
}

function complexMultiply(a, b) {
  return { re: a.re * b.re - a.im * b.im, im: a.re * b.im + a.im * b.re };
}

function complexScale(value, factor) {
  return { re: value.re * factor, im: value.im * factor };
}

function complexTanh(value) {
  const doubleReal = 2 * value.re;
  const doubleImag = 2 * value.im;
  const denominator = Math.cosh(doubleReal) + Math.cos(doubleImag);
  return { re: Math.sinh(doubleReal) / denominator, im: Math.sin(doubleImag) / denominator };
}

function complexSqrt(value) {
  const magnitude = Math.hypot(value.re, value.im);
  let root = {
    re: Math.sqrt(Math.max(0, (magnitude + value.re) / 2)),
    im: Math.sign(value.im || 1) * Math.sqrt(Math.max(0, (magnitude - value.re) / 2)),
  };
  if (root.re < 0 || (root.re === 0 && root.im < 0)) root = { re: -root.re, im: -root.im };
  return root;
}

function impedanceToGamma(z) {
  return complexDivide({ re: z.re - 1, im: z.im }, { re: z.re + 1, im: z.im });
}

function calculate() {
  const sign = elements.sign.value === "minus" ? -1 : 1;
  const thicknessMeters = Number(elements.thickness.value) * 1e-3;
  if (elements.mode.value === "layer" && (!Number.isFinite(thicknessMeters) || thicknessMeters <= 0)) {
    throw new Error("吸波介质厚度必须大于 0 mm。");
  }
  const frequencyFactor = { GHz: 1e9, MHz: 1e6, Hz: 1 }[elements.unit.value];
  const lightSpeed = 299792458;
  state.results = state.rows.map(([frequency, epsilonReal, epsilonLoss, muReal, muLoss]) => {
    const epsilon = { re: epsilonReal, im: sign * epsilonLoss };
    const mu = { re: muReal, im: sign * muLoss };
    const intrinsic = complexSqrt(complexDivide(mu, epsilon));
    let impedance = intrinsic;
    if (elements.mode.value === "layer") {
      const propagationRoot = complexSqrt(complexMultiply(mu, epsilon));
      const phaseScale = 2 * Math.PI * frequency * frequencyFactor * thicknessMeters / lightSpeed;
      const argument = complexMultiply({ re: 0, im: phaseScale }, propagationRoot);
      impedance = complexMultiply(intrinsic, complexTanh(argument));
    }
    const gamma = impedanceToGamma(impedance);
    const gammaMagnitude = Math.hypot(gamma.re, gamma.im);
    const reflectionLoss = 20 * Math.log10(Math.max(gammaMagnitude, 1e-15));
    return { frequency, epsilonReal, epsilonLoss, muReal, muLoss, impedance, gamma, gammaMagnitude, reflectionLoss };
  });
}

function recalculate() {
  if (!state.rows.length) return;
  try {
    calculate();
    updateModeUi();
    updateSummary();
    renderChart();
    setStatus(`已重新计算 · ${elements.mode.value === "intrinsic" ? "本征阻抗" : `${formatNumber(Number(elements.thickness.value), 4)} mm 金属背板吸波层`}`, "success");
  } catch (error) {
    setStatus(error.message, "error");
  }
}

function updateModeUi() {
  const layerMode = elements.mode.value === "layer";
  elements.thicknessField.hidden = !layerMode;
  elements.chartTitle.textContent = layerMode ? "有限厚度吸波层输入阻抗轨迹" : "归一化本征阻抗轨迹";
  elements.formulaPrimary.innerHTML = layerMode
    ? "z<sub>in</sub> = √(μ<sub>r</sub>/ε<sub>r</sub>) tanh(jk<sub>0</sub>d√(μ<sub>r</sub>ε<sub>r</sub>))"
    : "z = √(μ<sub>r</sub> / ε<sub>r</sub>)";
  elements.modelNote.textContent = layerMode
    ? "厚度模式：自由空间正入射、单层均匀吸波介质、理想金属背板。"
    : "本征模式用于材料与自由空间的固有阻抗匹配分析。";
}

function setStatus(message, type = "") {
  elements.status.textContent = message;
  elements.status.className = `status ${type}`.trim();
}

async function loadFile(file) {
  if (!file) return;
  if (file.size > 25 * 1024 * 1024) { setStatus("文件超过 25 MB。", "error"); return; }
  setStatus(`正在识别 ${file.name}…`);
  try {
    const extension = file.name.split(".").pop().toLowerCase();
    let parsed;
    if (["xlsx", "xlsm", "xls"].includes(extension)) {
      parsed = await parseExcel(file);
    } else if (["csv", "txt", "tsv"].includes(extension)) {
      parsed = parseDelimited(await file.text());
    } else throw new Error("仅支持 XLSX、XLSM、CSV、TXT 和 TSV 文件。");

    state.sourceName = file.name.replace(/\.[^.]+$/, "");
    state.rows = parsed.rows;
    state.meta = parsed;
    calculate();
    updateModeUi();
    updateSummary();
    renderChart();
    setStatus(`已识别 ${file.name}${parsed.sheet ? ` · 工作表 ${parsed.sheet}` : ""}`, "success");
  } catch (error) {
    setStatus(error.message, "error");
  }
}

function updateSummary() {
  const frequencies = state.results.map(row => row.frequency);
  const zReal = state.results.map(row => row.impedance.re);
  const zImag = state.results.map(row => row.impedance.im);
  elements.summary.hidden = false;
  elements.pointCount.textContent = state.results.length.toLocaleString("zh-CN");
  elements.frequencyRange.textContent = `${formatNumber(Math.min(...frequencies))}–${formatNumber(Math.max(...frequencies))} ${elements.unit.value}`;
  elements.ignoredRows.textContent = state.meta.ignoredRows ?? 0;
  elements.impedanceRange.textContent = `z′ ${formatNumber(Math.min(...zReal), 3)}–${formatNumber(Math.max(...zReal), 3)} · z″ ${formatNumber(Math.min(...zImag), 3)}–${formatNumber(Math.max(...zImag), 3)}`;
  [elements.exportCsv, elements.exportTxt, elements.exportSvg, elements.exportPng].forEach(button => { button.disabled = false; });
}

function formatNumber(value, digits = 4) {
  return Number(value).toLocaleString("zh-CN", { maximumFractionDigits: digits });
}

function svgElement(name, attributes = {}) {
  const element = document.createElementNS("http://www.w3.org/2000/svg", name);
  Object.entries(attributes).forEach(([key, value]) => element.setAttribute(key, value));
  return element;
}

function gammaFromImpedance(resistance, reactance) {
  return impedanceToGamma({ re: resistance, im: reactance });
}

function frequencyColor(fraction) {
  const stops = [
    { at: 0, color: [21, 61, 104] },
    { at: 0.5, color: [20, 132, 122] },
    { at: 1, color: [240, 198, 87] },
  ];
  const left = fraction <= 0.5 ? stops[0] : stops[1];
  const right = fraction <= 0.5 ? stops[1] : stops[2];
  const local = (fraction - left.at) / (right.at - left.at);
  const channels = left.color.map((value, index) => Math.round(value + (right.color[index] - value) * local));
  return `rgb(${channels.join(",")})`;
}

function pathFromPoints(points, scale, centerX, centerY) {
  return points.map((point, index) => `${index ? "L" : "M"}${centerX + point.re * scale},${centerY - point.im * scale}`).join(" ");
}

function drawCompleteSmithGrid(svg, scale, centerX, centerY) {
  const resistanceMinor = [0.01, 0.02, 0.03, 0.05, 0.07, 0.1, 0.15, 0.2, 0.3, 0.4, 0.5, 0.7, 1, 1.5, 2, 3, 5, 7, 10, 20, 50];
  const resistanceMajor = new Set([0.1, 0.2, 0.5, 1, 2, 5, 10]);
  const reactanceMinor = [0.01, 0.02, 0.03, 0.05, 0.07, 0.1, 0.15, 0.2, 0.3, 0.4, 0.5, 0.7, 1, 1.5, 2, 3, 5, 7, 10, 20, 50];
  const reactanceMajor = new Set([0.1, 0.2, 0.5, 1, 2, 5, 10]);
  const impedanceGrid = svgElement("g", { fill: "none" });
  resistanceMinor.forEach(resistance => {
    const points = Array.from({ length: 900 }, (_, index) => gammaFromImpedance(resistance, -100 + index * 200 / 899));
    impedanceGrid.append(svgElement("path", {
      d: pathFromPoints(points, scale, centerX, centerY),
      stroke: resistanceMajor.has(resistance) ? "#8d928f" : "#c7cac7",
      "stroke-width": resistanceMajor.has(resistance) ? "0.75" : "0.38",
    }));
  });
  reactanceMinor.forEach(reactance => [-1, 1].forEach(sign => {
    const points = Array.from({ length: 800 }, (_, index) => gammaFromImpedance(index * 100 / 799, sign * reactance));
    impedanceGrid.append(svgElement("path", {
      d: pathFromPoints(points, scale, centerX, centerY),
      stroke: reactanceMajor.has(reactance) ? "#969b98" : "#ced0ce",
      "stroke-width": reactanceMajor.has(reactance) ? "0.7" : "0.35",
    }));
  }));
  svg.append(impedanceGrid);

  const admittanceGrid = svgElement("g", { fill: "none", opacity: "0.38" });
  [0.1, 0.2, 0.5, 1, 2, 5, 10].forEach(conductance => {
    const points = Array.from({ length: 700 }, (_, index) => {
      const admittance = { re: conductance, im: -50 + index * 100 / 699 };
      return impedanceToGamma(complexDivide({ re: 1, im: 0 }, admittance));
    });
    admittanceGrid.append(svgElement("path", { d: pathFromPoints(points, scale, centerX, centerY), stroke: "#7d8783", "stroke-width": "0.45", "stroke-dasharray": "2 2" }));
  });
  [0.1, 0.2, 0.5, 1, 2, 5, 10].forEach(susceptance => [-1, 1].forEach(sign => {
    const points = Array.from({ length: 650 }, (_, index) => {
      const admittance = { re: index * 50 / 649, im: sign * susceptance };
      return impedanceToGamma(complexDivide({ re: 1, im: 0 }, admittance));
    });
    admittanceGrid.append(svgElement("path", { d: pathFromPoints(points, scale, centerX, centerY), stroke: "#7d8783", "stroke-width": "0.4", "stroke-dasharray": "2 2" }));
  }));
  svg.append(admittanceGrid);

  svg.append(svgElement("circle", { cx: centerX, cy: centerY, r: scale, fill: "none", stroke: "#252c29", "stroke-width": "1.5" }));
  svg.append(svgElement("line", { x1: centerX - scale, y1: centerY, x2: centerX + scale, y2: centerY, stroke: "#575e5a", "stroke-width": "0.8" }));

  [0, 0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50].forEach(resistance => {
    const gamma = gammaFromImpedance(resistance, 0);
    const label = svgElement("text", { x: centerX + gamma.re * scale, y: centerY - 5, fill: "#515955", "font-size": "8.5", "text-anchor": "middle" });
    label.textContent = resistance;
    svg.append(label);
  });
  [0.1, 0.2, 0.5, 1, 2, 5, 10].forEach(reactance => [-1, 1].forEach(sign => {
    const gamma = gammaFromImpedance(0, sign * reactance);
    const label = svgElement("text", {
      x: centerX + gamma.re * scale * 1.015,
      y: centerY - gamma.im * scale * 1.015 + (sign > 0 ? -2 : 7),
      fill: "#69716d", "font-size": "7.5", "text-anchor": "end",
    });
    label.textContent = `${sign > 0 ? "+j" : "−j"}${reactance}`;
    svg.append(label);
  }));

  const phaseGroup = svgElement("g", { stroke: "#555d59", fill: "#555d59" });
  for (let degrees = -180; degrees < 180; degrees += 5) {
    const radians = degrees * Math.PI / 180;
    const major = degrees % 30 === 0;
    const inner = scale + (major ? 3 : 5);
    const outer = scale + (major ? 11 : 9);
    phaseGroup.append(svgElement("line", {
      x1: centerX + inner * Math.cos(radians), y1: centerY - inner * Math.sin(radians),
      x2: centerX + outer * Math.cos(radians), y2: centerY - outer * Math.sin(radians),
      "stroke-width": major ? "0.8" : "0.45",
    }));
    if (major) {
      const labelRadius = scale + 20;
      const label = svgElement("text", {
        x: centerX + labelRadius * Math.cos(radians), y: centerY - labelRadius * Math.sin(radians) + 3,
        "font-size": "7.5", "text-anchor": "middle", stroke: "none",
      });
      label.textContent = `${degrees}°`;
      phaseGroup.append(label);
    }
  }
  svg.append(phaseGroup);
}

function renderChart() {
  const svg = elements.chart;
  svg.replaceChildren();
  elements.emptyState.hidden = true;
  const width = Math.max(380, elements.chartWrap.clientWidth);
  const height = elements.chartWrap.clientHeight;
  const margin = { top: 34, right: 84, bottom: 38, left: 34 };
  const scale = Math.min((width - margin.left - margin.right) / 2, (height - margin.top - margin.bottom) / 2);
  const centerX = margin.left + scale;
  const centerY = margin.top + scale;
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);

  const defs = svgElement("defs");
  const gradient = svgElement("linearGradient", { id: "frequency-gradient", x1: "0%", y1: "100%", x2: "0%", y2: "0%" });
  [[0, "#153d68"], [50, "#14847a"], [100, "#f0c657"]].forEach(([offset, color]) => gradient.append(svgElement("stop", { offset: `${offset}%`, "stop-color": color })));
  defs.append(gradient);
  svg.append(defs);

  drawCompleteSmithGrid(svg, scale, centerX, centerY);

  const gammaPoints = state.results.map(row => row.gamma);
  const trajectory = svgElement("g", { fill: "none", "stroke-width": "3.2", "stroke-linecap": "round" });
  for (let index = 0; index < gammaPoints.length - 1; index += 1) {
    trajectory.append(svgElement("path", {
      d: pathFromPoints([gammaPoints[index], gammaPoints[index + 1]], scale, centerX, centerY),
      stroke: frequencyColor(index / Math.max(1, gammaPoints.length - 2)),
    }));
  }
  svg.append(trajectory);
  const markerGroup = svgElement("g");
  state.results.forEach((row, index) => {
    const marker = svgElement("circle", { cx: centerX + row.gamma.re * scale, cy: centerY - row.gamma.im * scale, r: "8", fill: "transparent", "data-index": index, tabindex: "0" });
    marker.addEventListener("pointerenter", showTooltip);
    marker.addEventListener("pointermove", showTooltip);
    marker.addEventListener("pointerleave", hideTooltip);
    marker.addEventListener("focus", showTooltip);
    marker.addEventListener("blur", hideTooltip);
    markerGroup.append(marker);
  });
  svg.append(markerGroup);

  const match = svgElement("g", { stroke: "#cf4b3f", "stroke-width": "2" });
  match.append(svgElement("line", { x1: centerX - 8, y1: centerY, x2: centerX + 8, y2: centerY }));
  match.append(svgElement("line", { x1: centerX, y1: centerY - 8, x2: centerX, y2: centerY + 8 }));
  svg.append(match);

  const legendX = centerX + scale + 25;
  const legendHeight = Math.min(280, scale * 1.2);
  svg.append(svgElement("rect", { x: legendX, y: centerY - legendHeight / 2, width: 10, height: legendHeight, fill: "url(#frequency-gradient)" }));
  const lowLabel = svgElement("text", { x: legendX + 16, y: centerY + legendHeight / 2, fill: "#69716d", "font-size": "11" });
  lowLabel.textContent = formatNumber(state.results[0].frequency);
  const highLabel = svgElement("text", { x: legendX + 16, y: centerY - legendHeight / 2 + 10, fill: "#69716d", "font-size": "11" });
  highLabel.textContent = formatNumber(state.results.at(-1).frequency);
  svg.append(lowLabel, highLabel);
}

function showTooltip(event) {
  const row = state.results[Number(event.currentTarget.dataset.index)];
  const rect = elements.chartWrap.getBoundingClientRect();
  const pointerX = event.clientX || rect.left + Number(event.currentTarget.getAttribute("cx"));
  const pointerY = event.clientY || rect.top + Number(event.currentTarget.getAttribute("cy"));
  elements.tooltip.innerHTML = `<strong>${formatNumber(row.frequency)} ${elements.unit.value}</strong><br>z = ${formatNumber(row.impedance.re, 5)} ${row.impedance.im >= 0 ? "+" : "−"} j${formatNumber(Math.abs(row.impedance.im), 5)}<br>|Γ| = ${formatNumber(row.gammaMagnitude, 5)}<br>RL = ${formatNumber(row.reflectionLoss, 3)} dB`;
  elements.tooltip.style.left = `${pointerX - rect.left}px`;
  elements.tooltip.style.top = `${pointerY - rect.top}px`;
  elements.tooltip.hidden = false;
}

function hideTooltip() { elements.tooltip.hidden = true; }

function exportRows(delimiter) {
  const sign = elements.sign.value === "minus" ? -1 : 1;
  const headers = ["frequency", "frequency_unit", "calculation_mode", "thickness_mm", "epsilon_real", "epsilon_loss", "epsilon_imag", "mu_real", "mu_loss", "mu_imag", "z_real", "z_imag", "gamma_real", "gamma_imag", "gamma_magnitude", "reflection_loss_dB"];
  const lines = [headers.join(delimiter)];
  state.results.forEach(row => lines.push([
    Number(row.frequency).toPrecision(12), elements.unit.value, elements.mode.value,
    elements.mode.value === "layer" ? Number(elements.thickness.value).toPrecision(12) : "",
    Number(row.epsilonReal).toPrecision(12), Number(row.epsilonLoss).toPrecision(12), Number(sign * row.epsilonLoss).toPrecision(12),
    Number(row.muReal).toPrecision(12), Number(row.muLoss).toPrecision(12), Number(sign * row.muLoss).toPrecision(12),
    Number(row.impedance.re).toPrecision(12), Number(row.impedance.im).toPrecision(12), Number(row.gamma.re).toPrecision(12), Number(row.gamma.im).toPrecision(12),
    Number(row.gammaMagnitude).toPrecision(12), Number(row.reflectionLoss).toPrecision(12),
  ].join(delimiter)));
  return lines.join("\r\n");
}

function downloadBlob(content, type, filename) {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const anchor = document.createElement("a");
  anchor.href = url; anchor.download = filename; anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function outputStem() {
  return elements.mode.value === "layer"
    ? `${state.sourceName}_layer_${String(elements.thickness.value).replace(".", "p")}mm`
    : `${state.sourceName}_intrinsic`;
}

function exportSvg() {
  const clone = elements.chart.cloneNode(true);
  clone.querySelectorAll("circle[fill='transparent']").forEach(node => node.remove());
  clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  downloadBlob(new XMLSerializer().serializeToString(clone), "image/svg+xml", `${outputStem()}_smith.svg`);
}

function exportPng() {
  const clone = elements.chart.cloneNode(true);
  clone.querySelectorAll("circle[fill='transparent']").forEach(node => node.remove());
  clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  const svgBlob = new Blob([new XMLSerializer().serializeToString(clone)], { type: "image/svg+xml" });
  const url = URL.createObjectURL(svgBlob);
  const image = new Image();
  image.onload = () => {
    const viewBox = elements.chart.viewBox.baseVal;
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(viewBox.width * 2); canvas.height = Math.round(viewBox.height * 2);
    const context = canvas.getContext("2d");
    context.fillStyle = "#fbfaf6"; context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    canvas.toBlob(blob => downloadBlob(blob, "image/png", `${outputStem()}_smith.png`), "image/png");
    URL.revokeObjectURL(url);
  };
  image.src = url;
}

elements.fileInput.addEventListener("change", event => loadFile(event.target.files[0]));
["dragenter", "dragover"].forEach(type => elements.dropZone.addEventListener(type, event => { event.preventDefault(); elements.dropZone.classList.add("is-dragging"); }));
["dragleave", "drop"].forEach(type => elements.dropZone.addEventListener(type, event => { event.preventDefault(); elements.dropZone.classList.remove("is-dragging"); }));
elements.dropZone.addEventListener("drop", event => loadFile(event.dataTransfer.files[0]));
elements.sign.addEventListener("change", recalculate);
elements.unit.addEventListener("change", recalculate);
elements.mode.addEventListener("change", () => { updateModeUi(); recalculate(); });
elements.thickness.addEventListener("change", recalculate);
elements.exportCsv.addEventListener("click", () => downloadBlob(`\uFEFF${exportRows(",")}`, "text/csv;charset=utf-8", `${outputStem()}_origin.csv`));
elements.exportTxt.addEventListener("click", () => downloadBlob(exportRows("\t"), "text/plain;charset=utf-8", `${outputStem()}_origin.txt`));
elements.exportSvg.addEventListener("click", exportSvg);
elements.exportPng.addEventListener("click", exportPng);
window.addEventListener("resize", () => { if (state.results.length) renderChart(); });
updateModeUi();
