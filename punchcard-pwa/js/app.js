// === CONSTANTS ===
const PATTERN_COLS = [
  19.25, 23.75, 28.25, 32.75, 37.25, 41.75, 46.25, 50.75,
  55.25, 59.75, 64.25, 68.75, 73.25, 77.75, 82.25, 86.75,
  91.25, 95.75, 100.25, 104.75, 109.25, 113.75, 118.25, 122.75
];
const EDGE_GUIDE_X = [13.5, 128.5];
const TRANSPORT_X = [6.75, 135.25];
const PATTERN_R = 1.75;
const EDGE_R = 1.5;
const TRANSPORT_R = 1.75;
const CARD_WIDTH = 142.0;
const STROKE_W = 0.1;

// === STATE ===
let stitchCount = parseInt(localStorage.getItem('sf_stitchCount')) || 24;
function getStitchCount() { return stitchCount; }

let currentGrid = [];
let originalImageData = null;
let imageWidth = 24;
let imageHeight = 0;
let generatedSVGString = null;
let imageProcessingMode = 'simple';
let sourceImage = null;
let cropRect = { x: 0, y: 0, w: 1, h: 1 };
let cropDragState = null;
let detectedCells = null;
let chartScanImageData = null;
let gridOverlayInfo = null;
let pgComposition = [];
let currentCatId = 'all';
let currentRowColors = [];
let pgPalette = [];
const PG_PALETTE_KEY = 'pgPaletteColors';
const DEFAULT_PALETTE = ['#111011','#2e2627','#3b1720','#7f0e1e','#bc312d','#9e2e50','#18322b','#71b0a2','#ba9373','#f5e0ca'];
let cellColorOverrides = {};
let activePaintColor = null;
let isPainting = false;
const DEFAULT_COLORS = { a: '#27200f', b: '#e8ddd0' };

// Section accordion state (no sidebar needed — panels always visible)

// === UNDO/REDO ===
const UNDO_LIMIT = 30;
let undoStack = [];
let redoStack = [];

function saveUndoState() {
  undoStack.push({
    grid: currentGrid.map(r => r),
    rowColors: currentRowColors.map(rc => ({ a: rc.a, b: rc.b })),
    cellOverrides: Object.assign({}, cellColorOverrides)
  });
  if (undoStack.length > UNDO_LIMIT) undoStack.shift();
  redoStack = [];
  updateUndoButtons();
}

function undo() {
  if (undoStack.length === 0) return;
  redoStack.push({
    grid: currentGrid.map(r => r),
    rowColors: currentRowColors.map(rc => ({ a: rc.a, b: rc.b })),
    cellOverrides: Object.assign({}, cellColorOverrides)
  });
  const prev = undoStack.pop();
  currentGrid = prev.grid;
  currentRowColors = prev.rowColors;
  cellColorOverrides = prev.cellOverrides;
  renderPreview(currentGrid);
  generatedSVGString = null;
  svgOutput.innerHTML = '';
  updateUndoButtons();
}

function redo() {
  if (redoStack.length === 0) return;
  undoStack.push({
    grid: currentGrid.map(r => r),
    rowColors: currentRowColors.map(rc => ({ a: rc.a, b: rc.b })),
    cellOverrides: Object.assign({}, cellColorOverrides)
  });
  const next = redoStack.pop();
  currentGrid = next.grid;
  currentRowColors = next.rowColors;
  cellColorOverrides = next.cellOverrides;
  renderPreview(currentGrid);
  generatedSVGString = null;
  svgOutput.innerHTML = '';
  updateUndoButtons();
}

function updateUndoButtons() {
  var undoBtn = document.getElementById('undoBtn');
  var redoBtn = document.getElementById('redoBtn');
  if (currentGrid.length > 0) {
    undoBtn.style.display = '';
    redoBtn.style.display = '';
    undoBtn.disabled = undoStack.length === 0;
    redoBtn.disabled = redoStack.length === 0;
  } else {
    undoBtn.style.display = 'none';
    redoBtn.style.display = 'none';
  }
}

function updateActivePaintChip() {
  var chip = document.getElementById('activePaintChip');
  if (!chip) return;
  if (!activePaintColor) { chip.style.display = 'none'; return; }
  chip.style.display = '';
  if (activePaintColor === 'eraser') {
    chip.style.background = '';
    chip.textContent = '🧹';
    chip.title = 'Active: Eraser';
  } else {
    chip.style.background = activePaintColor;
    chip.textContent = '';
    chip.title = 'Active color: ' + activePaintColor;
  }
}

function ensureRowColors(grid) {
  if (currentRowColors.length === grid.length) return;
  currentRowColors = grid.map(() => ({ a: DEFAULT_COLORS.a, b: DEFAULT_COLORS.b }));
  cellColorOverrides = {};
}

// === DOM REFS ===
const blankRows = document.getElementById('blankRows');
const createBlankBtn = document.getElementById('createBlankBtn');
const txtFileInput = document.getElementById('txtFileInput');
const textInput = document.getElementById('textInput');
const textError = document.getElementById('textError');
const imageInput = document.getElementById('imageInput');
const hiddenCanvas = document.getElementById('hiddenCanvas');
const analysisCanvas = document.getElementById('analysisCanvas');
const imageModeToggle = document.getElementById('imageModeToggle');
const simpleModeBtn = document.getElementById('simpleModeBtn');
const chartScanBtn = document.getElementById('chartScanBtn');
const simplePreviewContainer = document.getElementById('simplePreviewContainer');
const simplePreviewImg = document.getElementById('simplePreviewImg');
const cropContainer = document.getElementById('cropContainer');
const cropCanvas = document.getElementById('cropCanvas');
const chartScanControls = document.getElementById('chartScanControls');
const sampleSlider = document.getElementById('sampleSlider');
const sampleValueEl = document.getElementById('sampleValue');
const chartRowsInput = document.getElementById('chartRows');
const detectGridBtn = document.getElementById('detectGridBtn');
const gridStatus = document.getElementById('gridStatus');
const imageModeHint = document.getElementById('imageModeHint');
const imageControls = document.getElementById('imageControls');
const thresholdSlider = document.getElementById('thresholdSlider');
const thresholdValue = document.getElementById('thresholdValue');
const invertToggle = document.getElementById('invertToggle');
const previewGrid = document.getElementById('previewGrid');
const rowCount = document.getElementById('rowCount');
const previewHint = document.getElementById('previewHint');
const generateBtn = document.getElementById('generateBtn');
const downloadCutBtn = document.getElementById('downloadCutBtn');
const downloadDrawBtn = document.getElementById('downloadDrawBtn');
const downloadCombinedBtn = document.getElementById('downloadCombinedBtn');
const downloadActions = document.getElementById('downloadActions');
const svgOutput = document.getElementById('svgOutput');
const jacquardToggle = document.getElementById('jacquardToggle');
const refreshAllBtn = document.getElementById('refreshAllBtn');

// Layout DOM refs (3-column, no sidebar)
const canvasEmptyState = document.getElementById('canvasEmptyState');
const statusText = document.getElementById('statusText');
const txtFileImport = document.getElementById('txtFileImport');

// === ACCORDION SECTION MANAGEMENT ===
document.querySelectorAll('.sf-section-header').forEach(header => {
  header.addEventListener('click', () => {
    const section = header.closest('.sf-section');
    if (section) section.classList.toggle('collapsed');
  });
});

// === INPUT TAB SWITCHING (3-way selector) ===
(function initInputTabs() {
  const tabs = document.querySelectorAll('.sf-input-tab');
  const panels = {
    image: document.getElementById('inputImage'),
    text: document.getElementById('inputText'),
    blank: document.getElementById('inputBlank')
  };
  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      const target = tab.dataset.input;
      Object.keys(panels).forEach(key => {
        if (panels[key]) panels[key].style.display = key === target ? '' : 'none';
      });
    });
  });
})();

// === MOBILE PANEL NAVIGATION ===
(function initMobileNav() {
  const leftPanel = document.getElementById('leftPanel');
  const rightPanel = document.getElementById('rightPanel');
  const mobileNavBtns = document.querySelectorAll('.sf-mobile-nav-btn');

  function isMobile() {
    return window.matchMedia('(max-width: 800px)').matches;
  }

  function closeMobilePanels() {
    leftPanel.classList.remove('mobile-open');
    rightPanel.classList.remove('mobile-open');
  }

  function setActiveMobileTab(panel) {
    mobileNavBtns.forEach(btn => {
      btn.classList.toggle('active', btn.dataset.mobilePanel === panel);
    });
  }

  mobileNavBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      if (!isMobile()) return;
      const panel = btn.dataset.mobilePanel;

      if (panel === 'canvas') {
        closeMobilePanels();
        setActiveMobileTab('canvas');
        return;
      }

      if (panel === 'left') {
        rightPanel.classList.remove('mobile-open');
        leftPanel.classList.add('mobile-open');
        setActiveMobileTab('left');
        return;
      }

      if (panel === 'right') {
        leftPanel.classList.remove('mobile-open');
        rightPanel.classList.add('mobile-open');
        setActiveMobileTab('right');
      }
    });
  });

  // Reset mobile state when resizing to desktop
  window.addEventListener('resize', () => {
    if (!isMobile()) {
      closeMobilePanels();
      setActiveMobileTab('canvas');
    }
  });
})();

// .txt file import from dropdown
txtFileImport.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    const text = ev.target.result.replace(/\r/g, '');
    const { grid, errors } = parseTextInput(text);
    if (errors.length === 0 && grid.length > 0) {
      saveUndoState();
      currentGrid = grid;
      ensureRowColors(grid);
      renderPreview(grid);
    }
    txtFileImport.value = '';
  };
  reader.readAsText(file);
});

// === SETTINGS BUTTON — scrolls to and opens settings section ===
document.getElementById('settingsBtn').addEventListener('click', () => {
  const section = document.getElementById('sectionSettings');
  if (section) {
    section.classList.remove('collapsed');
    section.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
});

// === SETTINGS HANDLERS ===
(function initSettings() {
  const overlayModeSelect = document.getElementById('overlayModeSelect');
  const resetAllSettings = document.getElementById('resetAllSettings');
  const stitchCountInput = document.getElementById('stitchCountInput');

  // localStorage migration: sf_colNumbers → sf_overlayMode
  if (!localStorage.getItem('sf_overlayMode')) {
    if (localStorage.getItem('sf_colNumbers') === 'true') {
      localStorage.setItem('sf_overlayMode', 'colnums');
    }
    localStorage.removeItem('sf_colNumbers');
  }

  // Load saved settings
  const savedMachine = localStorage.getItem('sf_defaultMachine');
  if (savedMachine) {
    var machineProfile = document.getElementById('machineProfile');
    if (machineProfile) machineProfile.value = savedMachine;
  }

  // Load overlay mode
  const savedOverlay = localStorage.getItem('sf_overlayMode');
  if (savedOverlay && overlayModeSelect) overlayModeSelect.value = savedOverlay;

  // Load stitch count
  if (stitchCountInput) stitchCountInput.value = String(getStitchCount());

  // Overlay mode select
  if (overlayModeSelect) {
    overlayModeSelect.addEventListener('change', () => {
      localStorage.setItem('sf_overlayMode', overlayModeSelect.value);
      renderPreview(currentGrid);
    });
  }

  // Stitch count input
  if (stitchCountInput) {
    stitchCountInput.addEventListener('change', function() {
      var newVal = Math.max(1, Math.min(200, parseInt(this.value) || 24));
      this.value = String(newVal);
      if (newVal === stitchCount) return;
      if (currentGrid.length > 0 && !confirm('Changing column count will clear the current canvas. Continue?')) {
        this.value = String(stitchCount);
        return;
      }
      stitchCount = newVal;
      localStorage.setItem('sf_stitchCount', stitchCount);
      currentGrid = [];
      currentRowColors = [];
      cellColorOverrides = {};
      undoStack = [];
      redoStack = [];
      document.documentElement.style.setProperty('--stitch-count', stitchCount);
      updateSvgExportGuard();
      renderPreview([]);
      updateUndoButtons();
      updateEmptyState();
      updateStatusBar();
      var fragGrid = document.getElementById('pgFragGrid');
      if (fragGrid) renderFragGrid(currentCatId || 'all');
    });
  }

  // Reset all settings
  resetAllSettings.addEventListener('click', () => {
    localStorage.removeItem('sf_defaultMachine');
    localStorage.removeItem('sf_colNumbers');
    localStorage.removeItem('sf_overlayMode');
    if (overlayModeSelect) overlayModeSelect.value = 'none';
    var machineProfile = document.getElementById('machineProfile');
    if (machineProfile) machineProfile.value = '0';
    renderPreview(currentGrid);
  });
})();

function updateSvgExportGuard() {
  var note = document.getElementById('svgColNote');
  if (stitchCount !== 24) {
    if (note) note.style.display = '';
    if (generateBtn) { generateBtn.disabled = true; generateBtn.title = 'SVG export requires 24 columns'; }
  } else {
    if (note) note.style.display = 'none';
    if (generateBtn) { generateBtn.disabled = currentGrid.length === 0; generateBtn.title = ''; }
  }
}

// === EMPTY STATE MANAGEMENT ===
function updateEmptyState() {
  if (canvasEmptyState) {
    canvasEmptyState.style.display = currentGrid.length > 0 ? 'none' : '';
  }
}

function updateStatusBar() {
  if (!statusText) return;
  if (currentGrid.length === 0) { statusText.textContent = ''; return; }

  var parts = [currentGrid.length + ' rows \u00d7 ' + getStitchCount() + ' stitches'];

  var machineEl = document.getElementById('machineProfile');
  var offset = machineEl ? parseInt(machineEl.value) : 0;
  if (offset > 0) parts.push('offset\u00a0+' + offset);

  if (undoStack.length > 0) parts.push(undoStack.length + ' undo' + (undoStack.length !== 1 ? 's' : ''));

  statusText.textContent = parts.join('\u2003\u00b7\u2003');
}

// === CUSTOM FILE BUTTONS ===
document.getElementById('imageInputBtn').addEventListener('click', () => imageInput.click());
document.getElementById('txtFileBtn').addEventListener('click', () => txtFileInput.click());
imageInput.addEventListener('change', function() {
  document.getElementById('imageFileName').textContent = this.files[0] ? this.files[0].name : '';
});
txtFileInput.addEventListener('change', function() {
  document.getElementById('txtFileName').textContent = this.files[0] ? this.files[0].name : '';
});

// === REFRESH ALL ===
refreshAllBtn.addEventListener('click', () => {
  if (!confirm('Clear canvas? This will reset the grid to 20 blank rows. This cannot be undone.')) return;

  saveUndoState();
  currentGrid = Array.from({ length: 20 }, () => '-'.repeat(getStitchCount()));
  currentRowColors = [];
  ensureRowColors(currentGrid);
  cellColorOverrides = {};
  isPainting = false;
  activePaintColor = null;
  updateActivePaintChip();

  generatedSVGString = null;
  svgOutput.innerHTML = '';
  downloadCutBtn.disabled = true;
  downloadDrawBtn.disabled = true;
  downloadCombinedBtn.disabled = true;
  downloadActions.style.display = 'none';
  var previewWrap = document.getElementById('svgPreviewWrap');
  if (previewWrap) previewWrap.style.display = 'none';

  renderPreview(currentGrid);
  updateEmptyState();
  updateStatusBar();
});

// === BLANK CANVAS ===
function createBlankGrid() {
  const rows = Math.max(1, Math.min(200, parseInt(blankRows.value) || 20));
  blankRows.value = rows;
  saveUndoState();
  currentGrid = Array.from({ length: rows }, () => '-'.repeat(getStitchCount()));
  ensureRowColors(currentGrid);
  renderPreview(currentGrid);
}

createBlankBtn.addEventListener('click', createBlankGrid);

// === TEXT INPUT ===
let textDebounce = null;
textInput.addEventListener('input', () => {
  clearTimeout(textDebounce);
  textDebounce = setTimeout(handleTextChange, 300);
});

txtFileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    textInput.value = ev.target.result.replace(/\r/g, '');
    handleTextChange();
    txtFileInput.value = '';
  };
  reader.readAsText(file);
});

function handleTextChange() {
  const text = textInput.value;
  if (!text.trim()) {
    currentGrid = [];
    textError.textContent = '';
    renderPreview([]);
    generateBtn.disabled = true;
    return;
  }
  const { grid, errors } = parseTextInput(text);
  textError.textContent = errors.join('\n');
  if (errors.length === 0) {
    saveUndoState();
    currentGrid = grid;
    ensureRowColors(grid);
    renderPreview(grid);
  } else {
    currentGrid = [];
    renderPreview([]);
  }
}

function parseTextInput(text) {
  const lines = text.split('\n');
  const errors = [];
  const grid = [];

  while (lines.length > 0 && lines[lines.length - 1].trim() === '') {
    lines.pop();
  }

  if (lines.length === 0) {
    errors.push('At least 1 row is required.');
    return { grid, errors };
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].toLowerCase().replace(/\s+$/, '');
    const sc = getStitchCount();
    if (line.length !== sc) {
      errors.push('Line ' + (i + 1) + ': expected ' + sc + ' characters, got ' + line.length);
    }
    if (!/^[x\-]+$/.test(line)) {
      errors.push('Line ' + (i + 1) + ': only \'x\' and \'-\' characters allowed');
    }
    grid.push(line.padEnd(sc, '-').substring(0, sc));
  }

  return { grid, errors };
}

// === IMAGE UPLOAD ===
imageInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    const img = new Image();
    img.onload = () => {
      sourceImage = img;
      imageModeToggle.classList.remove('hidden');
      imageModeToggle.style.display = 'flex';
      imageModeHint.classList.remove('hidden');
      if (imageProcessingMode === 'simple') {
        runSimpleMode(img);
      } else {
        initChartScanMode(img);
      }
    };
    img.src = ev.target.result;
  };
  reader.readAsDataURL(file);
});

function runSimpleMode(img) {
  cropContainer.classList.add('hidden');
  chartScanControls.classList.add('hidden');
  gridStatus.textContent = '';
  imageControls.style.display = 'block';
  simplePreviewImg.src = img.src;
  simplePreviewContainer.classList.remove('hidden');
  const targetWidth = getStitchCount();
  const aspect = img.height / img.width;
  const targetHeight = Math.round(targetWidth * aspect);
  imageWidth = targetWidth;
  imageHeight = targetHeight;
  const ctx = hiddenCanvas.getContext('2d');
  hiddenCanvas.width = targetWidth;
  hiddenCanvas.height = targetHeight;
  ctx.imageSmoothingEnabled = true;
  ctx.drawImage(img, 0, 0, targetWidth, targetHeight);
  originalImageData = ctx.getImageData(0, 0, targetWidth, targetHeight);
  processImage();
}

function initChartScanMode(img) {
  simplePreviewContainer.classList.add('hidden');
  cropContainer.classList.remove('hidden');
  chartScanControls.classList.remove('hidden');
  imageControls.style.display = 'block';
  cropRect = { x: 0, y: 0, w: 1, h: 1 };
  detectedCells = null;
  chartScanImageData = null;
  gridOverlayInfo = null;
  gridOverlayToggle.checked = false;
  gridOverlayToggleRow.style.display = 'none';
  drawCrop();
}

simpleModeBtn.addEventListener('click', () => {
  imageProcessingMode = 'simple';
  simpleModeBtn.classList.add('active');
  chartScanBtn.classList.remove('active');
  if (sourceImage) runSimpleMode(sourceImage);
});

chartScanBtn.addEventListener('click', () => {
  imageProcessingMode = 'chartScan';
  chartScanBtn.classList.add('active');
  simpleModeBtn.classList.remove('active');
  if (sourceImage) initChartScanMode(sourceImage);
});

thresholdSlider.addEventListener('input', () => {
  thresholdValue.textContent = thresholdSlider.value;
  if (imageProcessingMode === 'simple' && originalImageData) {
    processImage();
  } else if (imageProcessingMode === 'chartScan' && detectedCells && chartScanImageData) {
    reprocessChartScan();
  }
});

invertToggle.addEventListener('change', () => {
  if (imageProcessingMode === 'simple' && originalImageData) {
    processImage();
  } else if (imageProcessingMode === 'chartScan' && detectedCells && chartScanImageData) {
    reprocessChartScan();
  }
});

sampleSlider.addEventListener('input', () => {
  sampleValueEl.textContent = sampleSlider.value + '%';
  if (detectedCells && chartScanImageData) {
    reprocessChartScan();
  }
});

function reprocessChartScan() {
  const ratio = parseInt(sampleSlider.value) / 100;
  const threshold = parseInt(thresholdSlider.value);
  const invert = invertToggle.checked;
  const grid = processChartScanGrid(chartScanImageData, detectedCells, threshold, invert, ratio);
  saveUndoState();
  currentGrid = grid;
  ensureRowColors(grid);
  renderPreview(grid);
}

function processImage() {
  if (!originalImageData) return;
  const threshold = parseInt(thresholdSlider.value);
  const invert = invertToggle.checked;
  const pixels = originalImageData.data;
  const grid = [];

  const sc = getStitchCount();
  for (let row = 0; row < imageHeight; row++) {
    let line = '';
    for (let col = 0; col < sc; col++) {
      const idx = (row * sc + col) * 4;
      const r = pixels[idx];
      const g = pixels[idx + 1];
      const b = pixels[idx + 2];
      const luminance = 0.299 * r + 0.587 * g + 0.114 * b;
      const isDark = luminance < threshold;
      const isPunched = invert ? !isDark : isDark;
      line += isPunched ? 'x' : '-';
    }
    grid.push(line);
  }

  saveUndoState();
  currentGrid = grid;
  ensureRowColors(grid);
  renderPreview(grid);

  if (imageProcessingMode === 'simple' && sourceImage) {
    gridOverlayInfo = { img: sourceImage, cols: getStitchCount(), rows: imageHeight, cropRect: null };
    document.getElementById('gridOverlayToggleRow').style.display = '';
    drawGridOverlay();
  }
}

// === GRID OVERLAY ===
var simpleGridOverlay = document.getElementById('simpleGridOverlay');
var gridOverlayToggle = document.getElementById('gridOverlayToggle');
var gridOverlayToggleRow = document.getElementById('gridOverlayToggleRow');

function drawGridLines(ctx, x0, y0, w, h, cols, rows) {
  ctx.strokeStyle = 'rgba(255, 80, 80, 0.6)';
  ctx.lineWidth = 1;
  var colStep = w / cols;
  var rowStep = h / rows;

  for (var c = 0; c <= cols; c++) {
    var x = Math.round(x0 + c * colStep) + 0.5;
    ctx.beginPath();
    ctx.moveTo(x, y0);
    ctx.lineTo(x, y0 + h);
    ctx.stroke();
  }
  for (var r = 0; r <= rows; r++) {
    var y = Math.round(y0 + r * rowStep) + 0.5;
    ctx.beginPath();
    ctx.moveTo(x0, y);
    ctx.lineTo(x0 + w, y);
    ctx.stroke();
  }

  var fontSize = Math.max(9, Math.min(14, Math.floor(colStep * 0.6)));
  ctx.font = 'bold ' + fontSize + 'px sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  for (var c = 0; c < cols; c++) {
    var cx = x0 + (c + 0.5) * colStep;
    var ty = y0 + 2;
    var text = String(c + 1);
    var tw = ctx.measureText(text).width;
    ctx.fillStyle = 'rgba(0,0,0,0.55)';
    ctx.fillRect(cx - tw / 2 - 2, ty, tw + 4, fontSize + 2);
    ctx.fillStyle = '#fff';
    ctx.fillText(text, cx, ty + 1);
  }

  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  var rowFontSize = Math.max(8, Math.min(13, Math.floor(rowStep * 0.55)));
  ctx.font = 'bold ' + rowFontSize + 'px sans-serif';
  for (var r = 0; r < rows; r++) {
    var cy = y0 + (r + 0.5) * rowStep;
    var text = String(r + 1);
    var tw = ctx.measureText(text).width;
    ctx.fillStyle = 'rgba(0,0,0,0.55)';
    ctx.fillRect(x0 + 1, cy - rowFontSize / 2 - 1, tw + 4, rowFontSize + 2);
    ctx.fillStyle = '#fff';
    ctx.fillText(text, x0 + 3, cy);
  }
}

function drawSimpleGridOverlay() {
  if (!gridOverlayInfo || !gridOverlayToggle.checked || imageProcessingMode !== 'simple') {
    simpleGridOverlay.width = 0;
    simpleGridOverlay.height = 0;
    return;
  }
  var img = simplePreviewImg;
  var w = img.naturalWidth;
  var h = img.naturalHeight;
  if (!w || !h) return;
  simpleGridOverlay.width = w;
  simpleGridOverlay.height = h;
  var ctx = simpleGridOverlay.getContext('2d');
  drawGridLines(ctx, 0, 0, w, h, gridOverlayInfo.cols, gridOverlayInfo.rows);
}

function drawGridOverlay() {
  if (imageProcessingMode === 'simple') {
    drawSimpleGridOverlay();
  } else if (imageProcessingMode === 'chartScan') {
    drawCrop();
  }
}

gridOverlayToggle.addEventListener('change', drawGridOverlay);

// === CROP UI ===
function drawCrop() {
  if (!sourceImage) return;
  const containerWidth = cropContainer.clientWidth || 560;
  const aspect = sourceImage.height / sourceImage.width;
  const displayW = containerWidth;
  const displayH = Math.round(containerWidth * aspect);
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  cropCanvas.style.width = displayW + 'px';
  cropCanvas.style.height = displayH + 'px';
  cropCanvas.width = displayW * dpr;
  cropCanvas.height = displayH * dpr;
  const ctx = cropCanvas.getContext('2d');
  ctx.scale(dpr, dpr);
  ctx.drawImage(sourceImage, 0, 0, displayW, displayH);
  const cx = cropRect.x * displayW;
  const cy = cropRect.y * displayH;
  const cw = cropRect.w * displayW;
  const ch = cropRect.h * displayH;
  ctx.fillStyle = 'rgba(0,0,0,0.45)';
  ctx.fillRect(0, 0, displayW, cy);
  ctx.fillRect(0, cy + ch, displayW, displayH - cy - ch);
  ctx.fillRect(0, cy, cx, ch);
  ctx.fillRect(cx + cw, cy, displayW - cx - cw, ch);
  ctx.strokeStyle = '#ff6b9d';
  ctx.lineWidth = 2;
  ctx.strokeRect(cx, cy, cw, ch);
  const hs = 12;
  var handles = [
    [cx, cy], [cx + cw, cy], [cx, cy + ch], [cx + cw, cy + ch],
    [cx + cw / 2, cy], [cx + cw / 2, cy + ch],
    [cx, cy + ch / 2], [cx + cw, cy + ch / 2]
  ];
  for (var hi = 0; hi < handles.length; hi++) {
    var hx = Math.max(0, Math.min(displayW - hs, handles[hi][0] - hs / 2));
    var hy = Math.max(0, Math.min(displayH - hs, handles[hi][1] - hs / 2));
    ctx.fillStyle = '#ff6b9d';
    ctx.fillRect(hx, hy, hs, hs);
    ctx.strokeStyle = '#000';
    ctx.lineWidth = 1;
    ctx.strokeRect(hx, hy, hs, hs);
  }

  if (gridOverlayInfo && gridOverlayToggle.checked && imageProcessingMode === 'chartScan') {
    ctx.save();
    ctx.beginPath();
    ctx.rect(cx, cy, cw, ch);
    ctx.clip();
    drawGridLines(ctx, cx, cy, cw, ch, gridOverlayInfo.cols, gridOverlayInfo.rows);
    ctx.restore();
  }
}

function getCropPointerPos(e) {
  const rect = cropCanvas.getBoundingClientRect();
  return {
    nx: (e.clientX - rect.left) / rect.width,
    ny: (e.clientY - rect.top) / rect.height
  };
}

function hitTestHandle(nx, ny) {
  const tol = 0.04;
  const cr = cropRect;
  const cx = cr.x, cy = cr.y, cw = cr.w, ch = cr.h;
  var corners = [
    { id: 'tl', x: cx, y: cy }, { id: 'tr', x: cx + cw, y: cy },
    { id: 'bl', x: cx, y: cy + ch }, { id: 'br', x: cx + cw, y: cy + ch }
  ];
  for (var i = 0; i < corners.length; i++) {
    if (Math.abs(nx - corners[i].x) < tol && Math.abs(ny - corners[i].y) < tol)
      return corners[i].id;
  }
  var edges = [
    { id: 't', x: cx + cw / 2, y: cy }, { id: 'b', x: cx + cw / 2, y: cy + ch },
    { id: 'l', x: cx, y: cy + ch / 2 }, { id: 'r', x: cx + cw, y: cy + ch / 2 }
  ];
  for (var i = 0; i < edges.length; i++) {
    if (Math.abs(nx - edges[i].x) < tol && Math.abs(ny - edges[i].y) < tol)
      return edges[i].id;
  }
  if (nx > cx && nx < cx + cw && ny > cy && ny < cy + ch) return 'move';
  return null;
}

cropCanvas.addEventListener('pointerdown', (e) => {
  const { nx, ny } = getCropPointerPos(e);
  const handle = hitTestHandle(nx, ny);
  if (!handle) return;
  cropDragState = { handle: handle, startNx: nx, startNy: ny, startRect: Object.assign({}, cropRect) };
  cropCanvas.setPointerCapture(e.pointerId);
  e.preventDefault();
});

cropCanvas.addEventListener('pointermove', (e) => {
  if (!cropDragState) return;
  const { nx, ny } = getCropPointerPos(e);
  const dx = nx - cropDragState.startNx;
  const dy = ny - cropDragState.startNy;
  const sr = cropDragState.startRect;
  const minSize = 0.08;
  var r = Object.assign({}, cropRect);
  switch (cropDragState.handle) {
    case 'move':
      r.x = Math.max(0, Math.min(1 - sr.w, sr.x + dx));
      r.y = Math.max(0, Math.min(1 - sr.h, sr.y + dy));
      r.w = sr.w; r.h = sr.h;
      break;
    case 'tl':
      r.x = Math.max(0, Math.min(sr.x + sr.w - minSize, sr.x + dx));
      r.y = Math.max(0, Math.min(sr.y + sr.h - minSize, sr.y + dy));
      r.w = sr.x + sr.w - r.x; r.h = sr.y + sr.h - r.y;
      break;
    case 'tr':
      r.w = Math.max(minSize, Math.min(1 - sr.x, sr.w + dx));
      r.y = Math.max(0, Math.min(sr.y + sr.h - minSize, sr.y + dy));
      r.h = sr.y + sr.h - r.y;
      break;
    case 'bl':
      r.x = Math.max(0, Math.min(sr.x + sr.w - minSize, sr.x + dx));
      r.w = sr.x + sr.w - r.x;
      r.h = Math.max(minSize, Math.min(1 - sr.y, sr.h + dy));
      break;
    case 'br':
      r.w = Math.max(minSize, Math.min(1 - sr.x, sr.w + dx));
      r.h = Math.max(minSize, Math.min(1 - sr.y, sr.h + dy));
      break;
    case 't':
      r.y = Math.max(0, Math.min(sr.y + sr.h - minSize, sr.y + dy));
      r.h = sr.y + sr.h - r.y;
      break;
    case 'b':
      r.h = Math.max(minSize, Math.min(1 - sr.y, sr.h + dy));
      break;
    case 'l':
      r.x = Math.max(0, Math.min(sr.x + sr.w - minSize, sr.x + dx));
      r.w = sr.x + sr.w - r.x;
      break;
    case 'r':
      r.w = Math.max(minSize, Math.min(1 - sr.x, sr.w + dx));
      break;
  }
  cropRect = r;
  drawCrop();
  e.preventDefault();
});

cropCanvas.addEventListener('pointerup', () => { cropDragState = null; });
cropCanvas.addEventListener('pointercancel', () => { cropDragState = null; });

// === CHART SCAN: GRID DETECTION ===
function getCroppedImageData() {
  const ctx = analysisCanvas.getContext('2d');
  const sx = Math.round(cropRect.x * sourceImage.width);
  const sy = Math.round(cropRect.y * sourceImage.height);
  const sw = Math.max(1, Math.round(cropRect.w * sourceImage.width));
  const sh = Math.max(1, Math.round(cropRect.h * sourceImage.height));
  const analysisWidth = Math.min(sw, 1200);
  const scale = analysisWidth / sw;
  const analysisHeight = Math.max(1, Math.round(sh * scale));
  analysisCanvas.width = analysisWidth;
  analysisCanvas.height = analysisHeight;
  ctx.imageSmoothingEnabled = true;
  ctx.drawImage(sourceImage, sx, sy, sw, sh, 0, 0, analysisWidth, analysisHeight);
  return ctx.getImageData(0, 0, analysisWidth, analysisHeight);
}

function computeProjections(imageData) {
  var data = imageData.data, w = imageData.width, h = imageData.height;
  var hProj = new Float32Array(h);
  var vProj = new Float32Array(w);
  for (var y = 0; y < h; y++) {
    var sum = 0;
    for (var x = 0; x < w; x++) {
      var idx = (y * w + x) * 4;
      var lum = 0.299 * data[idx] + 0.587 * data[idx + 1] + 0.114 * data[idx + 2];
      var darkness = 1.0 - lum / 255.0;
      sum += darkness;
      vProj[x] += darkness;
    }
    hProj[y] = sum / w;
  }
  for (var x = 0; x < w; x++) vProj[x] /= h;
  return { hProj: hProj, vProj: vProj };
}

function boxSmooth(arr, radius) {
  var result = new Float32Array(arr.length);
  for (var i = 0; i < arr.length; i++) {
    var sum = 0, count = 0;
    for (var j = Math.max(0, i - radius); j <= Math.min(arr.length - 1, i + radius); j++) {
      sum += arr[j]; count++;
    }
    result[i] = sum / count;
  }
  return result;
}

function findGridLines(profile) {
  var smoothed = boxSmooth(profile, 3);
  var n = smoothed.length;
  if (n < 10) return null;
  var mean = 0;
  for (var i = 0; i < n; i++) mean += smoothed[i];
  mean /= n;
  var variance = 0;
  for (var i = 0; i < n; i++) variance += (smoothed[i] - mean) * (smoothed[i] - mean);
  var stddev = Math.sqrt(variance / n);
  var peakThreshold = mean + 0.5 * stddev;
  var candidates = [];
  for (var i = 2; i < n - 2; i++) {
    if (smoothed[i] > peakThreshold &&
        smoothed[i] >= smoothed[i - 1] && smoothed[i] >= smoothed[i + 1] &&
        smoothed[i] >= smoothed[i - 2] && smoothed[i] >= smoothed[i + 2]) {
      candidates.push(i);
    }
  }
  if (candidates.length < 3) return null;

  var period = estimatePeriodAutocorrelation(smoothed);
  if (!period || period < 5) {
    period = estimatePeriod(candidates);
  }
  if (!period || period < 5) return null;
  return buildRegularGrid(candidates, period, n);
}

function estimatePeriod(candidates) {
  var gaps = [];
  for (var i = 1; i < candidates.length; i++) gaps.push(candidates[i] - candidates[i - 1]);
  if (gaps.length === 0) return null;
  gaps.sort(function(a, b) { return a - b; });
  var median = gaps[Math.floor(gaps.length / 2)];
  var filtered = gaps.filter(function(g) { return Math.abs(g - median) < median * 0.4; });
  if (filtered.length === 0) return median;
  var sum = 0;
  for (var i = 0; i < filtered.length; i++) sum += filtered[i];
  return sum / filtered.length;
}

function estimatePeriodAutocorrelation(profile) {
  var n = profile.length;
  if (n < 20) return null;

  var mean = 0;
  for (var i = 0; i < n; i++) mean += profile[i];
  mean /= n;
  var centered = new Float32Array(n);
  for (var i = 0; i < n; i++) centered[i] = profile[i] - mean;

  var minLag = 5;
  var maxLag = Math.floor(n / 3);
  if (maxLag <= minLag) return null;

  var acorr = new Float32Array(maxLag + 1);
  var ac0 = 0;
  for (var i = 0; i < n; i++) ac0 += centered[i] * centered[i];
  if (ac0 === 0) return null;

  for (var lag = minLag; lag <= maxLag; lag++) {
    var sum = 0;
    for (var i = 0; i < n - lag; i++) {
      sum += centered[i] * centered[i + lag];
    }
    acorr[lag] = sum / ac0;
  }

  var searchStart = minLag;
  for (var i = minLag; i < maxLag; i++) {
    if (acorr[i] < acorr[i + 1]) {
      searchStart = i;
      break;
    }
  }

  var bestLag = -1;
  var bestVal = 0;
  for (var i = searchStart + 1; i < maxLag; i++) {
    if (acorr[i] > bestVal &&
        acorr[i] >= acorr[i - 1] && acorr[i] >= acorr[i + 1]) {
      bestVal = acorr[i];
      bestLag = i;
      if (bestVal > 0.15) break;
    }
  }

  if (bestLag < minLag || bestVal < 0.05) return null;

  if (bestLag > minLag && bestLag < maxLag) {
    var a = acorr[bestLag - 1];
    var b = acorr[bestLag];
    var c = acorr[bestLag + 1];
    var denom = 2 * (2 * b - a - c);
    if (denom !== 0) {
      var offset = (a - c) / denom;
      if (Math.abs(offset) < 0.5) {
        return bestLag + offset;
      }
    }
  }

  return bestLag;
}

function buildRegularGrid(candidates, period, totalLength) {
  var bestLines = null, bestScore = -1;
  var startCount = Math.min(candidates.length, 5);
  for (var s = 0; s < startCount; s++) {
    var lines = [candidates[s]];
    var score = 1;
    var pos = candidates[s] + period;
    while (pos < totalLength - 2) {
      var snap = findNearestCandidate(candidates, Math.round(pos), period * 0.3);
      if (snap !== null) { lines.push(snap); pos = snap + period; score++; }
      else { lines.push(Math.round(pos)); pos += period; }
    }
    pos = candidates[s] - period;
    while (pos > 2) {
      var snap = findNearestCandidate(candidates, Math.round(pos), period * 0.3);
      if (snap !== null) { lines.unshift(snap); pos = snap - period; score++; }
      else { lines.unshift(Math.round(pos)); pos -= period; }
    }
    if (score > bestScore) { bestScore = score; bestLines = lines; }
  }
  return bestLines;
}

function findNearestCandidate(candidates, target, tolerance) {
  var best = null, bestDist = tolerance + 1;
  for (var i = 0; i < candidates.length; i++) {
    var dist = Math.abs(candidates[i] - target);
    if (dist < bestDist) { bestDist = dist; best = candidates[i]; }
  }
  return bestDist <= tolerance ? best : null;
}

function adjustLinesToCount(lines, target, totalLength) {
  if (lines.length === target) return lines;
  var medianGap = getMedianGap(lines);
  var merged = [lines[0]];
  for (var i = 1; i < lines.length; i++) {
    if (lines[i] - merged[merged.length - 1] < medianGap * 0.5) {
      merged[merged.length - 1] = Math.round((merged[merged.length - 1] + lines[i]) / 2);
    } else {
      merged.push(lines[i]);
    }
  }
  if (merged.length < target) {
    var newMedian = getMedianGap(merged);
    var result = [merged[0]];
    for (var i = 1; i < merged.length; i++) {
      var gap = merged[i] - merged[i - 1];
      if (gap > newMedian * 1.5 && result.length < target - (merged.length - i)) {
        var divisions = Math.round(gap / newMedian);
        for (var d = 1; d < divisions; d++) {
          result.push(Math.round(merged[i - 1] + gap * d / divisions));
        }
      }
      result.push(merged[i]);
    }
    merged = result;
  }
  if (merged.length > target) merged = merged.slice(0, target);
  return merged;
}

function getMedianGap(lines) {
  var gaps = [];
  for (var i = 1; i < lines.length; i++) gaps.push(lines[i] - lines[i - 1]);
  gaps.sort(function(a, b) { return a - b; });
  return gaps[Math.floor(gaps.length / 2)] || 1;
}

function computeCellBounds(vLines, hLines) {
  var cells = [];
  for (var r = 0; r < hLines.length - 1; r++) {
    var row = [];
    for (var c = 0; c < vLines.length - 1; c++) {
      row.push({ x: vLines[c], y: hLines[r], w: vLines[c + 1] - vLines[c], h: hLines[r + 1] - hLines[r] });
    }
    cells.push(row);
  }
  return cells;
}

// === CHART SCAN: CENTER SAMPLING ===
function getCellCenterLuminance(imageData, cell, ratio) {
  var data = imageData.data, w = imageData.width;
  var marginX = cell.w * (1 - ratio) / 2;
  var marginY = cell.h * (1 - ratio) / 2;
  var sx = Math.round(cell.x + marginX);
  var sy = Math.round(cell.y + marginY);
  var sw = Math.max(1, Math.round(cell.w * ratio));
  var sh = Math.max(1, Math.round(cell.h * ratio));
  var totalLum = 0, pixelCount = 0;
  for (var py = sy; py < sy + sh && py < imageData.height; py++) {
    for (var px = sx; px < sx + sw && px < w; px++) {
      var idx = (py * w + px) * 4;
      totalLum += 0.299 * data[idx] + 0.587 * data[idx + 1] + 0.114 * data[idx + 2];
      pixelCount++;
    }
  }
  return pixelCount > 0 ? totalLum / pixelCount : 255;
}

function processChartScanGrid(imageData, cells, threshold, invert, ratio) {
  var grid = [];
  for (var r = 0; r < cells.length; r++) {
    var line = '';
    for (var c = 0; c < Math.min(cells[r].length, getStitchCount()); c++) {
      var avgLum = getCellCenterLuminance(imageData, cells[r][c], ratio);
      var isDark = avgLum < threshold;
      var isPunched = invert ? !isDark : isDark;
      line += isPunched ? 'x' : '-';
    }
    line = line.padEnd(getStitchCount(), '-').substring(0, getStitchCount());
    grid.push(line);
  }
  return grid;
}

function otsuThreshold(values) {
  var hist = new Array(256).fill(0);
  for (var i = 0; i < values.length; i++) hist[Math.round(Math.max(0, Math.min(255, values[i])))]++;
  var total = values.length;
  var sumAll = 0;
  for (var i = 0; i < 256; i++) sumAll += i * hist[i];
  var bestThresh = 128, bestVariance = 0;
  var sumBg = 0, countBg = 0;
  for (var t = 0; t < 256; t++) {
    countBg += hist[t];
    if (countBg === 0) continue;
    var countFg = total - countBg;
    if (countFg === 0) break;
    sumBg += t * hist[t];
    var meanBg = sumBg / countBg;
    var meanFg = (sumAll - sumBg) / countFg;
    var bv = countBg * countFg * (meanBg - meanFg) * (meanBg - meanFg);
    if (bv > bestVariance) { bestVariance = bv; bestThresh = t; }
  }
  return bestThresh;
}

// === DETECT GRID & PROCESS BUTTON ===
detectGridBtn.addEventListener('click', function() {
  if (!sourceImage) return;

  var userRows = parseInt(chartRowsInput.value);
  if (!(userRows > 0)) {
    gridStatus.textContent = 'Please enter the number of rows in the chart.';
    gridStatus.style.color = 'var(--error)';
    chartRowsInput.focus();
    return;
  }

  gridStatus.textContent = 'Analyzing image...';
  gridStatus.style.color = 'var(--muted)';
  requestAnimationFrame(function() {
    try {
      var imgData = getCroppedImageData();
      var w = imgData.width, h = imgData.height;

      var vLines = [];
      var sc = getStitchCount();
      var colStep = w / sc;
      for (var i = 0; i <= sc; i++) vLines.push(Math.round(i * colStep));

      var hLines = [];
      var rowStep = h / userRows;
      for (var i = 0; i <= userRows; i++) hLines.push(Math.round(i * rowStep));

      var cells = computeCellBounds(vLines, hLines);
      detectedCells = cells;
      chartScanImageData = imgData;

      var ratio = parseInt(sampleSlider.value) / 100;
      var luminances = [];
      for (var r = 0; r < cells.length; r++) {
        for (var c = 0; c < Math.min(cells[r].length, getStitchCount()); c++) {
          luminances.push(getCellCenterLuminance(imgData, cells[r][c], ratio));
        }
      }

      var autoThresh = otsuThreshold(luminances);
      thresholdSlider.value = autoThresh;
      thresholdValue.textContent = autoThresh;

      var invert = invertToggle.checked;
      var grid = processChartScanGrid(imgData, cells, autoThresh, invert, ratio);
      saveUndoState();
      currentGrid = grid;
      ensureRowColors(grid);
      renderPreview(grid);

      var numRows = cells.length;
      gridStatus.textContent = getStitchCount() + ' cols x ' + numRows + ' rows. Auto-threshold: ' + autoThresh;
      gridStatus.style.color = 'var(--accent)';

      gridOverlayInfo = { img: sourceImage, cols: getStitchCount(), rows: userRows, cropRect: { x: cropRect.x, y: cropRect.y, w: cropRect.w, h: cropRect.h } };
      gridOverlayToggleRow.style.display = '';
      drawGridOverlay();
    } catch (err) {
      gridStatus.textContent = 'Error: ' + err.message + '. Try adjusting the crop.';
      gridStatus.style.color = 'var(--error)';
    }
  });
});

// === PREVIEW ===
function renderPreview(grid) {
  previewGrid.innerHTML = '';
  // Update CSS variable for stitch count
  document.documentElement.style.setProperty('--stitch-count', getStitchCount());
  if (grid.length === 0) {
    rowCount.textContent = '';
    previewHint.style.display = 'none';
    document.getElementById('editRowsToggleRow').style.display = 'none';
    // Hide toolbar tool buttons
    document.getElementById('fullscreenPreviewBtn').style.display = 'none';
    document.getElementById('invertGridBtn').style.display = 'none';
    document.getElementById('exportPngBtn').disabled = true;
    updateEmptyState();
    updateStatusBar();
    return;
  }
  previewHint.style.display = 'block';
  document.getElementById('editRowsToggleRow').style.display = '';

  // Show toolbar tool buttons
  document.getElementById('fullscreenPreviewBtn').style.display = '';
  document.getElementById('invertGridBtn').style.display = '';
  document.getElementById('exportPngBtn').disabled = false;

  var overlayMode = (document.getElementById('overlayModeSelect') || {}).value || 'none';
  previewGrid.classList.toggle('show-col-nums',   overlayMode === 'colnums');
  previewGrid.classList.toggle('show-punch-dots', overlayMode === 'punchdots');

  var editRows = document.getElementById('editRowsToggle').checked;
  previewGrid.classList.toggle('edit-rows', editRows);

  function addGutters(rowIdx) {
    if (!editRows) return;
    var gl = document.createElement('div');
    gl.className = 'row-gutter left';
    gl.innerHTML = '<span>' + (rowIdx + 1) + '</span><button class="row-del-btn" data-del-row="' + rowIdx + '" title="Delete row ' + (rowIdx + 1) + '">\u00d7</button>';
    previewGrid.appendChild(gl);
    return function() {
      if (!editRows) return;
      var gr = document.createElement('div');
      gr.className = 'row-gutter right';
      gr.innerHTML = '<button class="row-add-btn" data-add-after="' + rowIdx + '" title="Add row after ' + (rowIdx + 1) + '">+</button>';
      previewGrid.appendChild(gr);
    };
  }

  // Pattern rows
  for (let r = 0; r < grid.length; r++) {
    var addRight = addGutters(r);
    var rowClr = currentRowColors[r] || null;
    for (let c = 0; c < getStitchCount(); c++) {
      const cell = document.createElement('div');
      const isPunched = grid[r][c] === 'x';
      cell.className = 'cell editable' + (isPunched ? ' punched' : '');
      cell.dataset.row = r;
      cell.dataset.col = c;
      var overrideKey = r + ',' + c;
      if (cellColorOverrides[overrideKey]) {
        cell.style.background = cellColorOverrides[overrideKey];
      } else if (rowClr) {
        cell.style.background = isPunched ? rowClr.a : rowClr.b;
      }
      var numSpan = document.createElement('span');
      numSpan.className = 'col-num';
      numSpan.textContent = c + 1;
      cell.appendChild(numSpan);
      var dotSpan = document.createElement('span');
      dotSpan.className = 'punch-dot';
      cell.appendChild(dotSpan);
      previewGrid.appendChild(cell);
    }
    if (addRight) addRight();
  }

  rowCount.textContent = grid.length + ' rows';

  updateEmptyState();
  updateStatusBar();
  updateSvgExportGuard();
  renderRepeatPreview(grid);
}

// === 5× REPEAT PREVIEW ===
function renderRepeatPreview(grid) {
  var container = document.getElementById('repeatPreview');
  var rpGrid = document.getElementById('repeatPreviewGrid');
  if (!container || !rpGrid) return;
  if (grid.length === 0) {
    container.style.display = 'none';
    return;
  }
  container.style.display = '';
  rpGrid.innerHTML = '';
  var cols = getStitchCount();
  var REPEATS = 5;
  rpGrid.style.gridTemplateColumns = 'repeat(' + (cols * REPEATS) + ', 1fr)';
  for (var r = 0; r < grid.length; r++) {
    var colors = currentRowColors[r] || DEFAULT_COLORS;
    for (var rep = 0; rep < REPEATS; rep++) {
      for (var c = 0; c < cols; c++) {
        var cell = document.createElement('div');
        cell.className = 'rpc';
        var punched = grid[r][c] === 'x';
        var overrideClr = cellColorOverrides[r + ',' + c];
        if (overrideClr) {
          cell.style.background = overrideClr;
        } else {
          cell.style.background = punched ? colors.a : colors.b;
        }
        rpGrid.appendChild(cell);
      }
    }
  }
}

// === EDIT ROWS TOGGLE ===
document.getElementById('editRowsToggle').addEventListener('change', function() {
  renderPreview(currentGrid);
});

// === ROW DELETE / ADD ===
previewGrid.addEventListener('click', function(e) {
  var delBtn = e.target.closest('[data-del-row]');
  if (delBtn) {
    var row = parseInt(delBtn.dataset.delRow);
    if (currentGrid.length <= 1) return;
    saveUndoState();
    currentGrid.splice(row, 1);
    if (currentRowColors.length > row) currentRowColors.splice(row, 1);
    renderPreview(currentGrid);
    // Sync text input if text panel is open
    textInput.value = currentGrid.join('\n');
    generatedSVGString = null;
    svgOutput.innerHTML = '';
    downloadCutBtn.disabled = true; downloadDrawBtn.disabled = true; downloadCombinedBtn.disabled = true; downloadActions.style.display = 'none';
    return;
  }
  var addBtn = e.target.closest('[data-add-after]');
  if (addBtn) {
    var afterRow = parseInt(addBtn.dataset.addAfter);
    saveUndoState();
    currentGrid.splice(afterRow + 1, 0, '-'.repeat(getStitchCount()));
    if (currentRowColors.length > afterRow) {
      var clr = currentRowColors[afterRow];
      currentRowColors.splice(afterRow + 1, 0, { a: clr.a, b: clr.b });
    }
    renderPreview(currentGrid);
    textInput.value = currentGrid.join('\n');
    generatedSVGString = null;
    svgOutput.innerHTML = '';
    downloadCutBtn.disabled = true; downloadDrawBtn.disabled = true; downloadCombinedBtn.disabled = true; downloadActions.style.display = 'none';
    return;
  }
});

// === PREVIEW EDITING ===
function applyPaintToCell(cell) {
  if (!cell || !cell.classList.contains('editable')) return;
  var row = parseInt(cell.dataset.row);
  var col = parseInt(cell.dataset.col);
  var key = row + ',' + col;

  if (activePaintColor && activePaintColor !== 'eraser') {
    cellColorOverrides[key] = activePaintColor;
    cell.style.background = activePaintColor;
  } else if (activePaintColor === 'eraser') {
    delete cellColorOverrides[key];
    var rowClr = currentRowColors[row] || null;
    var ch = currentGrid[row][col];
    if (rowClr) {
      cell.style.background = ch === 'x' ? rowClr.a : rowClr.b;
    } else {
      cell.style.background = '';
    }
  }
}

previewGrid.addEventListener('click', (e) => {
  var cell = e.target;
  if (cell.classList.contains('col-num') || cell.classList.contains('punch-dot')) cell = cell.parentElement;
  if (!cell.classList.contains('editable')) return;

  const row = parseInt(cell.dataset.row);
  const col = parseInt(cell.dataset.col);

  if (activePaintColor) {
    saveUndoState();
    applyPaintToCell(cell);
    return;
  }

  saveUndoState();
  const line = currentGrid[row];
  const newChar = line[col] === 'x' ? '-' : 'x';
  currentGrid[row] = line.substring(0, col) + newChar + line.substring(col + 1);

  cell.classList.toggle('punched');

  var rowClr = currentRowColors[row] || null;
  if (rowClr) {
    cell.style.background = newChar === 'x' ? rowClr.a : rowClr.b;
  }
  var overKey = row + ',' + col;
  if (cellColorOverrides[overKey]) {
    delete cellColorOverrides[overKey];
  }

  // Sync text input (always visible in left panel)
  textInput.value = currentGrid.join('\n');

  generatedSVGString = null;
  svgOutput.innerHTML = '';
  downloadCutBtn.disabled = true; downloadDrawBtn.disabled = true; downloadCombinedBtn.disabled = true; downloadActions.style.display = 'none';
});

// === DRAG-TO-PAINT ===
previewGrid.addEventListener('mousedown', function(e) {
  if (!activePaintColor) return;
  var cell = e.target;
  if (cell.classList.contains('col-num') || cell.classList.contains('punch-dot')) cell = cell.parentElement;
  if (!cell.classList.contains('editable')) return;
  e.preventDefault();
  isPainting = true;
  applyPaintToCell(cell);
});
document.addEventListener('mousemove', function(e) {
  if (!isPainting) return;
  var el = document.elementFromPoint(e.clientX, e.clientY);
  if (el && (el.classList.contains('col-num') || el.classList.contains('punch-dot'))) el = el.parentElement;
  if (el && el.classList.contains('editable')) {
    applyPaintToCell(el);
  }
});
document.addEventListener('mouseup', function() {
  isPainting = false;
});
previewGrid.addEventListener('touchstart', function(e) {
  if (!activePaintColor) return;
  var touch = e.touches[0];
  var el = document.elementFromPoint(touch.clientX, touch.clientY);
  if (el && (el.classList.contains('col-num') || el.classList.contains('punch-dot'))) el = el.parentElement;
  if (el && el.classList.contains('editable')) {
    e.preventDefault();
    isPainting = true;
    applyPaintToCell(el);
  }
}, { passive: false });
previewGrid.addEventListener('touchmove', function(e) {
  if (!isPainting) return;
  e.preventDefault();
  var touch = e.touches[0];
  var el = document.elementFromPoint(touch.clientX, touch.clientY);
  if (el && (el.classList.contains('col-num') || el.classList.contains('punch-dot'))) el = el.parentElement;
  if (el && el.classList.contains('editable')) {
    applyPaintToCell(el);
  }
}, { passive: false });
previewGrid.addEventListener('touchend', function() {
  isPainting = false;
});

// === SVG GENERATION ===
function buildPolygonPoints(h) {
  return [
    '3.0,0', '139.0,0',
    '140.0,1.7', '140.0,20', '142.0,22',
    '142.0,' + fmtNum(h - 22),
    '140.0,' + fmtNum(h - 20),
    '140.0,' + fmtNum(h - 1.7),
    '139.0,' + fmtNum(h),
    '3.0,' + fmtNum(h),
    '2.0,' + fmtNum(h - 1.7),
    '2.0,' + fmtNum(h - 20),
    '0,' + fmtNum(h - 22),
    '0,22', '2.0,20', '2.0,1.7'
  ].join(' ');
}

function fmtNum(n) {
  if (Number.isInteger(n)) return n.toFixed(1);
  return String(n);
}

function circle(cx, cy, r, fill) {
  return '<circle cx="' + fmtNum(cx) + '" cy="' + fmtNum(cy) +
    '" fill="' + (fill || 'white') + '" r="' + fmtNum(r) + '" stroke="black" stroke-width="' + STROKE_W + '" />';
}

function generateSVG(grid, rowTypes, textOpts) {
  const dataRows = grid.length;
  const totalRows = dataRows + 4;
  const h = totalRows * 5.0;

  var showRowNums = textOpts && textOpts.rowNumbers;
  var cardTitle = textOpts && textOpts.title ? textOpts.title.trim() : '';
  var fontFamily = (textOpts && textOpts.fontFamily) || 'Arial, Helvetica, sans-serif';
  var rowNumFontSize = 3;
  var titleFontSize = 5;

  var rowNumX = 131.5;
  var titleX = 7.75;

  var parts = [];

  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="no"?>');
  parts.push('<svg baseProfile="full" height="' + fmtNum(h) + 'mm" preserveAspectRatio="none" version="1.1" viewBox="0 0 ' + fmtNum(CARD_WIDTH) + ' ' + fmtNum(h) + '" width="' + fmtNum(CARD_WIDTH) + 'mm" xmlns="http://www.w3.org/2000/svg" xmlns:ev="http://www.w3.org/2001/xml-events" xmlns:xlink="http://www.w3.org/1999/xlink">');
  parts.push('<defs />');

  parts.push('<g id="cut">');

  parts.push('<polygon fill="white" points="' + buildPolygonPoints(h) + '" stroke="black" stroke-width="' + STROKE_W + '" />');

  for (let i = 0; i < totalRows; i++) {
    const patternY = 2.5 + i * 5.0;

    parts.push(circle(EDGE_GUIDE_X[0], patternY, EDGE_R));

    const isFullRow = (i < 2) || (i >= totalRows - 2);
    var holeFill = 'white';
    if (!isFullRow && rowTypes && rowTypes[i - 2] === 'inverse') {
      holeFill = '#ab96af';
    }
    for (let c = 0; c < getStitchCount(); c++) {
      let punched;
      if (isFullRow) {
        punched = true;
      } else {
        punched = grid[i - 2][c] === 'x';
      }
      if (punched) {
        parts.push(circle(PATTERN_COLS[c], patternY, PATTERN_R, holeFill));
      }
    }

    parts.push(circle(EDGE_GUIDE_X[1], patternY, EDGE_R));

    if (i < totalRows - 1) {
      const transportY = 5.0 + i * 5.0;
      if (i === 0 || i === 2 || i === totalRows - 4 || i === totalRows - 2) {
        for (const tx of TRANSPORT_X) {
          parts.push(circle(tx, transportY, TRANSPORT_R));
        }
      }
    }
  }

  var hasDrawContent = showRowNums || cardTitle;
  if (hasDrawContent) {
    parts.push('<g id="draw">');

    var machineOffset = (textOpts && textOpts.machineOffset) || 0;
    var patternLen = (textOpts && textOpts.preShiftLen) || dataRows;
    var shift = patternLen > 0 ? (machineOffset % patternLen) : 0;
    var row1DrFound = -1;
    if (showRowNums) {
      for (var dr = 0; dr < dataRows - 2; dr++) {
        var rowY = 2.5 + (dr + 2) * 5.0;
        var bottomUp = dataRows - 1 - dr;
        var rowLabel = ((patternLen - shift + bottomUp) % patternLen) + 1;
        if (rowLabel === 1) row1DrFound = dr;
        parts.push('<text x="' + fmtNum(rowNumX) + '" y="' + fmtNum(rowY) + '" font-family="' + fontFamily + '" font-size="' + rowNumFontSize + '" fill="black" dominant-baseline="central" text-anchor="start">' + rowLabel + '</text>');
      }
      if (row1DrFound < 0) {
        for (var dr = dataRows - 2; dr < dataRows; dr++) {
          var bottomUp = dataRows - 1 - dr;
          var rowLabel = ((patternLen - shift + bottomUp) % patternLen) + 1;
          if (rowLabel === 1) { row1DrFound = dr; break; }
        }
      }
      var topLabel = ((patternLen - shift + dataRows - 1) % patternLen) + 1;
      for (var li = 0; li < 2; li++) {
        var leadY = 2.5 + (1 - li) * 5.0;
        var leadLabel = ((topLabel + li) % patternLen) + 1;
        parts.push('<text x="' + fmtNum(rowNumX) + '" y="' + fmtNum(leadY) + '" font-family="' + fontFamily + '" font-size="' + rowNumFontSize + '" fill="black" dominant-baseline="central" text-anchor="start">' + leadLabel + '</text>');
      }
    }

    if (row1DrFound >= 0) {
      var sepY = 2.5 + (row1DrFound + 2) * 5.0 + 2.5;
      parts.push('<line x1="' + fmtNum(EDGE_GUIDE_X[0] - EDGE_R) + '" y1="' + fmtNum(sepY) + '" x2="' + fmtNum(EDGE_GUIDE_X[1] + EDGE_R) + '" y2="' + fmtNum(sepY) + '" stroke="black" stroke-width="0.5" />');
    }

    if (cardTitle) {
      var titleCenterY = h / 2;
      parts.push('<text x="' + fmtNum(titleX) + '" y="' + fmtNum(titleCenterY) + '" font-family="' + fontFamily + '" font-size="' + titleFontSize + '" fill="black" dominant-baseline="central" text-anchor="middle" transform="rotate(-90,' + fmtNum(titleX) + ',' + fmtNum(titleCenterY) + ')">' + escapeXml(cardTitle) + '</text>');
    }

    parts.push('</g>');
  }

  parts.push('</svg>');
  return parts.join('');
}

function escapeXml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function extractSVGLayer(svgString, layerId) {
  var parser = new DOMParser();
  var doc = parser.parseFromString(svgString, 'image/svg+xml');
  var svg = doc.querySelector('svg');
  if (!svg) return svgString;

  var groups = svg.querySelectorAll('g[id]');
  for (var i = 0; i < groups.length; i++) {
    if (groups[i].id !== layerId) {
      groups[i].parentNode.removeChild(groups[i]);
    }
  }

  return '<?xml version="1.0" encoding="UTF-8" standalone="no"?>' + new XMLSerializer().serializeToString(svg);
}

// === ACTIONS ===
function expandJacquard(grid) {
  // Pattern per row pair (1-indexed):
  //   odd row:  inverse, original
  //   even row: original, inverse
  // Which gives the sequence:
  //   row1-inv, row1-orig, row2-orig, row2-inv, row3-inv, row3-orig, row4-orig, row4-inv, ...
  var result = [];
  var types = [];
  function invert(row) {
    var inv = '';
    for (var c = 0; c < row.length; c++) inv += row[c] === 'x' ? '-' : 'x';
    return inv;
  }
  for (var i = 0; i < grid.length; i++) {
    var rowNum = i + 1; // 1-indexed
    var inv = invert(grid[i]);
    if (rowNum % 2 === 1) {
      // odd row: inverse first, then original
      result.push(inv);   types.push('inverse');
      result.push(grid[i]); types.push('original');
    } else {
      // even row: original first, then inverse
      result.push(grid[i]); types.push('original');
      result.push(inv);   types.push('inverse');
    }
  }
  return { grid: result, types: types };
}

function getMachineOffset() {
  const sel = document.getElementById('machineProfile');
  return sel ? parseInt(sel.value) : 0;
}

function shiftGridForMachine(grid, offset, types) {
  if (offset <= 0 || grid.length === 0) return { grid: grid, types: types };
  const shift = offset % grid.length;
  if (shift === 0) return { grid: grid, types: types };
  return {
    grid: grid.slice(shift).concat(grid.slice(0, shift)),
    types: types ? types.slice(shift).concat(types.slice(0, shift)) : null
  };
}

function repeatGrid(grid, times, types) {
  if (times <= 1) return { grid: grid, types: types };
  var resultGrid = [];
  var resultTypes = types ? [] : null;
  for (var i = 0; i < times; i++) {
    resultGrid = resultGrid.concat(grid);
    if (types) resultTypes = resultTypes.concat(types);
  }
  return { grid: resultGrid, types: resultTypes };
}

function getTextOpts() {
  return {
    rowNumbers: document.getElementById('rowNumbersToggle').checked,
    title: document.getElementById('cardTitleInput').value,
    fontFamily: document.getElementById('fontPicker').value
  };
}

generateBtn.addEventListener('click', () => {
  if (currentGrid.length === 0) return;
  var repeats = Math.max(1, Math.min(50, parseInt(document.getElementById('repeatCount').value) || 1));
  var grid, rowTypes = null;
  if (jacquardToggle.checked) {
    var jq = expandJacquard(currentGrid);
    grid = jq.grid;
    rowTypes = jq.types;
  } else {
    grid = currentGrid;
  }
  var rp = repeatGrid(grid, repeats, rowTypes);
  grid = rp.grid;
  rowTypes = rp.types;
  var preShiftLen = grid.length;
  var offset = getMachineOffset();
  var textOpts = getTextOpts();
  textOpts.preShiftLen = preShiftLen;
  textOpts.machineOffset = offset;
  generatedSVGString = generateSVG(grid, rowTypes, textOpts);
  const displaySVG = generatedSVGString.replace(/<\?xml[^?]*\?>/, '');
  svgOutput.innerHTML = displaySVG;

  // Show SVG preview thumbnail in export panel
  var previewWrap = document.getElementById('svgPreviewWrap');
  var previewThumb = document.getElementById('svgPreviewThumb');
  if (previewWrap && previewThumb) {
    previewThumb.innerHTML = displaySVG;
    previewWrap.style.display = '';
  }

  var hasDrawContent = textOpts.rowNumbers || textOpts.title.trim();
  downloadActions.style.display = 'flex';
  downloadCombinedBtn.disabled = false;
  downloadCutBtn.disabled = false;
  downloadDrawBtn.disabled = !hasDrawContent;
});

function downloadSVG(svgStr, suffix) {
  var blob = new Blob([svgStr], { type: 'image/svg+xml' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = 'punchcard-' + suffix + '-' + Date.now() + '.svg';
  a.click();
  URL.revokeObjectURL(url);
}

downloadCombinedBtn.addEventListener('click', () => {
  if (!generatedSVGString) return;
  downloadSVG(generatedSVGString, 'combined');
});

downloadCutBtn.addEventListener('click', () => {
  if (!generatedSVGString) return;
  downloadSVG(extractSVGLayer(generatedSVGString, 'cut'), 'cut');
});

downloadDrawBtn.addEventListener('click', () => {
  if (!generatedSVGString) return;
  downloadSVG(extractSVGLayer(generatedSVGString, 'draw'), 'draw');
});

// === FRAGMENT LIBRARY ===
function buildFragThumb(frag, thumbWidth) {
  const w = thumbWidth || 4;
  const el = document.createElement('div');
  el.className = 'frag-thumb';
  el.style.gridTemplateColumns = 'repeat(' + getStitchCount() + ', ' + w + 'px)';
  for (let r = 0; r < frag.grid.length; r++) {
    for (let c = 0; c < getStitchCount(); c++) {
      const t = document.createElement('div');
      t.className = 't ' + (frag.grid[r][c] === 'x' ? 'p' : 'b');
      t.style.width = w + 'px';
      t.style.height = w + 'px';
      el.appendChild(t);
    }
  }
  return el;
}

function renderFragGrid(catId) {
  const container = document.getElementById('pgFragGrid');
  container.innerHTML = '';
  const query = document.getElementById('pgFragSearch').value.trim().toLowerCase();
  let frags = catId === 'all'
    ? PATTERN_LIB
    : PATTERN_LIB.filter(f => f.cat === catId);
  if (query) frags = frags.filter(f => f.name.toLowerCase().includes(query));

  frags.forEach(frag => {
    const card = document.createElement('div');
    card.className = 'pg-frag-card';
    card.dataset.fragId = frag.id;
    card.title = 'Add ' + frag.name + ' to sequence';
    card.appendChild(buildFragThumb(frag));
    const name = document.createElement('div');
    name.className = 'pg-frag-name';
    const patCols = frag.grid[0] ? frag.grid[0].length : getStitchCount();
    const widthNote = patCols !== getStitchCount() ? ' ⚠\u202f' + patCols + 'col' : '';
    name.textContent = frag.name + ' (' + frag.grid.length + 'r)' + widthNote;
    card.appendChild(name);
    container.appendChild(card);
  });
}

function addToComposition(fragId) {
  pgComposition.push({ fragId: fragId, repeat: 1, colorA: DEFAULT_COLORS.a, colorB: DEFAULT_COLORS.b, offset: 0 });
  renderComposition();
}

function tintThumb(thumbEl, colorA, colorB) {
  thumbEl.querySelectorAll('.t.p').forEach(function(t) { t.style.background = colorA; });
  thumbEl.querySelectorAll('.t.b').forEach(function(t) { t.style.background = colorB; });
}

function renderComposition() {
  const list = document.getElementById('pgCompList');
  const controls = document.getElementById('pgCompControls');
  const addTopBtn = document.getElementById('pgAddTopBtn');
  const addBottomBtn = document.getElementById('pgAddBottomBtn');
  const replaceBtn = document.getElementById('pgReplaceBtn');

  list.innerHTML = '';
  if (pgComposition.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'pg-comp-empty';
    empty.textContent = 'Use "+ Add pattern" above to build a sequence';
    list.appendChild(empty);
    controls.style.display = 'none';
    addTopBtn.disabled = true;
    addBottomBtn.disabled = true;
    replaceBtn.disabled = true;
    return;
  }

  controls.style.display = '';
  addTopBtn.disabled = false;
  addBottomBtn.disabled = false;
  replaceBtn.disabled = false;

  let totalRows = 0;
  pgComposition.forEach((item, idx) => {
    const frag = PATTERN_LIB.find(f => f.id === item.fragId);
    if (!frag) return;
    const rows = frag.grid.length * item.repeat;
    totalRows += rows;

    const el = document.createElement('div');
    el.className = 'pg-comp-item';
    el.dataset.compIdx = idx;
    el.draggable = true;

    // --- Top row: drag handle, thumbnail, name, row badge, remove ---
    const topRow = document.createElement('div');
    topRow.className = 'pg-comp-top';

    const handle = document.createElement('span');
    handle.className = 'pg-comp-drag-handle';
    handle.textContent = '\u2807';
    handle.title = 'Drag to reorder';

    const miniThumb = buildFragThumb(frag, 2);
    miniThumb.style.flexShrink = '0';
    tintThumb(miniThumb, item.colorA, item.colorB);

    const nameEl = document.createElement('span');
    nameEl.className = 'pg-comp-name';
    nameEl.textContent = frag.name;

    const rowBadge = document.createElement('span');
    rowBadge.className = 'pg-comp-rows';
    rowBadge.textContent = rows + 'r';

    const removeBtn = document.createElement('button');
    removeBtn.className = 'pg-comp-remove';
    removeBtn.textContent = '\u00d7';
    removeBtn.title = 'Remove';
    removeBtn.addEventListener('click', () => {
      pgComposition.splice(idx, 1);
      renderComposition();
    });

    topRow.appendChild(handle);
    topRow.appendChild(miniThumb);
    topRow.appendChild(nameEl);
    topRow.appendChild(rowBadge);
    topRow.appendChild(removeBtn);
    el.appendChild(topRow);

    // --- Bottom row: repeat, offset, colors ---
    const bottomRow = document.createElement('div');
    bottomRow.className = 'pg-comp-bottom';

    // Repeat
    const repeatWrap = document.createElement('label');
    repeatWrap.className = 'pg-comp-field';
    repeatWrap.textContent = 'Repeat ';
    const repeatInput = document.createElement('input');
    repeatInput.type = 'number';
    repeatInput.className = 'pg-comp-repeat';
    repeatInput.min = 1;
    repeatInput.max = 50;
    repeatInput.value = item.repeat;
    repeatInput.addEventListener('change', () => {
      item.repeat = Math.max(1, Math.min(50, parseInt(repeatInput.value) || 1));
      renderComposition();
    });
    repeatWrap.appendChild(repeatInput);
    bottomRow.appendChild(repeatWrap);

    // Horizontal offset
    const offsetWrap = document.createElement('label');
    offsetWrap.className = 'pg-comp-field';
    offsetWrap.title = 'Shift columns horizontally by this many stitches';
    offsetWrap.textContent = 'H-offset ';
    const offsetInput = document.createElement('input');
    offsetInput.type = 'number';
    offsetInput.className = 'pg-comp-offset-input';
    offsetInput.min = 0;
    offsetInput.max = getStitchCount();
    offsetInput.value = item.offset || 0;
    offsetInput.addEventListener('change', () => {
      var max = getStitchCount();
      var val = parseInt(offsetInput.value) || 0;
      item.offset = Math.max(0, Math.min(max, val));
      offsetInput.value = item.offset;
    });
    offsetWrap.appendChild(offsetInput);
    bottomRow.appendChild(offsetWrap);

    // Color A
    const colorAInput = document.createElement('input');
    colorAInput.type = 'color';
    colorAInput.className = 'pg-comp-color';
    colorAInput.value = item.colorA;
    colorAInput.title = 'Color A (punched)';
    colorAInput.addEventListener('input', function() {
      item.colorA = this.value;
      tintThumb(miniThumb, item.colorA, item.colorB);
    });
    colorAInput.addEventListener('change', function() {
      item.colorA = this.value;
      tintThumb(miniThumb, item.colorA, item.colorB);
    });
    var colorAWrap = document.createElement('div');
    colorAWrap.style.cssText = 'position:relative;display:inline-block;flex-shrink:0';
    var colorABtn = document.createElement('div');
    colorABtn.style.cssText = 'width:20px;height:20px;border-radius:4px;border:1px solid var(--copper);cursor:pointer;background:' + item.colorA;
    colorABtn.title = 'Color A (punched) \u2014 click for palette';
    colorABtn.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      if (pgPalette.length > 0) {
        showColorPopover(colorAInput, item.colorA, function(hex) {
          item.colorA = hex;
          colorAInput.value = hex;
          colorABtn.style.background = hex;
          tintThumb(miniThumb, item.colorA, item.colorB);
        });
      } else {
        colorAInput.click();
      }
    });
    colorAInput.addEventListener('input', function() {
      colorABtn.style.background = this.value;
    });
    colorAInput.style.cssText = 'position:absolute;opacity:0;pointer-events:none;width:0;height:0';
    colorAWrap.appendChild(colorABtn);
    colorAWrap.appendChild(colorAInput);
    bottomRow.appendChild(colorAWrap);

    // Color B
    const colorBInput = document.createElement('input');
    colorBInput.type = 'color';
    colorBInput.className = 'pg-comp-color';
    colorBInput.value = item.colorB;
    colorBInput.title = 'Color B (blank)';
    colorBInput.addEventListener('input', function() {
      item.colorB = this.value;
      tintThumb(miniThumb, item.colorA, item.colorB);
    });
    colorBInput.addEventListener('change', function() {
      item.colorB = this.value;
      tintThumb(miniThumb, item.colorA, item.colorB);
    });
    var colorBWrap = document.createElement('div');
    colorBWrap.style.cssText = 'position:relative;display:inline-block;flex-shrink:0';
    var colorBBtn = document.createElement('div');
    colorBBtn.style.cssText = 'width:20px;height:20px;border-radius:4px;border:1px solid var(--copper);cursor:pointer;background:' + item.colorB;
    colorBBtn.title = 'Color B (blank) \u2014 click for palette';
    colorBBtn.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      if (pgPalette.length > 0) {
        showColorPopover(colorBInput, item.colorB, function(hex) {
          item.colorB = hex;
          colorBInput.value = hex;
          colorBBtn.style.background = hex;
          tintThumb(miniThumb, item.colorA, item.colorB);
        });
      } else {
        colorBInput.click();
      }
    });
    colorBInput.addEventListener('input', function() {
      colorBBtn.style.background = this.value;
    });
    colorBInput.style.cssText = 'position:absolute;opacity:0;pointer-events:none;width:0;height:0';
    colorBWrap.appendChild(colorBBtn);
    colorBWrap.appendChild(colorBInput);
    bottomRow.appendChild(colorBWrap);

    el.appendChild(bottomRow);

    list.appendChild(el);
  });

  document.getElementById('pgTotalRows').textContent = totalRows + ' total rows';
}

// === COMPOSITION DRAG-AND-DROP REORDER ===
(function initCompDrag() {
  var list = document.getElementById('pgCompList');
  var dragSrcIdx = -1;

  list.addEventListener('dragstart', function(e) {
    var item = e.target.closest('.pg-comp-item');
    if (!item) return;
    dragSrcIdx = parseInt(item.dataset.compIdx);
    item.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', dragSrcIdx);
  });
  list.addEventListener('dragend', function(e) {
    var item = e.target.closest('.pg-comp-item');
    if (item) item.classList.remove('dragging');
    list.querySelectorAll('.drag-over').forEach(function(el) { el.classList.remove('drag-over'); });
    dragSrcIdx = -1;
  });
  list.addEventListener('dragover', function(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    var item = e.target.closest('.pg-comp-item');
    if (!item || parseInt(item.dataset.compIdx) === dragSrcIdx) return;
    list.querySelectorAll('.drag-over').forEach(function(el) { el.classList.remove('drag-over'); });
    item.classList.add('drag-over');
  });
  list.addEventListener('dragleave', function(e) {
    var item = e.target.closest('.pg-comp-item');
    if (item) item.classList.remove('drag-over');
  });
  list.addEventListener('drop', function(e) {
    e.preventDefault();
    var targetItem = e.target.closest('.pg-comp-item');
    if (!targetItem) return;
    var fromIdx = parseInt(e.dataTransfer.getData('text/plain'));
    var toIdx = parseInt(targetItem.dataset.compIdx);
    if (isNaN(fromIdx) || isNaN(toIdx) || fromIdx === toIdx) return;
    var moved = pgComposition.splice(fromIdx, 1)[0];
    pgComposition.splice(toIdx, 0, moved);
    renderComposition();
  });

  var touchDragIdx = -1;
  var touchClone = null;
  var touchCurrentTarget = null;

  list.addEventListener('touchstart', function(e) {
    var handle = e.target.closest('.pg-comp-drag-handle');
    if (!handle) return;
    var item = handle.closest('.pg-comp-item');
    if (!item) return;
    touchDragIdx = parseInt(item.dataset.compIdx);
    item.classList.add('dragging');
    var rect = item.getBoundingClientRect();
    touchClone = item.cloneNode(true);
    touchClone.style.cssText = 'position:fixed;left:' + rect.left + 'px;top:' + rect.top + 'px;width:' + rect.width + 'px;opacity:0.8;pointer-events:none;z-index:9999;box-shadow:0 4px 12px rgba(0,0,0,0.2);';
    document.body.appendChild(touchClone);
  }, { passive: true });

  list.addEventListener('touchmove', function(e) {
    if (touchDragIdx < 0 || !touchClone) return;
    e.preventDefault();
    var touch = e.touches[0];
    touchClone.style.left = (touch.clientX - touchClone.offsetWidth / 2) + 'px';
    touchClone.style.top = (touch.clientY - touchClone.offsetHeight / 2) + 'px';
    var elUnder = document.elementFromPoint(touch.clientX, touch.clientY);
    var targetItem = elUnder ? elUnder.closest('.pg-comp-item') : null;
    list.querySelectorAll('.drag-over').forEach(function(el) { el.classList.remove('drag-over'); });
    if (targetItem && parseInt(targetItem.dataset.compIdx) !== touchDragIdx) {
      targetItem.classList.add('drag-over');
      touchCurrentTarget = targetItem;
    } else {
      touchCurrentTarget = null;
    }
  }, { passive: false });

  list.addEventListener('touchend', function(e) {
    if (touchDragIdx < 0) return;
    list.querySelectorAll('.dragging').forEach(function(el) { el.classList.remove('dragging'); });
    list.querySelectorAll('.drag-over').forEach(function(el) { el.classList.remove('drag-over'); });
    if (touchClone) { touchClone.remove(); touchClone = null; }
    if (touchCurrentTarget) {
      var toIdx = parseInt(touchCurrentTarget.dataset.compIdx);
      if (!isNaN(toIdx) && toIdx !== touchDragIdx) {
        var moved = pgComposition.splice(touchDragIdx, 1)[0];
        pgComposition.splice(toIdx, 0, moved);
        renderComposition();
      }
    }
    touchDragIdx = -1;
    touchCurrentTarget = null;
  }, { passive: true });
})();

// === PALETTE MANAGEMENT ===
function loadPalette() {
  try {
    var stored = localStorage.getItem(PG_PALETTE_KEY);
    if (stored) {
      pgPalette = JSON.parse(stored);
    } else {
      pgPalette = DEFAULT_PALETTE.slice();
      savePalette();
    }
  } catch (e) { pgPalette = DEFAULT_PALETTE.slice(); }
}
function savePalette() {
  try { localStorage.setItem(PG_PALETTE_KEY, JSON.stringify(pgPalette)); } catch (e) {}
}
function renderPalette() {
  var container = document.getElementById('pgPaletteSwatches');
  var empty = document.getElementById('pgPaletteEmpty');
  var clearBtn = document.getElementById('pgPaletteClearBtn');
  container.innerHTML = '';

  if (pgPalette.length === 0) {
    var emptyEl = document.createElement('span');
    emptyEl.className = 'pg-palette-empty';
    emptyEl.id = 'pgPaletteEmpty';
    emptyEl.textContent = 'Add colors to build your palette';
    container.appendChild(emptyEl);
    clearBtn.style.display = 'none';
    activePaintColor = null;
    previewGrid.classList.remove('paint-mode');
    updateActivePaintChip();
    return;
  }
  clearBtn.style.display = '';

  // If no active color or active color removed, default to first
  if (!activePaintColor || (activePaintColor !== 'eraser' && pgPalette.indexOf(activePaintColor) === -1)) {
    activePaintColor = pgPalette[0];
  }

  pgPalette.forEach(function(hex, idx) {
    var swatch = document.createElement('div');
    swatch.className = 'pg-palette-swatch' + (hex === activePaintColor ? ' active' : '');
    swatch.style.background = hex;
    swatch.title = hex;
    swatch.addEventListener('click', function(e) {
      if (e.target.classList.contains('pg-swatch-remove')) return;
      activePaintColor = hex;
      previewGrid.classList.add('paint-mode');
      container.querySelectorAll('.pg-palette-swatch, .pg-palette-eraser').forEach(function(s) { s.classList.remove('active'); });
      swatch.classList.add('active');
      updateActivePaintChip();
    });
    var removeBtn = document.createElement('button');
    removeBtn.className = 'pg-swatch-remove';
    removeBtn.textContent = '\u00d7';
    removeBtn.addEventListener('click', function(e) {
      e.stopPropagation();
      pgPalette.splice(idx, 1);
      savePalette();
      renderPalette();
    });
    swatch.appendChild(removeBtn);
    container.appendChild(swatch);
  });

  // Eraser swatch at the end
  var eraser = document.createElement('div');
  eraser.className = 'pg-palette-eraser' + (activePaintColor === 'eraser' ? ' active' : '');
  eraser.title = 'Eraser';
  eraser.textContent = '🧹';
  eraser.addEventListener('click', function() {
    activePaintColor = 'eraser';
    container.querySelectorAll('.pg-palette-swatch, .pg-palette-eraser').forEach(function(s) { s.classList.remove('active'); });
    eraser.classList.add('active');
    updateActivePaintChip();
  });
  container.appendChild(eraser);

  // Ensure paint mode is on when palette has colors
  if (activePaintColor) {
    previewGrid.classList.add('paint-mode');
  }
  updateActivePaintChip();
}
function parsePaletteInput(raw) {
  var colors = [];
  var trimmed = raw.trim();
  if (!trimmed) return colors;

  var coolorsMatch = trimmed.match(/coolors\.co\/(?:palette\/)?([0-9a-f]{3,8}(?:-[0-9a-f]{3,8})+)/i);
  if (coolorsMatch) {
    var parts = coolorsMatch[1].split('-');
    parts.forEach(function(p) {
      var hex = '#' + p.substring(0, 6).toLowerCase();
      if (/^#[0-9a-f]{6}$/.test(hex) && colors.indexOf(hex) === -1) {
        colors.push(hex);
      }
    });
    return colors;
  }

  var hexRegex = /#([0-9a-f]{3,8})\b/gi;
  var match;
  while ((match = hexRegex.exec(trimmed)) !== null) {
    var val = match[1].toLowerCase();
    var hex;
    if (val.length === 3) {
      hex = '#' + val[0] + val[0] + val[1] + val[1] + val[2] + val[2];
    } else if (val.length >= 6) {
      hex = '#' + val.substring(0, 6);
    } else {
      continue;
    }
    if (/^#[0-9a-f]{6}$/.test(hex) && colors.indexOf(hex) === -1) {
      colors.push(hex);
    }
  }
  if (colors.length > 0) return colors;

  var bareRegex = /\b([0-9a-f]{6})\b/gi;
  while ((match = bareRegex.exec(trimmed)) !== null) {
    var h = '#' + match[1].toLowerCase();
    if (colors.indexOf(h) === -1) colors.push(h);
  }

  return colors;
}

// === PAINT MODE ===
// Paint mode is always on when palette has colors — no open/close needed

function initPalette() {
  loadPalette();
  renderPalette();

  var addBtn = document.getElementById('pgPaletteAddBtn');
  var colorInput = document.getElementById('pgPaletteColorInput');

  addBtn.addEventListener('click', function() {
    colorInput.click();
  });

  document.getElementById('pgPaletteClearBtn').addEventListener('click', function() {
    pgPalette = [];
    savePalette();
    renderPalette();
  });
  colorInput.addEventListener('input', function() {
    // Live preview placeholder
  });
  colorInput.addEventListener('change', function() {
    var hex = this.value;
    if (pgPalette.indexOf(hex) === -1) {
      pgPalette.push(hex);
      savePalette();
      renderPalette();
    }
  });

  var importToggle = document.getElementById('pgImportToggle');
  var importArea = document.getElementById('pgImportArea');
  importToggle.addEventListener('click', function() {
    importArea.classList.toggle('open');
  });

  var importBtn = document.getElementById('pgImportBtn');
  var importInput = document.getElementById('pgImportInput');
  var importMsg = document.getElementById('pgImportMsg');

  importBtn.addEventListener('click', function() {
    var raw = importInput.value;
    var parsed = parsePaletteInput(raw);
    if (parsed.length === 0) {
      importMsg.style.color = 'var(--error)';
      importMsg.textContent = 'No colors found \u2014 try a Coolors URL or CSS hex values';
      return;
    }
    var added = 0;
    parsed.forEach(function(hex) {
      if (pgPalette.indexOf(hex) === -1) {
        pgPalette.push(hex);
        added++;
      }
    });
    savePalette();
    renderPalette();
    importInput.value = '';
    importMsg.style.color = 'var(--accent)';
    if (added === 0) {
      importMsg.textContent = parsed.length + ' color(s) already in palette';
    } else {
      importMsg.textContent = 'Added ' + added + ' color' + (added > 1 ? 's' : '') + '!';
    }
    setTimeout(function() { importMsg.textContent = ''; }, 3000);
  });
}

// === PALETTE POPOVER FOR COMP COLOR PICKERS ===
var activePopover = null;
function closePopover() {
  if (activePopover) {
    activePopover.remove();
    activePopover = null;
  }
}
function showColorPopover(anchorEl, currentValue, onSelect) {
  closePopover();
  if (pgPalette.length === 0) return;

  var pop = document.createElement('div');
  pop.className = 'pg-color-popover';

  pgPalette.forEach(function(hex) {
    var sw = document.createElement('div');
    sw.className = 'pg-pop-swatch';
    sw.style.background = hex;
    sw.title = hex;
    if (hex === currentValue) {
      sw.style.borderColor = 'var(--primary)';
      sw.style.borderWidth = '2.5px';
    }
    sw.addEventListener('click', function(e) {
      e.stopPropagation();
      onSelect(hex);
      closePopover();
    });
    pop.appendChild(sw);
  });

  var customBtn = document.createElement('div');
  customBtn.className = 'pg-pop-custom';
  customBtn.title = 'Custom color...';
  customBtn.textContent = '\u270e';
  customBtn.addEventListener('click', function(e) {
    e.stopPropagation();
    closePopover();
    anchorEl.click();
  });
  pop.appendChild(customBtn);

  var rect = anchorEl.getBoundingClientRect();
  pop.style.position = 'fixed';
  pop.style.left = Math.min(rect.left, window.innerWidth - 190) + 'px';
  pop.style.top = (rect.bottom + 4) + 'px';
  document.body.appendChild(pop);
  activePopover = pop;

  setTimeout(function() {
    function handler(e) {
      if (pop.parentNode && !pop.contains(e.target) && e.target !== anchorEl) {
        closePopover();
        document.removeEventListener('click', handler, true);
      }
    }
    document.addEventListener('click', handler, true);
  }, 0);
}

function initLibrary() {
  // === Build category filter buttons ===
  const catFilter = document.getElementById('pgCatFilter');
  PATTERN_CATS.forEach(cat => {
    const btn = document.createElement('button');
    btn.className = 'pg-cat-btn';
    btn.dataset.cat = cat.id;
    btn.textContent = cat.name;
    catFilter.appendChild(btn);
  });

  // Category filter clicks
  catFilter.addEventListener('click', e => {
    const btn = e.target.closest('.pg-cat-btn');
    if (!btn) return;
    catFilter.querySelectorAll('.pg-cat-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentCatId = btn.dataset.cat;
    renderFragGrid(btn.dataset.cat);
  });

  // Search filter
  document.getElementById('pgFragSearch').addEventListener('input', () => {
    const activeCat = catFilter.querySelector('.pg-cat-btn.active');
    renderFragGrid(activeCat ? activeCat.dataset.cat : 'all');
  });

  // === Fragment card click → add to composition ===
  document.getElementById('pgFragGrid').addEventListener('click', e => {
    const card = e.target.closest('.pg-frag-card');
    if (!card) return;
    const fragId = card.dataset.fragId;
    addToComposition(fragId);
  });

  // === Composition controls ===
  // Clear all
  document.getElementById('pgClearAll').addEventListener('click', () => {
    pgComposition = [];
    renderComposition();
  });

  // Build rows from composition (with optional half-offset)
  function buildCompositionRows() {
    const rows = [];
    const colors = [];
    const stitches = getStitchCount();
    pgComposition.forEach(item => {
      const frag = PATTERN_LIB.find(f => f.id === item.fragId);
      if (!frag) return;
      const off = (parseInt(item.offset) || 0) % stitches;
      for (let r = 0; r < item.repeat; r++) {
        frag.grid.forEach(row => {
          let r = off > 0 ? row.substring(off) + row.substring(0, off) : row;
          if (r.length < stitches) r = r + '-'.repeat(stitches - r.length);
          else if (r.length > stitches) r = r.substring(0, stitches);
          rows.push(r);
          colors.push({ a: item.colorA, b: item.colorB });
        });
      }
    });
    return { rows, colors };
  }

  function afterAppend() {
    isPainting = false;
    renderPreview(currentGrid);
    // generateBtn state is managed by updateSvgExportGuard() called from renderPreview
    generatedSVGString = null;
    svgOutput.innerHTML = '';
    downloadCutBtn.disabled = true;
    downloadDrawBtn.disabled = true;
    downloadCombinedBtn.disabled = true;
    downloadActions.style.display = 'none';
  }

  // Add to bottom
  document.getElementById('pgAddBottomBtn').addEventListener('click', () => {
    const { rows, colors } = buildCompositionRows();
    if (rows.length === 0) return;
    saveUndoState();
    currentGrid = currentGrid.concat(rows);
    currentRowColors = currentRowColors.concat(colors);
    afterAppend();
  });

  // Add to top
  document.getElementById('pgAddTopBtn').addEventListener('click', () => {
    const { rows, colors } = buildCompositionRows();
    if (rows.length === 0) return;
    saveUndoState();
    // Shift existing cell overrides down by the number of new rows
    const shift = rows.length;
    const shifted = {};
    for (const key in cellColorOverrides) {
      const parts = key.split(',');
      const newRow = parseInt(parts[0], 10) + shift;
      shifted[newRow + ',' + parts[1]] = cellColorOverrides[key];
    }
    cellColorOverrides = shifted;
    currentGrid = rows.concat(currentGrid);
    currentRowColors = colors.concat(currentRowColors);
    afterAppend();
  });

  // Replace canvas
  document.getElementById('pgReplaceBtn').addEventListener('click', () => {
    const { rows, colors } = buildCompositionRows();
    if (rows.length === 0) return;
    saveUndoState();
    currentGrid = rows;
    currentRowColors = colors;
    cellColorOverrides = {};
    afterAppend();
  });

  // Initial render
  renderFragGrid('all');
}

// === INIT ===
initLibrary();
initPalette();

// Initialize with 20-row blank grid
(function initBlankGrid() {
  if (currentGrid.length === 0) {
    currentGrid = Array.from({ length: 20 }, () => '-'.repeat(getStitchCount()));
    ensureRowColors(currentGrid);
    renderPreview(currentGrid);
  }
})();
updateEmptyState();

// === PNG EXPORT ===
document.getElementById('exportPngBtn').addEventListener('click', function() {
  var CELL = 12;
  var cols = getStitchCount();
  var rows = currentGrid.length;
  if (rows === 0) return;

  var canvas = document.createElement('canvas');
  canvas.width = cols * CELL;
  canvas.height = rows * CELL;
  var ctx = canvas.getContext('2d');

  for (var r = 0; r < rows; r++) {
    var colors = currentRowColors[r] || DEFAULT_COLORS;
    for (var c = 0; c < cols; c++) {
      var punched = currentGrid[r][c] === 'x';
      var overrideClr = cellColorOverrides[r + ',' + c];
      if (overrideClr) {
        ctx.fillStyle = overrideClr;
      } else {
        ctx.fillStyle = punched ? colors.a : colors.b;
      }
      ctx.fillRect(c * CELL, r * CELL, CELL, CELL);
      ctx.strokeStyle = 'rgba(0,0,0,0.08)';
      ctx.lineWidth = 0.5;
      ctx.strokeRect(c * CELL, r * CELL, CELL, CELL);
    }
  }

  canvas.toBlob(function(blob) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'punchcard-preview.png';
    a.click();
    setTimeout(function() { URL.revokeObjectURL(url); }, 1000);
  }, 'image/png');
});

// === FULLSCREEN LIGHTBOX ===
document.getElementById('fullscreenPreviewBtn').addEventListener('click', function() {
  if (currentGrid.length === 0) return;
  var overlay = document.getElementById('lightboxOverlay');
  var content = document.getElementById('lightboxContent');
  content.innerHTML = '';

  var REPEATS = 5;
  for (var r = 0; r < currentGrid.length; r++) {
    var colors = currentRowColors[r] || DEFAULT_COLORS;
    for (var rep = 0; rep < REPEATS; rep++) {
      for (var c = 0; c < getStitchCount(); c++) {
        var cell = document.createElement('div');
        cell.className = 'lbc';
        var punched = currentGrid[r][c] === 'x';
        var lbOverride = cellColorOverrides[r + ',' + c];
        if (lbOverride) {
          cell.style.background = lbOverride;
        } else {
          cell.style.background = punched ? colors.a : colors.b;
        }
        content.appendChild(cell);
      }
    }
  }

  overlay.style.display = 'flex';
});

document.getElementById('lightboxOverlay').addEventListener('click', function(e) {
  if (e.target === e.currentTarget) {
    e.currentTarget.style.display = 'none';
  }
});

document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape') {
    var overlay = document.getElementById('lightboxOverlay');
    if (overlay.style.display !== 'none') {
      overlay.style.display = 'none';
    }
    // No sidebar to close — panels are always visible
  }
  // Undo: Ctrl+Z (or Cmd+Z on Mac)
  if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
    e.preventDefault();
    undo();
  }
  // Redo: Ctrl+Shift+Z (or Cmd+Shift+Z on Mac)
  if ((e.ctrlKey || e.metaKey) && e.key === 'z' && e.shiftKey) {
    e.preventDefault();
    redo();
  }
  // Redo alt: Ctrl+Y
  if ((e.ctrlKey || e.metaKey) && e.key === 'y') {
    e.preventDefault();
    redo();
  }
});

// Wire up undo/redo toolbar buttons
document.getElementById('undoBtn').addEventListener('click', undo);
document.getElementById('redoBtn').addEventListener('click', redo);

// === INVERT GRID ===
document.getElementById('invertGridBtn').addEventListener('click', () => {
  if (currentGrid.length === 0) return;
  saveUndoState();
  currentGrid = currentGrid.map(row =>
    row.split('').map(ch => ch === 'x' ? '-' : 'x').join('')
  );
  cellColorOverrides = {};
  renderPreview(currentGrid);
  generatedSVGString = null;
  svgOutput.innerHTML = '';
  downloadCutBtn.disabled = true;
  downloadDrawBtn.disabled = true;
  downloadCombinedBtn.disabled = true;
  downloadActions.style.display = 'none';
});

// === RESIZABLE PANELS ===
(function initResizablePanels() {
  const leftPanel = document.getElementById('leftPanel');
  const rightPanel = document.getElementById('rightPanel');
  const resizeLeft = document.getElementById('resizeLeft');
  const resizeRight = document.getElementById('resizeRight');
  const MIN_PANEL = 180;
  const MAX_PANEL = 500;

  function startDrag(handleEl, panel, side) {
    return function (downEvt) {
      if (window.matchMedia('(max-width: 800px)').matches) return;
      downEvt.preventDefault();
      const startX = downEvt.clientX || (downEvt.touches && downEvt.touches[0].clientX);
      const startW = panel.getBoundingClientRect().width;
      handleEl.classList.add('dragging');
      document.body.style.cursor = 'col-resize';
      document.body.style.userSelect = 'none';

      function onMove(moveEvt) {
        const cx = moveEvt.clientX || (moveEvt.touches && moveEvt.touches[0].clientX);
        let delta = cx - startX;
        if (side === 'right') delta = -delta;  // right panel grows leftward
        const newW = Math.min(MAX_PANEL, Math.max(MIN_PANEL, startW + delta));
        panel.style.width = newW + 'px';
      }

      function onUp() {
        handleEl.classList.remove('dragging');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        document.removeEventListener('touchmove', onMove);
        document.removeEventListener('touchend', onUp);
      }

      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
      document.addEventListener('touchmove', onMove, { passive: false });
      document.addEventListener('touchend', onUp);
    };
  }

  resizeLeft.addEventListener('mousedown', startDrag(resizeLeft, leftPanel, 'left'));
  resizeLeft.addEventListener('touchstart', startDrag(resizeLeft, leftPanel, 'left'), { passive: false });
  resizeRight.addEventListener('mousedown', startDrag(resizeRight, rightPanel, 'right'));
  resizeRight.addEventListener('touchstart', startDrag(resizeRight, rightPanel, 'right'), { passive: false });
})();

// === PWA ===
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('./sw.js')
    .then((reg) => console.log('SW registered:', reg.scope))
    .catch((err) => console.log('SW registration failed:', err));
}
