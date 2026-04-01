const dropZone   = document.getElementById('drop-zone');
const fileInput  = document.getElementById('docx-input');
const fileNameEl = document.getElementById('file-name');
const startInput = document.getElementById('start-fn');
const endInput   = document.getElementById('end-fn');
const submitBtn  = document.getElementById('submit-btn');
const btnText    = document.getElementById('btn-text');
const btnSpinner = document.getElementById('btn-spinner');
const xlsxBtn    = document.getElementById('xlsx-btn');
const xlsxText   = document.getElementById('xlsx-text');
const xlsxSpinner = document.getElementById('xlsx-spinner');
const errorMsg   = document.getElementById('error-msg');
const form       = document.getElementById('upload-form');

let selectedFile = null;

// ── File selection ─────────────────────────────────────────────────────────

function setFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.docx')) {
    showError('Please select a .docx file.');
    return;
  }
  selectedFile = file;
  fileNameEl.textContent = '\uD83D\uDCCE ' + file.name;
  dropZone.classList.add('has-file');
  dropZone.classList.remove('drag-over');
  hideError();
  updateButton();
}

dropZone.addEventListener('click', (e) => {
  // The label[for="docx-input"] already opens the picker natively — skip to
  // avoid triggering two file dialogs.
  if (e.target === fileInput || e.target.closest('label')) return;
  fileInput.click();
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) setFile(fileInput.files[0]);
});

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  const file = e.dataTransfer.files[0];
  if (file) setFile(file);
});

// ── Validation ─────────────────────────────────────────────────────────────

function rangeValid() {
  const s = parseInt(startInput.value, 10);
  const e = parseInt(endInput.value, 10);
  return Number.isFinite(s) && Number.isFinite(e) && s >= 1 && e >= s;
}

function updateButton() {
  const ready = !!(selectedFile && rangeValid());
  submitBtn.disabled = !ready;
  xlsxBtn.disabled   = !ready;
}

startInput.addEventListener('input', updateButton);
endInput.addEventListener('input',   updateButton);

// ── Submission ─────────────────────────────────────────────────────────────

async function submitTo(endpoint, ext, setLoadingFn) {
  if (!selectedFile || !rangeValid()) return;

  const s  = parseInt(startInput.value, 10);
  const en = parseInt(endInput.value, 10);

  setLoadingFn(true);
  hideError();

  const data = new FormData();
  data.append('docx',     selectedFile);
  data.append('start_fn', s);
  data.append('end_fn',   en);

  try {
    const res = await fetch(endpoint, { method: 'POST', body: data });

    if (res.ok) {
      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      const stem = selectedFile.name.replace(/\.docx$/i, '');
      a.href     = url;
      a.download = `${stem}_footnotes_${s}-${en}.${ext}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } else {
      let msg = 'Processing failed. Please try again.';
      try { msg = (await res.json()).error || msg; } catch (_) {}
      showError(msg);
    }
  } catch (err) {
    showError('Network error: ' + err.message);
  } finally {
    setLoadingFn(false);
  }
}

form.addEventListener('submit', (e) => {
  e.preventDefault();
  submitTo('/process', 'pdf', setPdfLoading);
});

xlsxBtn.addEventListener('click', () => {
  submitTo('/process_xlsx', 'xlsx', setXlsxLoading);
});

// ── Helpers ────────────────────────────────────────────────────────────────

function setPdfLoading(on) {
  submitBtn.disabled  = on;
  xlsxBtn.disabled    = on;
  btnText.textContent = on ? 'Processing\u2026' : 'Generate PDF';
  btnSpinner.hidden   = !on;
}

function setXlsxLoading(on) {
  submitBtn.disabled   = on;
  xlsxBtn.disabled     = on;
  xlsxText.textContent = on ? 'Processing\u2026' : 'Export Excel';
  xlsxSpinner.hidden   = !on;
}

// keep old name as alias for any remaining references
const setLoading = setPdfLoading;

function showError(msg) {
  errorMsg.textContent = msg;
  errorMsg.hidden = false;
}

function hideError() {
  errorMsg.hidden = true;
}
