import { applyLetterhead } from './docx-letterhead.js';

const state = {
  file: null,
  templates: [],
  selectedTemplate: null
};

const views = {
  home: document.getElementById('view-home'),
  letterhead: document.getElementById('view-letterhead')
};

const fileInput = document.getElementById('source-file');
const fileName = document.getElementById('file-name');
const dropZone = document.getElementById('drop-zone');
const templateList = document.getElementById('template-list');
const generateButton = document.getElementById('generate-button');
const statusBox = document.getElementById('status-box');

function route(name) {
  Object.entries(views).forEach(([key, view]) => view.classList.toggle('is-active', key === name));
  history.replaceState(null, '', name === 'home' ? '#/' : '#/carta-intestata');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

document.querySelectorAll('[data-route]').forEach(el => {
  el.addEventListener('click', () => route(el.dataset.route));
});

function showStatus(message, type = '') {
  statusBox.hidden = !message;
  statusBox.textContent = message || '';
  statusBox.className = `status-box${type ? ` is-${type}` : ''}`;
}

function validateFile(file) {
  if (!file) return false;
  if (!file.name.toLowerCase().endsWith('.docx')) {
    showStatus('Seleziona un file Word in formato .docx. I vecchi file .doc non sono supportati.', 'error');
    return false;
  }
  return true;
}

function setFile(file) {
  showStatus('');
  if (!validateFile(file)) {
    state.file = null;
    fileInput.value = '';
    fileName.classList.remove('has-file');
    fileName.textContent = '';
    updateGenerateState();
    return;
  }
  state.file = file;
  fileName.textContent = file.name;
  fileName.classList.add('has-file');
  updateGenerateState();
}

fileInput.addEventListener('change', () => setFile(fileInput.files?.[0]));

['dragenter', 'dragover'].forEach(type => {
  dropZone.addEventListener(type, event => {
    event.preventDefault();
    dropZone.classList.add('is-dragging');
  });
});

['dragleave', 'drop'].forEach(type => {
  dropZone.addEventListener(type, event => {
    event.preventDefault();
    dropZone.classList.remove('is-dragging');
  });
});

dropZone.addEventListener('drop', event => {
  const file = event.dataTransfer?.files?.[0];
  if (file) setFile(file);
});

function renderTemplates() {
  templateList.innerHTML = '';
  for (const template of state.templates) {
    const label = document.createElement('label');
    label.className = 'template-option';
    label.innerHTML = `
      <input type="radio" name="letterhead-template" value="${template.id}">
      <span>${template.label}</span>
    `;
    label.querySelector('input').addEventListener('change', () => {
      state.selectedTemplate = template;
      updateGenerateState();
      showStatus('');
    });
    templateList.appendChild(label);
  }
}

async function loadTemplates() {
  try {
    const response = await fetch('templates/letterheads/catalog.json', { cache: 'no-store' });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    state.templates = await response.json();
    renderTemplates();
  } catch (error) {
    templateList.innerHTML = '<div class="loading-row">Impossibile caricare l’elenco delle carte intestate.</div>';
    showStatus('Controlla che templates/letterheads/catalog.json sia presente nel repository.', 'error');
  }
}

function updateGenerateState() {
  generateButton.disabled = !(state.file && state.selectedTemplate);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function outputName(inputName, templateLabel) {
  const base = inputName.replace(/\.docx$/i, '');
  const cleanTemplate = templateLabel
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  return `${base}_carta_intestata_${cleanTemplate}.docx`;
}

generateButton.addEventListener('click', async () => {
  if (!state.file || !state.selectedTemplate) return;

  generateButton.disabled = true;
  generateButton.textContent = 'Elaborazione…';
  showStatus('Sto applicando intestazione e piè di pagina. Il file resta sul dispositivo.');

  try {
    const [sourceBuffer, templateResponse] = await Promise.all([
      state.file.arrayBuffer(),
      fetch(`templates/letterheads/${encodeURIComponent(state.selectedTemplate.file)}`)
    ]);

    if (!templateResponse.ok) {
      throw new Error(`Il modello “${state.selectedTemplate.label}” non è disponibile nel repository.`);
    }

    const templateBuffer = await templateResponse.arrayBuffer();
    const result = await applyLetterhead(sourceBuffer, templateBuffer);
    downloadBlob(result, outputName(state.file.name, state.selectedTemplate.label));
    showStatus('Documento generato. Il download dovrebbe essere iniziato automaticamente.', 'success');
  } catch (error) {
    console.error(error);
    showStatus(error?.message || 'Si è verificato un errore durante la generazione del documento.', 'error');
  } finally {
    generateButton.textContent = 'Genera documento';
    updateGenerateState();
  }
});

loadTemplates();
route(location.hash.includes('carta-intestata') ? 'letterhead' : 'home');
