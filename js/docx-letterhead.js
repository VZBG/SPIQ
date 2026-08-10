const XML = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types'
};

const HEADER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const FOOTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';

function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  const error = doc.querySelector('parsererror');
  if (error) throw new Error('Impossibile leggere la struttura XML del documento Word.');
  return doc;
}

function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function basename(path) {
  return path.split('/').pop();
}

function dirname(path) {
  const i = path.lastIndexOf('/');
  return i === -1 ? '' : path.slice(0, i + 1);
}

function normalizePath(path) {
  const parts = [];
  for (const p of path.split('/')) {
    if (!p || p === '.') continue;
    if (p === '..') parts.pop();
    else parts.push(p);
  }
  return parts.join('/');
}

function resolveRelative(baseFile, target) {
  return normalizePath(dirname(baseFile) + target);
}

function relationshipPartFor(partPath) {
  const dir = dirname(partPath);
  const name = basename(partPath);
  return `${dir}_rels/${name}.rels`;
}

function nextRelationshipId(relsDoc) {
  const used = new Set([...relsDoc.documentElement.children].map(n => n.getAttribute('Id')));
  let n = 1;
  while (used.has(`rId${n}`)) n += 1;
  return `rId${n}`;
}

function uniquePartName(zip, prefix, ext) {
  let n = 1;
  let path;
  do {
    path = `word/${prefix}${n}.${ext}`;
    n += 1;
  } while (zip.file(path));
  return path;
}

function uniqueMediaPath(zip, originalPath) {
  const name = basename(originalPath);
  const dot = name.lastIndexOf('.');
  const stem = dot >= 0 ? name.slice(0, dot) : name;
  const ext = dot >= 0 ? name.slice(dot) : '';
  let n = 1;
  let path = `word/media/${stem}${ext}`;
  while (zip.file(path)) {
    path = `word/media/${stem}_spiq_${n}${ext}`;
    n += 1;
  }
  return path;
}

function ensureContentTypeOverride(contentTypesDoc, partName, contentType) {
  const root = contentTypesDoc.documentElement;
  const normalized = partName.startsWith('/') ? partName : `/${partName}`;
  const exists = [...root.children].some(node => node.localName === 'Override' && node.getAttribute('PartName') === normalized);
  if (exists) return;
  const el = contentTypesDoc.createElementNS(XML.ct, 'Override');
  el.setAttribute('PartName', normalized);
  el.setAttribute('ContentType', contentType);
  root.appendChild(el);
}

function findContentType(contentTypesDoc, partName) {
  const normalized = partName.startsWith('/') ? partName : `/${partName}`;
  const override = [...contentTypesDoc.documentElement.children].find(node => node.localName === 'Override' && node.getAttribute('PartName') === normalized);
  return override?.getAttribute('ContentType') || null;
}

function removeHeaderFooterRefs(sectPr) {
  [...sectPr.children].forEach(node => {
    if (node.namespaceURI === XML.w && (node.localName === 'headerReference' || node.localName === 'footerReference')) {
      node.remove();
    }
  });
}

function addSectionReference(doc, sectPr, kind, type, rId) {
  const el = doc.createElementNS(XML.w, `w:${kind}Reference`);
  el.setAttributeNS(XML.w, 'w:type', type);
  el.setAttributeNS(XML.r, 'r:id', rId);
  const firstNonRef = [...sectPr.children].find(node => !(node.namespaceURI === XML.w && (node.localName === 'headerReference' || node.localName === 'footerReference')));
  sectPr.insertBefore(el, firstNonRef || null);
}

async function copyPartAndDependencies({ sourceZip, templateZip, templatePartPath, sourcePartPath }) {
  const partFile = templateZip.file(templatePartPath);
  if (!partFile) throw new Error(`Parte del modello non trovata: ${templatePartPath}`);
  sourceZip.file(sourcePartPath, await partFile.async('uint8array'));

  const templateRelsPath = relationshipPartFor(templatePartPath);
  const relsFile = templateZip.file(templateRelsPath);
  if (!relsFile) return;

  const relsDoc = parseXml(await relsFile.async('text'));
  const relationships = [...relsDoc.documentElement.children];

  for (const rel of relationships) {
    const mode = rel.getAttribute('TargetMode');
    const target = rel.getAttribute('Target');
    if (!target || mode === 'External') continue;

    const dependencyPath = resolveRelative(templatePartPath, target);
    const dependency = templateZip.file(dependencyPath);
    if (!dependency) continue;

    let newDependencyPath = dependencyPath;
    if (dependencyPath.startsWith('word/media/')) {
      newDependencyPath = uniqueMediaPath(sourceZip, dependencyPath);
    } else if (sourceZip.file(dependencyPath)) {
      const name = basename(dependencyPath);
      const dot = name.lastIndexOf('.');
      const stem = dot >= 0 ? name.slice(0, dot) : name;
      const ext = dot >= 0 ? name.slice(dot + 1) : 'bin';
      newDependencyPath = uniquePartName(sourceZip, stem + '_spiq_', ext);
    }

    sourceZip.file(newDependencyPath, await dependency.async('uint8array'));
    const relTargetBase = dirname(sourcePartPath);
    const relative = newDependencyPath.startsWith(relTargetBase)
      ? newDependencyPath.slice(relTargetBase.length)
      : newDependencyPath.replace(/^word\//, '');
    rel.setAttribute('Target', relative);
  }

  sourceZip.file(relationshipPartFor(sourcePartPath), serializeXml(relsDoc));
}

function getTemplateSectionRefs(templateDocumentXml, templateDocumentRels) {
  const doc = parseXml(templateDocumentXml);
  const relsDoc = parseXml(templateDocumentRels);
  const sectPrs = [...doc.getElementsByTagNameNS(XML.w, 'sectPr')];
  if (!sectPrs.length) throw new Error('Il modello selezionato non contiene impostazioni di sezione.');
  const sectPr = sectPrs[sectPrs.length - 1];

  const relMap = new Map([...relsDoc.documentElement.children].map(rel => [rel.getAttribute('Id'), rel]));
  const refs = [];
  for (const node of [...sectPr.children]) {
    if (node.namespaceURI !== XML.w) continue;
    if (node.localName !== 'headerReference' && node.localName !== 'footerReference') continue;
    const rId = node.getAttributeNS(XML.r, 'id');
    const rel = relMap.get(rId);
    if (!rel) continue;
    refs.push({
      kind: node.localName === 'headerReference' ? 'header' : 'footer',
      type: node.getAttributeNS(XML.w, 'type') || 'default',
      target: resolveRelative('word/document.xml', rel.getAttribute('Target'))
    });
  }
  return refs;
}

export async function applyLetterhead(sourceArrayBuffer, templateArrayBuffer) {
  if (!window.JSZip) throw new Error('La libreria necessaria per elaborare i file Word non è disponibile.');

  const [sourceZip, templateZip] = await Promise.all([
    window.JSZip.loadAsync(sourceArrayBuffer),
    window.JSZip.loadAsync(templateArrayBuffer)
  ]);

  const required = ['word/document.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml'];
  for (const path of required) {
    if (!sourceZip.file(path)) throw new Error('Il file caricato non sembra essere un documento .docx valido.');
    if (!templateZip.file(path)) throw new Error('Il modello di carta intestata non è un .docx valido.');
  }

  const [templateDocumentXml, templateDocumentRels, sourceDocumentXml, sourceDocumentRels, sourceContentTypes, templateContentTypes] = await Promise.all([
    templateZip.file('word/document.xml').async('text'),
    templateZip.file('word/_rels/document.xml.rels').async('text'),
    sourceZip.file('word/document.xml').async('text'),
    sourceZip.file('word/_rels/document.xml.rels').async('text'),
    sourceZip.file('[Content_Types].xml').async('text'),
    templateZip.file('[Content_Types].xml').async('text')
  ]);

  const templateRefs = getTemplateSectionRefs(templateDocumentXml, templateDocumentRels);
  if (!templateRefs.length) throw new Error('Nel modello selezionato non sono stati trovati header o footer.');

  const sourceDoc = parseXml(sourceDocumentXml);
  const sourceRelsDoc = parseXml(sourceDocumentRels);
  const sourceContentTypesDoc = parseXml(sourceContentTypes);
  const templateContentTypesDoc = parseXml(templateContentTypes);

  const copiedRefs = [];
  for (const ref of templateRefs) {
    const ext = ref.target.split('.').pop() || 'xml';
    const newPart = uniquePartName(sourceZip, `spiq_${ref.kind}_`, ext);
    await copyPartAndDependencies({ sourceZip, templateZip, templatePartPath: ref.target, sourcePartPath: newPart });

    const contentType = findContentType(templateContentTypesDoc, ref.target) ||
      (ref.kind === 'header'
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
        : 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
    ensureContentTypeOverride(sourceContentTypesDoc, newPart, contentType);

    const rId = nextRelationshipId(sourceRelsDoc);
    const rel = sourceRelsDoc.createElementNS(XML.rel, 'Relationship');
    rel.setAttribute('Id', rId);
    rel.setAttribute('Type', ref.kind === 'header' ? HEADER_REL : FOOTER_REL);
    rel.setAttribute('Target', newPart.replace(/^word\//, ''));
    sourceRelsDoc.documentElement.appendChild(rel);
    copiedRefs.push({ ...ref, rId });
  }

  const sourceSections = [...sourceDoc.getElementsByTagNameNS(XML.w, 'sectPr')];
  if (!sourceSections.length) throw new Error('Il documento caricato non contiene impostazioni di sezione compatibili.');

  for (const sectPr of sourceSections) {
    removeHeaderFooterRefs(sectPr);
    for (const ref of copiedRefs) {
      addSectionReference(sourceDoc, sectPr, ref.kind, ref.type, ref.rId);
    }
  }

  sourceZip.file('word/document.xml', serializeXml(sourceDoc));
  sourceZip.file('word/_rels/document.xml.rels', serializeXml(sourceRelsDoc));
  sourceZip.file('[Content_Types].xml', serializeXml(sourceContentTypesDoc));

  return sourceZip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
}
