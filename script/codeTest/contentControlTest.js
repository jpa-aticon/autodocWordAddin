const AdmZip = require('adm-zip');
const { XMLParser } = require('fast-xml-parser');

function extractDocProperties(docxPath) {
  const zip = new AdmZip(docxPath);
  // Extract core properties
  const coreEntry = zip.getEntry('docProps/core.xml');
  let status = null, tags = null;
  if (coreEntry) {
    const coreXml = coreEntry.getData().toString('utf8');
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '',
    });
    const coreJson = parser.parse(coreXml);
    // Status is often stored in <cp:contentStatus>
    status = coreJson['cp:coreProperties'] && coreJson['cp:coreProperties']['cp:contentStatus']
      ? coreJson['cp:coreProperties']['cp:contentStatus']
      : null;
    // Tags/Keywords are often stored in <cp:keywords>
    tags = coreJson['cp:coreProperties'] && coreJson['cp:coreProperties']['cp:keywords']
      ? coreJson['cp:coreProperties']['cp:keywords']
      : null;
  }
  return { status, tags };
}

function extractContentControlTags(docxPath) {
  const zip = new AdmZip(docxPath);
  const xmlEntry = zip.getEntry('word/document.xml');
  if (!xmlEntry) {
    throw new Error('document.xml not found in the DOCX file!');
  }
  const xml = xmlEntry.getData().toString('utf8');

  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '',
  });
  const json = parser.parse(xml);

  console.log(JSON.stringify(json, null, 2));

  // Recursively find all w:sdt elements and extract tag names from w:sdtPr > w:tag
  function findSDTTags(node, tags = []) {
    if (Array.isArray(node)) {
      node.forEach(child => findSDTTags(child, tags));
    } else if (typeof node === 'object' && node !== null) {
      // If this is a w:sdt with w:sdtPr containing w:tag
      if (node['w:sdtPr'] && node['w:sdtPr']['w:tag'] && node['w:sdtPr']['w:tag'].val) {
        tags.push(node['w:sdtPr']['w:tag'].val);
      }
      // Recursively check all properties
      Object.values(node).forEach(child => findSDTTags(child, tags));
    }
    return tags;
  }

  return findSDTTags(json);
}

function extractContentControlTagAndAlias(docxPath) {
  const zip = new AdmZip(docxPath);
  const xmlEntry = zip.getEntry('word/document.xml');
  if (!xmlEntry) {
    throw new Error('document.xml not found in the DOCX file!');
  }
  const xml = xmlEntry.getData().toString('utf8');

  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '',
  });
  const json = parser.parse(xml);

  // Recursively find all w:sdt elements and extract tag and alias from w:sdtPr
  function findSDTInfo(node, results = []) {
    if (Array.isArray(node)) {
      node.forEach(child => findSDTInfo(child, results));
    } else if (typeof node === 'object' && node !== null) {
      if (node['w:sdtPr']) {
        const tagObj = node['w:sdtPr']['w:tag'];
        const aliasObj = node['w:sdtPr']['w:alias'];
        const tag = tagObj && tagObj['w:val'] ? tagObj['w:val'] : null;
        const alias = aliasObj && aliasObj['w:val'] ? aliasObj['w:val'] : null;
        if (tag || alias) {
          results.push({ tag, alias });
        }
      }
      Object.values(node).forEach(child => findSDTInfo(child, results));
    }
    return results;
  }

  return findSDTInfo(json);
}

// Usage
const path = 'AddinTest.docx'; // Change this to your .docx file name if needed

// Extract and print document properties
const props = extractDocProperties(path);
console.log('Document Properties:');
console.log('Status:', props.status || '(none)');
console.log('Tags:', props.tags || '(none)');

// Extract and print content control tags
const tags = extractContentControlTags(path).filter(Boolean);

if (tags.length === 0) {
  console.log('No content control tags found.');
} else {
  console.log('Content control tags:');
  tags.forEach((tag, idx) => {
    console.log(`${idx + 1}. ${tag}`);
  });
}

// Extract and print content control tag and alias
const controls = extractContentControlTagAndAlias(path);
console.log('Content Controls (tag and alias):');
controls.forEach((cc, idx) => {
  console.log(`${idx + 1}. Tag: ${cc.tag || '(none)'}, Alias: ${cc.alias || '(none)'}`);
});