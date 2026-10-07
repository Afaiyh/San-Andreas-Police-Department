let docxArrayBuffer = null;
let fieldKeys = [];

// 1. Get requested filename from URL query parameter (?file=blabla.docx)
const urlParams = new URLSearchParams(window.location.search);
const targetFile = urlParams.get('file') || 'blabla.docx';

document.getElementById('file-title').innerText = `Editing: ${targetFile}`;

// Relative path to document file inside /form/ directory
const templatePath = `../${targetFile}`;

// 2. Fetch template file from static server
fetch(templatePath)
  .then(response => {
    if (!response.ok) throw new Error(`HTTP ${response.status}: Could not load ${templatePath}`);
    return response.arrayBuffer();
  })
  .then(buffer => {
    docxArrayBuffer = buffer;
    inspectAndBuildForm(buffer);
  })
  .catch(err => {
    console.error(err);
    document.getElementById('template-form').innerHTML = `
      <div class="p-4 bg-rose-950/50 border border-rose-800 rounded-lg text-rose-300 text-sm">
        <strong>Error:</strong> Could not load template <code>form/${targetFile}</code>. Ensure the file exists in your repository directory.
      </div>
    `;
  });

// 3. Scan DOCX for template placeholders {VARIABLE} and build inputs
function inspectAndBuildForm(buffer) {
  try {
    const zip = new PizZip(buffer);
    const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    
    // Scan text content for {VAR_NAME} placeholders
    const text = doc.getFullText();
    const matches = text.match(/\{([^}]+)\}/g) || [];
    
    // Deduplicate placeholder tags
    fieldKeys = [...new Set(matches.map(m => m.replace(/[{}]/g, '').trim()))];

    const formEl = document.getElementById('template-form');
    
    // Fallback default keys if document contains no bracket tags
    if (fieldKeys.length === 0) {
      fieldKeys = ['OFFICER_NAME', 'DATE', 'CASE_NUMBER', 'LOCATION_DETAILS'];
    }

    formEl.innerHTML = fieldKeys.map(key => `
      <div>
        <label class="block text-xs font-semibold uppercase tracking-wider text-slate-400 mb-1.5">${key.replace(/_/g, ' ')}</label>
        <input type="text" id="field-${key}" name="${key}" placeholder="Enter ${key.toLowerCase()}..." />
      </div>
    `).join('');

  } catch (e) {
    console.error("Template processing failed:", e);
  }
}

// 4. Populate template and export on button click
document.getElementById('export-btn').addEventListener('click', () => {
  if (!docxArrayBuffer) return;

  try {
    const zip = new PizZip(docxArrayBuffer);
    const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

    // Collect values from generated inputs
    const formData = {};
    fieldKeys.forEach(key => {
      const input = document.getElementById(`field-${key}`);
      formData[key] = input ? input.value : '';
    });

    // Render variables into document body
    doc.render(formData);

    // Generate output blob
    const blob = doc.getZip().generate({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    // Trigger client-side file download
    saveAs(blob, `filled_${targetFile}`);

  } catch (error) {
    alert("Error generating document: " + error.message);
  }
});
