import 'bulma/css/bulma.css';

import { XLSX_Wrapper } from './xlsxWrapper.js';
import { produceDataFromTemplate } from './template';
let spreadsheet;

const spreadsheetDropArea = document.getElementById('spreadsheetDropArea');
const sheetSelection = document.getElementById('sheetSelection');

const handleDragover = (e) => {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
};

const handleDrop = (e) => {
	e.stopPropagation();
	e.preventDefault();

	const files = e.dataTransfer.files,
		f = files[0];
	const reader = new FileReader();

	// Take file
	// Pass to XLSX wrapper for processing
	// Show sheet selector
	reader.onload = (e) => {
		const data = new Uint8Array(e.target.result);
		spreadsheet = new XLSX_Wrapper(data);
		document.spreadsheet = spreadsheet;

		let optionElems = spreadsheet.sheets.map(
			(sheet) => `<option value="${sheet}">${sheet}</option>`
		);
		sheetSelection.innerHTML = optionElems.join('\n\n');

		document.getElementById('previewSection').classList.remove('hidden');
	};

	reader.readAsArrayBuffer(f);
};

spreadsheetDropArea.addEventListener('drop', handleDrop, false);
spreadsheetDropArea.addEventListener('dragover', handleDragover, false);

// Show preview of document
document.getElementById('sheetPreviewForm').addEventListener('submit', (e) => {
	e.preventDefault();
	spreadsheet.selectedSheet = sheetSelection.value;
	const previewCode = spreadsheet.sheet_to_html;
	const previewElem = document.getElementById('preview');

	previewCode.classList.add('table', 'is-striped');
	previewElem.innerHTML = previewCode.outerHTML;

	previewElem.classList.remove('hidden');
	document.getElementById('headerSection').classList.remove('hidden');
});

// Processes the headers
document.getElementById('headerRowCountForm').addEventListener('submit', (e) => {
	e.preventDefault();
	spreadsheet.headerRowCount = parseInt(document.getElementById('headerRowCountInput').value);
	spreadsheet.clarifyHeaders();
	document.getElementById('templateSection').classList.remove('hidden');
});

// Uses mustache template to produce data
document.getElementById('templateForm').addEventListener('submit', (e) => {
	e.preventDefault();
	let templatedData = produceDataFromTemplate(
		document.getElementById('templateText').value,
		spreadsheet.sheet_to_json
	);
	XLSX_Wrapper.exportData(templatedData);
});
