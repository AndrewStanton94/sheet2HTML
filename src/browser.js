import 'bulma/css/bulma.css';
import './main.css';

import { XLSX_Wrapper } from './xlsxWrapper.js';
import { produceDataFromTemplate } from './template';
/**
 * @type XLSX_Wrapper
 */
let spreadsheet;

/**
 * @type HTMLSelectElement
 */
// @ts-ignore
const sheetSelection = document.getElementById('sheetSelection');
const spreadsheetDropArea = document.getElementById('spreadsheetDropArea');

/**
 * Neutralises the dragover event
 * Selects the copy dropEffect
 *
 * @param { DragEvent } e
 */
const handleDragover = (e) => {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
};

/**
 * Receive the spreadsheet from the drop event
 * Pass it to a file reader for processing
 *
 * @param { DragEvent} e
 */
const handleDrop = (e) => {
	e.stopPropagation();
	e.preventDefault();

	const files = e.dataTransfer.files,
		f = files[0],
		reader = new FileReader();

	/**
	 * Take the loaded file
	 * Pass it to a new XLSX_Wrapper instance
	 * Populate a list of sheet names and display it
	 * Generate and show sheet selector
	 *
	 * @param { ProgressEvent } e
	 */
	reader.onload = (e) => {
		const data = new Uint8Array(e.target.result);
		spreadsheet = new XLSX_Wrapper(data);
		document.spreadsheet = spreadsheet;

		/**
		 * Generate option tags from the sheet names
		 *
		 * @param {string} sheet
		 */
		let optionElems = spreadsheet.sheetNames.map(
			(sheet) => `<option value="${sheet}">${sheet}</option>`
		);
		sheetSelection.innerHTML = optionElems.join('\n\n');

		document.getElementById('previewSection').classList.remove('hidden');
	};

	reader.readAsArrayBuffer(f);
};

spreadsheetDropArea.addEventListener('drop', handleDrop, false);
spreadsheetDropArea.addEventListener('dragover', handleDragover, false);

/**
 * Show sheet preview
 * Gets the name of the selected sheet and sets the class property
 * The property is used to generate the table from the sheet
 * The table is styled by bulma and added to the page
 *
 * @param { Event} e
 */
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

/**
 * Extract the headers using the number of header rows as specified by the user
 *
 * @param {Event} e
 */
document.getElementById('headerRowCountForm').addEventListener('submit', (e) => {
	e.preventDefault();
	// @ts-ignore
	spreadsheet.headerRowCount = parseInt(document.getElementById('headerRowCountInput').value);
	spreadsheet.clarifyHeaders();
	document.getElementById('templateSection').classList.remove('hidden');
});

/**
 * Get the template from the page and the JSON version of the data
 * Generate the templated data
 * Download it
 *
 * @param {Event} e
 */
document.getElementById('templateForm').addEventListener('submit', (e) => {
	e.preventDefault();

	/**
	 * @type {HTMLTextAreaElement}
	 */
	// @ts-ignore
	let templateTextArea = document.getElementById('templateText');
	let templatedData = produceDataFromTemplate(templateTextArea.value, spreadsheet.sheet_to_json);
	XLSX_Wrapper.exportData(templatedData);
});
