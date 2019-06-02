import { inputSanitisation, calculators } from './utils.js';
import XLSX from 'xlsx';

/**
 * A class that represents a spreadsheet
 * Makes it easy to access some XLSX library functions
 * Adds some new features
 *
 * @export
 * @class XLSX_Wrapper
 */
export class XLSX_Wrapper {
	/**
	 * Creates an instance of XLSX_Wrapper.
	 *
	 * @constructs XLSX_Wrapper
	 * @param {Uint8Array} file
	 * @memberof XLSX_Wrapper
	 */
	constructor(file) {
		this.workbook = XLSX.read(file, {
			type: 'array',
		});
	}

	/**
	 *
	 * @readonly
	 * @memberof XLSX_Wrapper
	 * @returns {string[]} The sheets in the workbook
	 */
	get sheetNames() {
		return this.workbook.SheetNames;
	}

	/**
	 *
	 * @memberof XLSX_Wrapper
	 * @returns {string} The name of the current sheet
	 */
	get selectedSheet() {
		return this._selectedSheet;
	}

	/**
	 *
	 * @param {string} chosenSheet
	 * @memberof XLSX_Wrapper
	 */
	set selectedSheet(chosenSheet) {
		this._selectedSheet = chosenSheet;
	}

	/**
	 *
	 * @memberof XLSX_Wrapper
	 * @returns {XLSX.utils.Sheet} The data from the selected sheet
	 */
	get selectedSheetData() {
		return this.workbook.Sheets[this.selectedSheet];
	}

	/**
	 *
	 * @memberof XLSX_Wrapper
	 * @returns {number} Number of header rows
	 */
	get headerRowCount() {
		return this._headerRowCount;
	}

	/**
	 *
	 * @param {number} newCount
	 * @memberof XLSX_Wrapper
	 */
	set headerRowCount(newCount) {
		this._headerRowCount = newCount;
	}

	/**
	 *
	 * @memberof XLSX_Wrapper
	 * @returns {string[]} A list of headers
	 */
	get headers() {
		return this._headers;
	}

	/**
	 *
	 * @param {string[]} newHeaders
	 * @memberof XLSX_Wrapper
	 */
	set headers(newHeaders) {
		this._headers = newHeaders;
	}

	/**
	 * Returns HTML version of the selected sheet with the rows numbered
	 *
	 * @readonly
	 * @memberof XLSX_Wrapper
	 * @returns {HTMLTableElement}
	 */
	get sheet_to_html() {
		const dp = new DOMParser();
		const table = dp
			.parseFromString(XLSX.utils.sheet_to_html(this.selectedSheetData), 'text/html')
			.querySelector('table');

		[...table.children[0].children].forEach((row, i) => {
			const indexCell = document.createElement('td');
			indexCell.innerText = `${i + 1}`;
			row.insertAdjacentElement('afterbegin', indexCell);
		});
		return table;
	}

	/**
	 *
	 *
	 * @readonly
	 * @memberof XLSX_Wrapper
	 * @returns {Object[]}
	 */
	get sheet_to_json() {
		return XLSX.utils.sheet_to_json(this.selectedSheetData);
	}

	/**
	 * Calls utility functions to remove merges and join header rows
	 *
	 * @todo Should be more granular?
	 * @todo Needs more options
	 * @memberof XLSX_Wrapper
	 */
	clarifyHeaders() {
		if (!this.selectedSheet) {
			throw new Error('Need to select a sheet');
		}
		if (!this.headerRowCount) {
			throw new Error('Need to know how many columns are in the header');
		}

		let sheetData = inputSanitisation.unpackMergedCells(this.selectedSheetData);

		const { headerRange, bodyRange } = calculators.headerBodyRanges(
			sheetData,
			this.headerRowCount
		);

		const jsonHeaders = XLSX.utils.sheet_to_json(sheetData, {
			range: headerRange,
			header: 1,
		});

		this.headers = inputSanitisation.joinHeaderRows(jsonHeaders);
		this.headerRowCount = 1;

		const json = XLSX.utils.sheet_to_json(sheetData, {
			range: bodyRange,
			header: this.headers,
		});

		this.workbook.Sheets[this.selectedSheet] = XLSX.utils.json_to_sheet(json);

		return json;
	}

	/**
	 * Takes JSON data and exports it in a new spreadsheet
	 *
	 * @static
	 * @param {object[]} jsonData
	 * @memberof XLSX_Wrapper
	 */
	static exportData(jsonData) {
		let exportWorkbook = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(
			exportWorkbook,
			XLSX.utils.json_to_sheet(jsonData),
			'Data Export'
		);
		XLSX.writeFile(exportWorkbook, 'out.xlsx');
	}
}
