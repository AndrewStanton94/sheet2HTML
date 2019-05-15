import { inputSanitisation, calculators } from './utils.js';
import XLSX from 'xlsx';

export class XLSX_Wrapper {
	constructor(file) {
		this.workbook = XLSX.read(file, {
			type: 'array',
		});
	}

	get sheets() {
		return this.workbook.SheetNames;
	}

	get selectedSheet() {
		return this._selectedSheet;
	}

	set selectedSheet(chosenSheet) {
		this._selectedSheet = chosenSheet;
	}

	get selectedSheetData() {
		return this.workbook.Sheets[this.selectedSheet];
	}

	get headerRowCount() {
		return this._headerRowCount;
	}

	set headerRowCount(newCount) {
		return (this._headerRowCount = newCount);
	}

	get headers() {
		return this._headers;
	}

	set headers(newHeaders) {
		this._headers = newHeaders;
	}

	// Returns HTML version of the selected sheet with the rows numbered
	get sheet_to_html() {
		const dp = new DOMParser();
		const table = dp
			.parseFromString(
				XLSX.utils.sheet_to_html(this.selectedSheetData),
				'text/html'
			)
			.querySelector('table');

		[...table.children[0].children].forEach((row, i) => {
			const indexCell = document.createElement('td');
			indexCell.innerText = `${i + 1}`;
			row.insertAdjacentElement('afterbegin', indexCell);
		});
		document.table = table;
		return table;
	}

	get sheet_to_json() {
		return XLSX.utils.sheet_to_json(this.selectedSheetData);
	}

	clarifyHeaders() {
		if (!this.selectedSheet) {
			throw new Error('Need to select a sheet');
		}
		if (!this.headerRowCount) {
			throw new Error('Need to know how many columns are in the header');
		}

		let sheetData = inputSanitisation.unpackMergedCells(
			this.selectedSheetData
		);

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

		this.workbook.Sheets[this.selectedSheet] = XLSX.utils.json_to_sheet(
			json
		);

		return json;
	}

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
