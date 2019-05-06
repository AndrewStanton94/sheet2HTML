export let inputSanitisation = {
	//  Copies the value in a merged cell to every value in the merge
	//  and remove the merge
	//  Encode {r:1, c:1} => A1
	unpackMergedCells: (sheet) => {
		// This copies the value to all cells in the merge
		const populateMerge = ({ s, e }) => {
			const firstCellData = sheet[XLSX.utils.encode_cell(s)];
			for (let row = s.r; row <= e.r; row++) {
				for (let column = s.c; column <= e.c; column++) {
					const cellName = XLSX.utils.encode_cell({
						r: row,
						c: column,
					});
					if (!sheet[cellName]) {
						sheet[cellName] = firstCellData;
					}
				}
			}
		};

		// Run the populate merge for all merges
		sheet['!merges'].forEach((r) => populateMerge(r));
		//update the merge list
		sheet['!merges'] = [];

		return sheet;
	},

	// This takes JSON version of headers. Joins cells in collumn
	joinHeaderRows: (jsonHeaders) => {
		const COLUMNS = Math.max(...jsonHeaders.map((r) => r.length));
		let headers = [];

		const notSameAsPreceeding = (cell, i, array) =>
			i === 0 || cell !== array[i - 1];

		// For each column
		for (let c = 0; c < COLUMNS; c++) {
			let colTitles = [];
			// Get the header cells
			for (let r = 0; r < jsonHeaders.length; r++) {
				colTitles.push(jsonHeaders[r][c]);
			}

			// Combine them into one string
			headers.push(
				colTitles
					.filter(
						(cell, i, array) =>
							cell && notSameAsPreceeding(cell, i, array)
					)
					.join(' => ')
					.trim()
			);
		}
		return headers;
	},
};

export let calculators = {
	headerBodyRanges: (sheet, headerRowCount = 1) => {
		const { s: start, e: end } = XLSX.utils.decode_range(sheet['!ref']);

		let headerRange = XLSX.utils.encode_range({
			s: { r: start.r, c: start.c },
			e: { r: headerRowCount - 1, c: end.c },
		});

		let bodyRange = XLSX.utils.encode_range({
			s: { r: headerRowCount, c: start.c },
			e: { r: end.r, c: end.c },
		});
		return { headerRange, bodyRange };
	},
};
