import { type } from 'os';

// import { Sheet } from 'src/xlsxWrapper';

export namespace CFB {
	function find(cfb: any, path: any): any;
	function parse(file: any, options: any): any;
	function read(blob: any, options: any): any;
	namespace utils {
		function CheckField(hexstr: any, fld: any): void;
		class ReadShift {
			constructor(size: any, t: any);
			l: any;
		}
		function bconcat(bufs: any): any;
		function cfb_add(cfb: any, name: any, content: any, opts: any): any;
		function cfb_del(cfb: any, name: any): any;
		function cfb_gc(cfb: any): void;
		function cfb_mov(cfb: any, old_name: any, new_name: any): any;
		function cfb_new(opts: any): any;
		const consts: {
			DIFSECT: number;
			ENDOFCHAIN: number;
			EntryTypes: any[];
			FATSECT: number;
			FREESECT: number;
			HEADER_CLSID: string;
			HEADER_MINOR_VERSION: string;
			HEADER_SIGNATURE: string;
			MAXREGSECT: number;
			MAXREGSID: number;
			NOSTREAM: number;
		};
		function prep_blob(blob: any, pos: any): void;
		function use_zlib(zlib: any): void;
	}
	const version: string;
	function write(cfb: any, options: any): any;
	function writeFile(cfb: any, filename: any, options: any): void;
}
export namespace SSF {
	function format(fmt: any, v: any, o: any): any;
	function get_table(): any;
	function init_table(t: any): void;
	function is_date(fmt: any): any;
	function load(fmt: any, idx: any): any;
	function load_table(tbl: any): void;
	function parse_date_code(v: any, opts: any, b2: any): any;
	const version: string;
}
export function parse_fods(data: any, opts: any): any;
export function parse_ods(zip: any, opts: any): any;
export function parse_xlscfb(cfb: any, options: any): any;
export function parse_zip(zip: any, opts: any): any;
export function read(data: any, opts: any): any;
export function readFile(filename: any, opts: any): any;
export function readFileSync(filename: any, opts: any): any;
export namespace stream {
	function to_csv(sheet: Sheet, opts: any): any;
	function to_html(ws: any, opts: any): any;
	function to_json(sheet: Sheet, opts: any): Object[];
}
export namespace utils {
	interface Cell {
		r: number;
		c: number;
	}
	interface CellRange {
		s: Cell;
		e: Cell;
	}
	interface Book {
		SheetNames: string[];
		Sheets: Object;
	}
	interface Sheet {}
	type EncodedCell = string;
	type EncodedCellRange = string;
	function aoa_to_sheet(data: any, opts: any): Sheet;
	function book_append_sheet(wb: Book, ws: Sheet, name: string): void;
	function book_new(): Book;
	function book_set_sheet_visibility(wb: Book, sh: Sheet, vis: any): void;
	function cell_add_comment(cell: Cell, text: any, author: any): void;
	function cell_set_hyperlink(cell: Cell, target: any, tooltip: any): any;
	function cell_set_internal_link(cell: Cell, range: any, tooltip: any): any;
	function cell_set_number_format(cell: Cell, fmt: any): any;
	const consts: {
		SHEET_HIDDEN: number;
		SHEET_VERY_HIDDEN: number;
		SHEET_VISIBLE: number;
	};
	function decode_cell(cstr: any): any;
	function decode_col(colstr: any): any;
	function decode_range(range: any): CellRange;
	function decode_row(rowstr: any): any;
	function encode_cell(cell: Cell): EncodedCell;
	function encode_col(col: any): any;
	function encode_range(cs: Cell, ce: Cell): EncodedCellRange;
	function encode_row(row: any): any;
	function format_cell(cell: Cell, v: any, o: any): any;
	function get_formulae(sheet: Sheet): any;
	function json_to_sheet(js: Object[], opts?: any): Sheet;
	function make_csv(sheet: any, opts: any): any;
	function make_formulae(sheet: Sheet): any;
	function make_json(sheet: Sheet, opts: any): any;
	function sheet_add_aoa(_ws: Sheet, data: any, opts: any): any;
	function sheet_add_json(_ws: Sheet, js: Object[], opts: any): any;
	function sheet_set_array_formula(ws: Sheet, range: any, formula: any): any;
	function sheet_to_csv(sheet: Sheet, opts: any): any;
	function sheet_to_dif(ws: Sheet): any;
	function sheet_to_eth(ws: Sheet): any;
	function sheet_to_formulae(sheet: Sheet): any;
	function sheet_to_html(ws: Sheet, opts?: any): any;
	function sheet_to_json(sheet: Sheet, opts?: any): Object[];
	function sheet_to_row_object_array(sheet: Sheet, opts: any): any;
	function sheet_to_slk(ws: Sheet, opts: any): any;
	function sheet_to_txt(sheet: Sheet, opts: any): any;
	function split_cell(cstr: any): any;
	function table_to_book(table: any, opts: any): any;
	function table_to_sheet(table: any, _opts: any): Sheet;
}
interface Book {
	SheetNames: string[];
	Sheets: Object;
}
export const version: string;
export function write(wb: Book, opts: any): any;
export function writeFile(wb: Book, filename: string, opts?: any): any;
export function writeFileAsync(filename: string, wb: Book, opts: any, cb: any): any;
export function writeFileSync(wb: Book, filename: string, opts: any): any;
export function write_ods(wb: Book, opts: any): any;
