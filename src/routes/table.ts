import { utils, writeFile } from 'xlsx';

export const export_table = () => {
	const table = document.getElementById('grade_table');
	const wb = utils.table_to_book(table);

	/* Export to file (start a download) */
	writeFile(wb, 'grade_processed.xlsx');
};

export const check_header_loc = (sheet_data: Array<Array<string | number>>): number => {
	for (let row = 0; row < sheet_data.length; row++) {
		if (sheet_data[row].length > 3) {
			return row;
		}
	}
	return -1;
};

export const table_styling = (sheet_data: Array<Array<string | number>>) => {
	const table = document.getElementById('grade_table') as HTMLTableElement;
	const table_row = table.getElementsByTagName('tr');

	for (let row = 0; row < sheet_data.length - 1; row++) {
		if (row % 2 == 0) {
			table_row[row].style.backgroundColor = '#E3E3E3';
		}
	}
};

export const sheet_to_html = (sheet_data: Array<Array<string | number>>, final: boolean) => {
	let table_output = '';

	let row_anchor = 0;

	row_anchor = check_header_loc(sheet_data);
	if (row_anchor < 0) {
		alert('No header found in the excel sheet!');
	}

	for (let row = row_anchor; row < sheet_data.length; row++) {
		table_output += '<tr>';

		for (let cell = 0; cell < sheet_data[row].length; cell++) {
			if (row == row_anchor) {
				if (final) {
					table_output += '<th>' + sheet_data[row][cell] + '</th>';
				} else {
					table_output +=
						'<th><input type="checkbox" class="headers" name=' +
						cell.toString +
						' value-' +
						cell.toString +
						'/><br>' +
						sheet_data[row][cell] +
						'</th>';
				}
			} else {
				const n = Number(sheet_data[row][cell]);

				if (Number(n) === n && n % 1 !== 0) {
					sheet_data[row][cell] = Math.round(n * 100) / 100;
				}
				table_output += '<td>' + sheet_data[row][cell] + '</td>';
			}
		}
		table_output += '</tr>';
	}

	const table_el = document.getElementById('grade_table') as HTMLElement;
	table_el.innerHTML = table_output;
};
