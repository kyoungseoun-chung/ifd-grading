import { check_header_loc } from './table';

interface Statistics {
	min: number;
	max: number;
	mean: number;
	std: number;
	pass: number;
	fail: number;
	full: number;
	total: number;
}

export const statistics = (
	data: Array<number>,
	pass_grade: number,
	max_grade: number
): Statistics => {
	const mean: number = Math.round((100 * data.reduce((a, b) => a + b, 0)) / data.length) / 100;
	const std: number =
		Math.round(
			100 * Math.sqrt(data.map((x) => Math.pow(x - mean, 2)).reduce((a, b) => a + b) / data.length)
		) / 100;

	const min: number = Math.min(...data);
	const max: number = Math.max(...data);

	let fail = 0;
	let full = 0;

	for (let i = 0; i < data.length; i++) {
		if (data[i] < pass_grade) {
			fail += 1;
		}
		if (data[i] >= max_grade) {
			full += 1;
		}
	}

	return {
		min: min,
		max: max,
		mean: mean,
		std: std,
		pass: data.length - fail,
		fail: fail,
		full: full,
		total: data.length
	};
};

export const add_final_grade = (
	sheet_data: Array<Array<string | number>>,
	col_idx: number,
	min: number,
	max: number,
	pass: number,
	cut_4: number,
	cut_6: number,
	recompute = false
): Array<Array<string | number>> => {
	const row_anchor = check_header_loc(sheet_data);

	for (let row = row_anchor; row < sheet_data.length; row++) {
		if (row == row_anchor) {
			if (!recompute) {
				sheet_data[row].push('Grade');
			}
		} else {
			const val: number = sheet_data[row][col_idx] as number;
			const fail_gap: number = pass - min;
			const pass_gap: number = max - pass;

			let computed = 0;

			if (val < cut_4) {
				computed = 0.875 + (fail_gap * val) / cut_4;
			} else {
				computed = 3.875 + (pass_gap * (val - cut_4)) / (cut_6 - cut_4);
				if (computed > max) {
					computed = max;
				}
			}
			if (recompute) {
				sheet_data[row][sheet_data[row].length - 1] = Math.round(4 * computed) / 4;
			} else {
				sheet_data[row].push(Math.round(4 * computed) / 4);
			}
		}
	}

	return sheet_data;
};
