<script lang="ts">
	import { onMount } from 'svelte';
	import { read, utils, writeFile } from 'xlsx';

	let sheet_data: Array<Array<string | number>>;
	let table_original_width: number = 0;
	let first_call: boolean = true;

	let min_grade: number = 1;
	let max_grade: number = 6;
	let pass_grade: number = 4;
	let cut_4: number = 35;
	let cut_6: number = 75;
	let dist_data: Array<number> = [];

	const export_table = () => {
		const table = document.getElementById('grade_table');
		console.log(table);
		const wb = utils.table_to_book(table);

		/* Export to file (start a download) */
		writeFile(wb, 'grade_processed.xlsx');
	};

	const reset_data = () => {
		let plotDiv = document.getElementById('graph_holder')!;
		plotDiv.style.visibility = 'hidden';

		if (!first_call) {
			for (let row = 0; row < sheet_data.length; row++) {
				if (sheet_data[row].length > table_original_width) {
					sheet_data[row].pop();
				}
			}
		}

		sheet_to_table(sheet_data, false);
		table_styling();
	};

	const handleSubmit = () => {
		var check_box = document.getElementsByClassName('headers')!;

		let col_idx: number = -1;

		for (let i = 0; i < check_box.length; i++) {
			let checker = check_box[i] as HTMLInputElement;
			if (checker.checked) {
				col_idx = i;
				table_original_width = col_idx;
			}
		}

		if (col_idx < 0) {
			alert('No data column selected!');
			console.log(table_original_width);

			if (!first_call) {
				for (let row = 0; row < sheet_data.length; row++) {
					if (sheet_data[row].length > table_original_width) {
						sheet_data[row].pop();
					}
				}
			}

			sheet_to_table(sheet_data, false);
			table_styling();
			return;
		} else {
			if (first_call) {
				first_call = false;
			}
		}

		sheet_data = add_final_grade(
			sheet_data,
			col_idx,
			min_grade,
			max_grade,
			pass_grade,
			cut_4,
			cut_6
		);
		let plotDiv = document.getElementById('graph_holder')!;
		let row_anchor = check_header_loc(sheet_data);
		for (let row = row_anchor + 1; row < sheet_data.length; row++) {
			dist_data.push(sheet_data[row].at(-1) as number);
		}
		let data: Array<object> = [
			{
				x: dist_data,
				type: 'histogram',
				autobinx: false,
				xbins: {
					end: 6.125,
					size: 0.25,
					start: 0.875
				},
				marker: {
					color: '#2f90ed'
				}
			}
		];
		let mean: number =
			Math.round((100 * dist_data.reduce((a, b) => a + b, 0)) / dist_data.length) / 100;
		let std: number =
			Math.round(
				100 *
					Math.sqrt(
						dist_data.map((x) => Math.pow(x - mean, 2)).reduce((a, b) => a + b) / dist_data.length
					)
			) / 100;

		let num_fail: number = 0;
		let num_max: number = 0;

		for (let i = 0; i < dist_data.length; i++) {
			if (dist_data[i] < min_grade) {
				num_fail += 1;
			}
			if (dist_data[i] >= max_grade) {
				num_max += 1;
			}
		}

		let layout: object = {
			bargap: 0.05,
			xaxis: { title: 'Grade' },
			yaxis: { title: '#Counts' },
			title:
				'HS2023 TF Exam. Mean: ' +
				mean.toString() +
				' Std: ' +
				std.toString() +
				' Num_fail: ' +
				num_fail.toString() +
				' Num_max: ' +
				num_max.toString()
		};

		plotDiv.style.visibility = 'visible';
		Plotly.newPlot(plotDiv, data, layout);

		let table_data: string = sheet_to_table(sheet_data, true);
		document.getElementById('table_preview')!.innerHTML = '';
		document.getElementById('table_preview')!.innerHTML = table_data;
		table_styling();
	};

	const add_final_grade = (
		sheet_data: Array<Array<string | number>>,
		col_idx: number,
		min: number,
		max: number,
		pass: number,
		cut_4: number,
		cut_6: number
	): Array<Array<string | number>> => {
		let row_anchor = check_header_loc(sheet_data);
		for (var row = row_anchor; row < sheet_data.length; row++) {
			if (row == row_anchor) {
				sheet_data[row].push('Grade');
			} else {
				let val: number = sheet_data[row][col_idx] as number;
				let fail_gap: number = pass - min;
				let pass_gap: number = max - pass;

				let computed: number = 0;

				if (val < cut_4) {
					computed = 0.875 + (fail_gap * val) / cut_4;
				} else {
					computed = 3.875 + (pass_gap * (val - cut_4)) / (cut_6 - cut_4);
					if (computed > max) {
						computed = max;
					}
				}
				sheet_data[row].push(Math.round(4 * computed) / 4);
			}
		}

		return sheet_data;
	};

	const check_header_loc = (sheet_data: Array<Array<string | number>>): number => {
		for (let row = 0; row < sheet_data.length; row++) {
			if (sheet_data[row].length > 3) {
				return row;
			}
		}
		return -1;
	};

	const table_styling = () => {
		let table_header = document.getElementById('table_preview')!.getElementsByTagName('th');

		for (let i = 0; i < table_header.length; i++) {
			table_header[i].style.fontSize = '0.9em';
			table_header[i].style.padding = '5px';
			table_header[i].style.color = 'white';
			table_header[i].style.backgroundColor = '#2F90ED';
		}

		let table_data = document.getElementById('table_preview')!.getElementsByTagName('td');

		for (let i = 0; i < table_data.length; i++) {
			table_data[i].style.textAlign = 'center';
			table_data[i].style.fontSize = '0.8em';
		}

		for (let row = 1; row < sheet_data.length; row++) {
			if (row % 2 == 0) {
				document.getElementById('table_preview')!.getElementsByTagName('tr')[
					row
				].style.backgroundColor = '#E3E3E3';
			}
		}
	};

	const sheet_to_table = (sheet_data: Array<Array<string | number>>, final: boolean): string => {
		let table_output: string =
			'<table id="grade_table" class="table table-striped table-bordered">';

		let row_anchor = 0;

		row_anchor = check_header_loc(sheet_data);
		if (row_anchor < 0) {
			alert('No header found in the excel sheet!');
		}

		for (var row = row_anchor; row < sheet_data.length; row++) {
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
					let n = Number(sheet_data[row][cell]);

					if (Number(n) === n && n % 1 !== 0) {
						sheet_data[row][cell] = Math.round(n * 100) / 100;
					}
					table_output += '<td>' + sheet_data[row][cell] + '</td>';
				}
			}
			table_output += '</tr>';
		}

		table_output += '</table>';

		document.getElementById('table_preview')!.innerHTML = table_output;
		return table_output;
	};

	onMount(() => {
		const excel_file = document.getElementById('file_loader') as HTMLElement;

		excel_file.addEventListener('change', (event: Event) => {
			document.getElementById('form_inset')!.style.visibility = 'visible';

			let text_holder = document.getElementById('text_holder')!;
			text_holder.innerHTML = '<p>Loaded excel sheet:</p>';
			text_holder.style.marginTop = '20px';

			const target = event.target as HTMLInputElement;
			const file: File = (target.files as FileList)[0];

			let reader = new FileReader();
			console.log(file);
			console.log(reader);
			reader.readAsArrayBuffer(file);
			reader.addEventListener('load', (event: Event) => {
				let data = new Uint8Array(reader.result as ArrayBuffer);
				let work_book = read(data, { type: 'array' });
				let sheet_name = work_book.SheetNames;
				sheet_data = utils.sheet_to_json(work_book.Sheets[sheet_name[0]], {
					header: 1
				});
				if (sheet_data.length > 3) {
					first_call = true;
					sheet_to_table(sheet_data, false);
					table_styling();
				} else {
					document.getElementById('table_preview')!.innerHTML =
						'<div class="alert alert-danger" style="color: red; font-weight: 800";>Not enough data stored in the excel sheet! Check the data first.</div>';
					return false;
				}
			});
		});
	});
</script>

<svelte:head>
	<script src="https://cdn.plot.ly/plotly-latest.min.js" type="text/javascript"></script>
</svelte:head>

<div class="body">
	<div class="welcome">
		<h1>Welcome to IFD Grading System</h1>
	</div>

	<p>Please upload your grading excel sheet. (It only accepts .xls and .xlsx files)</p>

	<!-- (B) FILE PICKER -->
	<input type="file" id="file_loader" accept=".xls,.xlsx" />
	<div id="form_holder" />
	<form on:submit|preventDefault={handleSubmit}>
		<div id="form_inset">
			<p>Step 1: Select a column contains total points from the below excel sheet.</p>
			<p>Step 2: Set grading criteria from the below and press 'Process' button.</p>
			<p>Step 3: You can export processed data to xlsx file by pressing 'Export' button</p>
			<p>Step 4: If you want, you can reset the process. Press 'Reset'</p>
			<br />

			<label for="min">Grade min</label>
			<input type="text" name="min" id="min" size="4" bind:value={min_grade} />
			<label for="pass">Grade pass</label>
			<input type="text" name="pass" id="pass" size="4" bind:value={pass_grade} />
			<label for="max">Grade max:</label>
			<input type="text" name="max" id="max" size="4" bind:value={max_grade} />
			<br />
			<label for="cut_4">Lower cut:</label>
			<input type="text" name="cut_4" id="cut_4" size="4" bind:value={cut_4} />
			<label for="cut_6">Upper cut:</label>
			<input type="text" name="cut_6" id="cut_6" size="4" bind:value={cut_6} />
			<input type="submit" id="submit" value="Process" />
			<button type="button" id="export" on:click={export_table}>Export</button>
			<button type="button" id="reset" on:click={reset_data}>Reset</button>
		</div>
	</form>
	<div id="text_holder" />
	<div id="table_preview" />
	<div id="graph_holder" />
</div>

<style>
	@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;700&display=swap');

	* {
		font-family: 'Poppins', sans-serif;
		--default_margin: 20px;
		--default_half_margin: 10px;
	}

	.body {
		margin: 4em;
	}

	.welcome {
		display: flex;
		justify-content: center;
	}

	#table_preview {
		display: flex;
		justify-content: center;
		margin-top: var(--default_margin);
	}

	form {
		display: flex;
		justify-content: center;
		margin-top: var(--default_half_margin);
	}

	form input {
		margin: var(--default_half_margin);
	}

	form #form_inset {
		visibility: hidden;
	}

	#graph_holder {
		margin-top: var(--default_margin);
		display: flex;
		align-items: center;
		justify-content: center;
		visibility: hidden;
		height: auto;
	}
	#export {
		font-size: 0.7em;
		margin-left: var(--default_margin);
		font-weight: 800;
		background-color: #2f90ed;
		color: #e3e3e3;
		border-radius: 50%;
		padding: 7px;
	}
	#reset {
		font-size: 0.7em;
		margin-left: var(--default_margin);
		font-weight: 800;
		background-color: #cc1833;
		color: #e3e3e3;
		border-radius: 50%;
		padding: 7px;
	}
</style>
