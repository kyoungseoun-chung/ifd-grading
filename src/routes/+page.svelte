<script lang="ts">
	import { onMount } from 'svelte';
	import { read, utils } from 'xlsx';
	import { add_final_grade, statistics } from './grade';
	import { table_styling, check_header_loc, sheet_to_html, export_table } from './table';
	import { plot_d3 } from './plot';

	let sheet_data: Array<Array<string | number>>;
	let table_original_width: number = 0;
	let first_call: boolean = true;

	let min_grade: number = 1;
	let max_grade: number = 6;
	let pass_grade: number = 4;
	let cut_4: number = 35;
	let cut_6: number = 75;
	let dist_data: Array<number> = [];

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
			if (first_call) {
				return;
			}

			let row_anchor = check_header_loc(sheet_data);

			for (let row = row_anchor; row < sheet_data.length; row++) {
				if (sheet_data[row].length > table_original_width + 1) {
					sheet_data[row].pop();
				}
			}

			sheet_to_html(sheet_data, false);
			table_styling(sheet_data);
			return;
		} else {
			first_call = false;
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

		let idx = 0;
		for (let row = row_anchor + 1; row < sheet_data.length; row++) {
			if (first_call) {
				dist_data.push(sheet_data[row].at(-1) as number);
			} else {
				dist_data[idx] = sheet_data[row].at(-1) as number;
				idx += 1;
			}
		}

		let stat = statistics(dist_data, pass_grade, max_grade);

		let stats: string =
			'Mean: ' +
			stat.mean.toString() +
			', Std: ' +
			stat.std.toString() +
			', Num_fail: ' +
			stat.fail.toString() +
			', Num_max: ' +
			stat.full.toString() +
			', Num_pass: ' +
			stat.pass.toString() +
			', Among: ' +
			stat.total.toString() +
			' students';

		const g_holder = document.getElementById('graph_holder') as HTMLElement;
		g_holder.style.visibility = 'visible';
		g_holder.innerHTML = '<p>TF Grading Results</p><p>' + stats + '</p><br>';

		plot_d3(dist_data, min_grade, max_grade, 0.25);

		let table_data: string = sheet_to_html(sheet_data, true);
		document.getElementById('table_preview')!.innerHTML = '';
		document.getElementById('table_preview')!.innerHTML = table_data;
		table_styling(sheet_data);
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
					sheet_to_html(sheet_data, false);
					table_styling(sheet_data);
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
		</div>
	</form>

	<div id="graph_holder" />
	<div id="text_holder" />
	<div id="table_preview" />
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
		flex-direction: column;
	}

	#export {
		font-size: 0.7em;
		margin-left: var(--default_margin);
		font-weight: 800;
		background-color: #cc1833;
		color: #e3e3e3;
		border-radius: 50%;
		padding: 7px;
	}
</style>
