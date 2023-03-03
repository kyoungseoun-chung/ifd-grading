import * as d3 from 'd3';

export const plot_d3 = (data: Array<number>, min: number, max: number, interval: number) => {
	const margin = { top: 20, right: 20, bottom: 30, left: 40 },
		width = 600 - margin.left - margin.right,
		height = 400 - margin.top - margin.bottom;

	const svg = d3
		.select('#graph_holder')
		.append('svg')
		.attr('width', width + margin.left + margin.right)
		.attr('height', height + margin.top + margin.bottom)
		.append('g')
		.attr('transform', 'translate(' + margin.left + ',' + margin.top + ')');

	const x = d3.scaleLinear().domain([min, max]).range([0, width]);

	svg
		.append('g')
		.attr('transform', `translate(0, ${height - margin.bottom})`)
		.call(d3.axisBottom(x).tickSizeOuter(0));

	const num_bin = (max - min) / interval;
	console.log(num_bin);

	const histogram = d3
		.bin()
		.domain([min - interval / 2, max + interval / 2])
		.value((d) => d)
		.thresholds(d3.range(min - interval / 2, max + interval / 2, interval));

	const bins = histogram(data);
	const max_bin = d3.max(bins, (d) => d.length) as number;

	const y = d3
		.scaleLinear()
		.domain([0, max_bin])
		.nice()
		.range([height - margin.bottom, margin.top]);

	svg.append('g').call(d3.axisLeft(y));

	svg
		.append('g')
		.selectAll('rect')
		.data(bins)
		.join('rect')
		.attr('x', (d) => x(d.x0 as number))
		.attr('width', (d) => {
			console.log(d.x0, d.x1);
			return Math.max(0, x(d.x1 as number) - x(d.x0 as number) - 1);
		})
		.attr('y', (d) => y(d.length))
		.attr('height', (d) => y(0) - y(d.length))
		.attr('fill', '#2f90ed');
	console.log(num_bin);
	console.log(bins);
	console.log(svg);
};
