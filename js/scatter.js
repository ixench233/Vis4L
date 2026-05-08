"use strict";

var scatterCharts = {};
var scatterEntries = [];
var scatterLoaded = false;
var taxonomyColors = {};
window.vis4vlScatterCharts = scatterCharts;

// Colors follow Excel taxonomy "Cross-modal Task" block (#8B6B9D), light→dark purple
var scatterColors = [
	"#4A7A9C",
	"#A6A6A6",
	"#D9B830",
	"#DA711A",
	"#74AAA1",
	"#8B6B9D",
	"#D2AAAA"
];

var scatterChartConfigs = [
	{
		elementId: "contain",
		title: "SALON",
		field: "salon",
		schemaText: "SALON"
	},
	{
		elementId: "contain2",
		title: "Data Modality",
		field: "modality",
		schemaText: "Data Modality"
	},
	{
		elementId: "con3",
		title: "Visual Encoding",
		field: "vis_encoding",
		schemaText: "Visual Encoding"
	},
	{
		elementId: "contain4",
		title: "Interaction Technique",
		field: "interaction",
		schemaText: "Interaction Technique"
	}
];

var scatterTaskOrder = [
	"Model Interpretability",
	"Image-to-Text",
	"Text-to-Image",
	"Question Answering",
	"Grounded Dialogue",
	"Classification",
	"Annotation",
	"Retrieval",
	"Matching",
	"Grounding",
	"Multimodal Collaborative Analysis",
	"Other"
];

var scatterSymbols = [
	"circle",
	"path://M512 64a448 448 0 1 1-448 448 448 448 0 0 1 448-448m0-64a512 512 0 1 0 512 512 512 512 0 0 0-512-512z",
	"path://M512 68.266667c244.622222 0 443.733333 199.111111 443.733333 443.733333S756.622222 955.733333 512 955.733333 68.266667 756.622222 68.266667 512 267.377778 68.266667 512 68.266667z m0 91.022222c-194.56 0-352.711111 158.151111-352.711111 352.711111s158.151111 352.711111 352.711111 352.711111 352.711111-158.151111 352.711111-352.711111-158.151111-352.711111-352.711111-352.711111z m0 125.155555c25.031111 0 45.511111 20.48 45.511111 45.511112v136.533333h136.533333c25.031111 0 45.511111 20.48 45.511112 45.511111s-20.48 45.511111-45.511112 45.511111H557.511111v136.533333c0 25.031111-20.48 45.511111-45.511111 45.511112s-45.511111-20.48-45.511111-45.511112V557.511111H329.955556c-25.031111 0-45.511111-20.48-45.511112-45.511111s20.48-45.511111 45.511112-45.511111h136.533333V329.955556c0-25.031111 20.48-45.511111 45.511111-45.511112z",
	"diamond",
	"rect",
	"triangle"
];

$(document).ready(function () {
	$("#scatterModal").on("shown.bs.modal", initializeScatter);

	$(window).on("resize", function () {
		$.each(scatterCharts, function (_key, chart) {
			chart.resize();
		});
	});
});

function initializeScatter() {
	$.each(scatterChartConfigs, function (_i, config) {
		if (!scatterCharts[config.elementId])
			scatterCharts[config.elementId] = echarts.init(document.getElementById(config.elementId));
	});

	if (scatterLoaded) {
		renderAllScatterCharts();
		return;
	}

	$.when(
		$.getJSON("data/scatter.json?v=taxonomy-20260507"),
		$.getJSON("data/taxonomy_colors.json?v=taxonomy-20260507")
	).done(function (scatterResponse, colorResponse) {
		scatterEntries = scatterResponse[0] || [];
		taxonomyColors = colorResponse[0] || {};
		scatterColors = buildScatterPalette(taxonomyColors);
		scatterLoaded = true;
		renderAllScatterCharts();
	}).fail(function () {
		scatterColors = buildScatterPalette({});
	});
}

function renderAllScatterCharts() {
	appendScatterLegend();

	$.each(scatterChartConfigs, function (_i, config) {
		renderScatterChart(config);
	});
}

function renderScatterChart(config) {
	var chart = scatterCharts[config.elementId];
	var groupNames = getScatterTaskGroups(scatterEntries);
	var series = $.map(groupNames, function (taskName) {
		return {
			symbolSize: 20,
			name: taskName,
			type: "scatter",
			itemStyle: {
				opacity: 0.78
			},
			data: makeScatterSeriesData(scatterEntries, taskName, config)
		};
	});

	var symbolMap = makeScatterSymbolMap(scatterEntries, config.field);

	chart.setOption({
		title: {
			text: config.title,
			left: "center",
			top: 0
		},
		color: scatterColors,
		legend: {
			show: false,
			top: 20,
			textStyle: {
				fontSize: 10
			},
			data: groupNames
		},
		grid: {
			borderColor: "rgb(0,0,0)",
			show: true
		},
		tooltip: {
			padding: 10,
			backgroundColor: "rgba(255,255,255)",
			borderColor: "rgba(0,0,0,.5)",
			formatter: function (obj) {
				var value = obj.value;
				return "<b>Title</b>: " + escapeScatterHtml(value[2])
					+ "<br><b>" + escapeScatterHtml(config.schemaText) + "</b>: " + escapeScatterHtml(value[5])
					+ "<br><b>Major Category</b>: " + escapeScatterHtml(value[3])
					+ "<br><b>Minor Category</b>: " + escapeScatterHtml(value[4])
					+ "<br><b>Base Model</b>: " + escapeScatterHtml(value[6]);
			}
		},
		xAxis: {
			type: "value",
			gridIndex: 0,
			nameGap: 160,
			min: -6.2,
			max: 6.2,
			show: false,
			splitLine: {
				show: false
			},
			axisLabel: {
				formatter: "{value}"
			}
		},
		yAxis: {
			type: "value",
			gridIndex: 0,
			nameLocation: "end",
			nameGap: 20,
			min: -6.2,
			max: 6.2,
			show: false,
			splitLine: {
				show: false
			}
		},
		visualMap: [
			{
				bottom: 70,
				right: "10%",
				dimension: 5,
				categories: Object.keys(symbolMap),
				inRange: {
					symbol: symbolMap,
					symbolSize: 10
				},
				textStyle: {
					fontSize: 10
				},
				outOfRange: {
					symbol: symbolMap
				}
			}
		],
		series: series
	}, true);

	chart.off("click");
	chart.on("click", onScatterPointClick);
	chart.resize();
}

function makeScatterSeriesData(entries, taskName, config) {
	var filteredEntries = $.grep(entries, function (entry) {
		return getPrimaryScatterValue(entry.task) == taskName;
	}).sort(function (a, b) {
		return String(a.title).localeCompare(String(b.title));
	});

	return filteredEntries.map(function (entry) {
		var fieldValue = getPrimaryScatterValue(entry[config.field]);
		var jitter = getScatterCoordinateJitter(entry, config);
		return [
			Number(entry.x) + jitter.x,
			Number(entry.y) + jitter.y,
			entry.title,
			getPrimaryScatterValue(entry.salon),
			taskName,
			fieldValue,
			formatScatterValue(entry.model),
			entry.id
		];
	});
}

function getScatterCoordinateJitter(entry, config) {
	var hash = scatterHash(config.field + ":" + entry.id);
	return {
		x: ((hash % 1000) / 1000 - 0.5) * 0.08,
		y: (((hash >>> 12) % 1000) / 1000 - 0.5) * 0.08
	};
}

function getScatterTaskGroups(entries) {
	var available = {};

	$.each(entries, function (_i, entry) {
		available[getPrimaryScatterValue(entry.task)] = true;
	});

	return $.grep(scatterTaskOrder, function (task) {
		return available[task];
	}).concat(Object.keys(available).sort().filter(function (task) {
		return scatterTaskOrder.indexOf(task) == -1;
	}));
}

function makeScatterSymbolMap(entries, field) {
	var values = [];

	$.each(entries, function (_i, entry) {
		var value = getPrimaryScatterValue(entry[field]);
		if (values.indexOf(value) == -1)
			values.push(value);
	});

	values.sort();

	var map = {};
	$.each(values, function (i, value) {
		map[value] = scatterSymbols[i % scatterSymbols.length];
	});

	return map;
}

function appendScatterLegend() {
	var table = $("#scatterLegendTable");
	table.empty();

	var groups = getScatterTaskGroups(scatterEntries);
	for (var i = 0; i < groups.length; i += 4) {
		var row = $("<tr></tr>");
		for (var j = 0; j < 4 && i + j < groups.length; j++) {
			var color = scatterColors[(i + j) % scatterColors.length];
			row.append($("<td style=\"height:55px;width:50px;\"></td>")
				.append($("<div class=\"scatter-legend-swatch\"></div>").css("background-color", color)));
			row.append($("<td></td>").text(groups[i + j]));
		}
		table.append(row);
	}
}

function buildScatterPalette(colors) {
	var orderedKeys = ["modality", "model", "vis_encoding", "interaction", "evaluation", "task", "domain"];
	var fallback = ["#4A7A9C", "#A6A6A6", "#D9B830", "#DA711A", "#74AAA1", "#8B6B9D", "#D2AAAA"];
	var palette = [];

	$.each(orderedKeys, function (_index, key) {
		var color = colors[key];
		if (color && palette.indexOf(color) === -1)
			palette.push(color);
	});

	if (!palette.length)
		return fallback;

	$.each(fallback, function (_index, color) {
		if (palette.indexOf(color) === -1)
			palette.push(color);
	});

	return palette;
}

function onScatterPointClick(params) {
	if (!params || !params.value || !params.value[7])
		return;

	var id = params.value[7];
	$("#scatterModal").modal("hide");
	setTimeout(function () {
		displayEntryDetails(id);
	}, 250);
}

function getPrimaryScatterValue(value) {
	var values = normalizeScatterValues(value);
	return values.length ? values[0] : "Other";
}

function normalizeScatterValues(value) {
	if ($.isArray(value))
		return $.grep($.map(value, cleanScatterValue), Boolean);

	value = cleanScatterValue(value);
	if (!value)
		return [];

	return [value];
}

function cleanScatterValue(value) {
	value = $.trim(String(value || ""));
	if (!value || value == "undefined" || value == "null")
		return "";
	return value;
}

function formatScatterValue(value) {
	if ($.isArray(value))
		return value.join(", ");
	return value || "";
}

function scatterHash(value) {
	var hash = 2166136261;
	value = String(value || "");

	for (var i = 0; i < value.length; i++) {
		hash ^= value.charCodeAt(i);
		hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
	}

	return hash >>> 0;
}

function escapeScatterHtml(value) {
	return String(value == null ? "" : value)
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/"/g, "&quot;")
		.replace(/'/g, "&#39;");
}
