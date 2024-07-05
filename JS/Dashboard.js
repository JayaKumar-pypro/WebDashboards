//Global variables

//Get ready document
$(document).ready(function() {
	//Call display Function
	dataview('dashboardarea');
	//for layout
	let documentHeight = document.documentElement.offsetHeight;
	document.getElementById("menu").style.height = (documentHeight) + "px";
	document.getElementById("dashboardarea").style.height = (documentHeight) + "px";
	document.getElementById("tablearea").style.height = (documentHeight) + "px";

	//prepare data
	fetchAndConvertExcel();
});

//Function to fetch the Excel file and convert it to JSON
async function fetchAndConvertExcel() {
	const response = await fetch('Assets/MyMartDemoData.xlsx'); //path to map excel file
	const arrayBuffer = await response.arrayBuffer();
	const data = new Uint8Array(arrayBuffer);
	const workbook = XLSX.read(data, {
		type: 'array'
	});

	//Convert the first sheet to JSON
	const sheetName = workbook.SheetNames[0];
	const worksheet = workbook.Sheets[sheetName];
	const json = XLSX.utils.sheet_to_json(worksheet);

	//Store the JSON data
	JsonData = JSON.stringify(json, null, 2);
	processdata(JsonData);
}

//Process the data
function processdata(JsonData) {
	var parsedData = JSON.parse(JsonData);
	chart_1_donut(parsedData);
	chart_2_bar(parsedData);
	chart_3_bar(parsedData);
	chart_4_line(parsedData);
	Prepare_TableView(parsedData);
}

function dataview(type) {
	if (type == "dashboardarea") {
		document.getElementById("dashboardarea").style.display = "block";
		document.getElementById("DashboardClick").setAttribute("class", "tabmenu selectedMenu");
		document.getElementById("tablearea").style.display = "none";
		document.getElementById("TableClick").setAttribute("class", "tabmenu")
	} else if (type == "tablearea") {
		document.getElementById("dashboardarea").style.display = "none";
		document.getElementById("DashboardClick").setAttribute("class", "tabmenu");
		document.getElementById("tablearea").style.display = "block";
		document.getElementById("TableClick").setAttribute("class", "tabmenu selectedMenu");
	}
}

function chart_1_donut(JsonData) {
	//calculate region wise sales
	var region_sales_all = [];
	var regions = [];
	for (let i = 0; i < JsonData.length; i++) {
		regions.push(JsonData[i].Region);
	}
	var regions_uni = [...new Set(regions)];
	//Calculate the respective regions' Sales
	for (let r = 0; r < regions_uni.length; r++) {
		var temp_data = JsonData.filter(item => item['Region'] == regions_uni[r]);
		var sales_amount = 0;
		for (let rs = 0; rs < temp_data.length; rs++) {
			sales_amount += temp_data[rs].Sales;
		}
		var region_sales = {};
		region_sales['name'] = regions_uni[r];
		region_sales['y'] = sales_amount;
		region_sales_all.push(region_sales);
	}

	//Now build the chart
	console.log(region_sales_all)
	Highcharts.chart('chart_1_donut', {
		chart: {
			type: 'pie'
		},
		title: {
			text: 'Sales by Region',
			style: {
				fontSize: '12px',
				fontWeight: '500'
			}
		},
		credits: {
			enabled: false
		},
		plotOptions: {
			pie: {
				innerSize: '50%',
				depth: 45,

				dataLabels: {
					distance: 15,
					enabled: true,
					formatter: function() {
						return '<span>' + this.point.name + '<br>' + Highcharts.numberFormat(this.point.y, 2, '.', ',');
					}
				}
			}
		},
		series: [{
			name: 'Sales',
			data: region_sales_all
		}]
	});
	//End
}

function chart_2_bar(JsonData) {
	//calculate profit and sales
	var sales = 0;
	var profit = 0;
	for (let i = 0; i < JsonData.length; i++) {
		sales += JsonData[i].Sales;
		profit += JsonData[i].Profit;
	}

	//build chart
	Highcharts.chart('chart_2_vbar', {
		chart: {
			type: 'column',
			spacingTop: 10,
		},
		title: {
			text: 'Sales Vs Profit',
			style: {
				fontSize: '12px',
				fontWeight: '500'
			}
		},
		xAxis: {
			categories: ['Sales', 'Profit'],
			title: {
				text: null
			},
			labels: {
				style: {
					fontSize: '11px',
				}
			},
			gridLineWidth: 0,
			lineWidth: 1,
			lineColor: '#ccc'
		},
		yAxis: {
			min: 0,
			title: {
				text: 'Sum of Currency'
			},
			labels: {
				enabled: false,
				overflow: 'justify',
				style: {
					fontSize: '11px',
				},
				formatter: function() {
					return '$' + Highcharts.numberFormat(this.value, 1, '.', ',');
				}
			},
			gridLineWidth: 0,
			lineWidth: 0
		},
		tooltip: {
			enabled: true
		},
		plotOptions: {
			column: {
				borderRadius: '3%',
				dataLabels: {
					enabled: true,
					formatter: function() {
						return '$' + Highcharts.numberFormat(this.y, 2, '.', ',');
					}
				},
				colorByPoint: true,
			}
		},
		legend: {
			enabled: false
		},
		credits: {
			enabled: false
		},
		colors: ['#1ab2ff', '#39ac73'],
		series: [{
			name: 'Sales vs Profit',
			data: [sales, profit],
			pointWidth: 45
		}]
	});
	//end
}

function chart_3_bar(JsonData) {
	//Calculate Segment wise sales
	var Segment_sales = new Map();
	var Segment = [];
	for (let i = 0; i < JsonData.length; i++) {
		Segment.push(JsonData[i].Segment);
	}
	var Segment_uni = [...new Set(Segment)];
	//Calculate the respective Segment' Sales
	for (let r = 0; r < Segment_uni.length; r++) {
		var temp_data = JsonData.filter(item => item['Segment'] == Segment_uni[r]);
		var sales_amount = 0;
		for (let rs = 0; rs < temp_data.length; rs++) {
			sales_amount += temp_data[rs].Sales;
		}
		Segment_sales.set(Segment_uni[r], sales_amount);
	}
	//build chart
	Highcharts.chart('chart_3_vbar', {
		chart: {
			type: 'bar',
			spacingTop: 10,
		},
		title: {
			text: 'Sales by Category',
			style: {
				fontSize: '12px',
				fontWeight: '500'
			}
		},
		xAxis: {
			categories: [...new Set(Segment_sales.keys())],
			title: {
				text: null
			},
			labels: {
				style: {
					fontSize: '11px',
				}
			},
			gridLineWidth: 0,
			lineWidth: 1,
			lineColor: '#ccc'
		},
		yAxis: {
			min: 0,
			title: {
				text: 'Sum of Sales'
			},
			labels: {
				enabled: false,
				overflow: 'justify',
				style: {
					fontSize: '11px',
				},
				formatter: function() {
					return '$' + Highcharts.numberFormat(this.value, 2, '.', ',');
				}
			},
			gridLineWidth: 0,
			lineWidth: 0
		},
		tooltip: {
			enabled: true,
			formatter: function() {
				return this.series.name + ': $' + Highcharts.numberFormat(this.y, 0, '.', ',');
			}
		},
		plotOptions: {
			bar: {
				borderRadius: '3%',
				dataLabels: {
					enabled: true,
					formatter: function() {
						return '$' + Highcharts.numberFormat(this.y, 2, '.', ',');
					}
				},
				colorByPoint: true,
			}
		},
		legend: {
			enabled: false
		},
		credits: {
			enabled: false
		},
		series: [{
			name: 'Sales',
			data: [...new Set(Segment_sales.values())],
			pointWidth: 35,
		}]
	});
	//end
}

function chart_4_line(JsonData) {
	//Calculate State wise sales
	var Category = [];
	var Sales = [];
	var Profits = [];
	var State = [];
	for (let i = 0; i < JsonData.length; i++) {
		State.push(JsonData[i].State);
	}
	var State_uni = [...new Set(State)];
	//Calculate the respective State' Sales
	for (let r = 0; r < State_uni.length; r++) {
		var temp_data = JsonData.filter(item => item['State'] == State_uni[r]);
		var sales_amount = 0;
		var profit_amount = 0;
		for (let rs = 0; rs < temp_data.length; rs++) {
			sales_amount += temp_data[rs].Sales;
			profit_amount += temp_data[rs].Profit;
		}
		Category.push(State_uni[r]);
		Sales.push(sales_amount);
		Profits.push(profit_amount);
	}

	//build chart
	Highcharts.chart('chart_4_line', {
		chart: {
			type: 'spline',
			spacingTop: 10,
		},
		title: {
			text: 'Sales by State',
			style: {
				fontSize: '12px',
				fontWeight: '500'
			}
		},
		xAxis: {
			categories: Category,
			title: {
				text: null
			},
			labels: {
				style: {
					fontSize: '11px',
				}
			},
			gridLineWidth: 0,
			lineWidth: 1,
			lineColor: '#ccc'
		},
		yAxis: [{
			min: 0,
			title: {
				text: 'Sum of Sales'
			},
			labels: {
				enabled: false,
				overflow: 'justify',
				style: {
					fontSize: '11px',
				},
				formatter: function() {
					return '$' + Highcharts.numberFormat(this.value, 0, '.', ',');
				}
			},
			gridLineWidth: 0,
			lineWidth: 0
		}, {
			min: 0,
			opposite: true,
			title: {
				text: 'Sum of Profit'
			},
			labels: {
				enabled: false,
				overflow: 'justify',
				style: {
					fontSize: '11px',
				},
				formatter: function() {
					return '$' + Highcharts.numberFormat(this.value, 0, '.', ',');
				}
			},
			gridLineWidth: 0,
			lineWidth: 0
		}],
		tooltip: {
			enabled: true,
			formatter: function() {
				return this.series.name + ': $' + Highcharts.numberFormat(this.y, 0, '.', ',');
			}
		},
		plotOptions: {
			spline: {
				connectNulls: false,
				dataLabels: {
					enabled: true,
					formatter: function() {
						return '$' + Highcharts.numberFormat(this.y, 2, '.', ',');
					}
				},
			},
			column: {
				dataLabels: {
					enabled: true,
					formatter: function() {
						return '$' + Highcharts.numberFormat(this.y, 2, '.', ',');
					}
				},
			}
		},
		legend: {
			enabled: true,
		},
		credits: {
			enabled: false
		},
		series: [{
			yAxis: 0,
			type: 'column',
			lineWidth: 1,
			name: 'Sales',
			data: Sales,
		}, {
			color: 'green',
			yAxis: 1,
			type: 'spline',
			lineWidth: 1,
			name: 'Profit',
			data: Profits,
		}]
	});
	//end
}

function Prepare_TableView(JsonData) {
	$("#salesdetails").igGrid({
		autoGenerateColumns: false,
		renderCheckboxes: true,
		columns: [{
				headerText: "Region",
				key: "Region",
				dataType: "string",
				width: "15%"
			},
			{
				headerText: "State",
				key: "State",
				dataType: "string",
				width: "15%"
			},
			{
				headerText: "Category",
				key: "Category",
				dataType: "string",
				width: "15%"
			},
			{
				headerText: "Segment",
				key: "Segment",
				dataType: "string",
				width: "15%"
			},
			{
				headerText: "Sales",
				key: "Sales",
				dataType: "Currency",
				width: "15%"
			},
			{
				headerText: "Profit",
				key: "Profit",
				dataType: "Currency",
				width: "15%"
			}
		],
		dataSource: JsonData,
		dataSourceType: "json",
		height: "100%",
		width: "100%",
		tabIndex: 1,
		features: [{
			name: "Paging",
			pageSize: 15
		}, ]
	});
}
