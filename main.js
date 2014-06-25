var rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined",
	drop = document.getElementById('drop'),
	XLSX = window.XLSX,
	output,
	to_sort = [],
	activity_names = [];

// Randomize array element order in-place. Using Fisher-Yates shuffle algorithm.
function shuffle(array) {
	var m = array.length,
		t, i;

	// While there remain elements to shuffleâ€¦
	while (m) {

		// Pick a remaining elementâ€¦
		i = Math.floor(Math.random() * (m--));

		// And swap it with the current element.
		t = array[m];
		array[m] = array[i];
		array[i] = t;
	}

	return array;
}

function count_spots() {

	var span = document.querySelector('#info-area #spots'),
		inputs = document.querySelectorAll('#info-area input'),
		spots = 0,
		kids = parseFloat(document.querySelector('#info-area #total').innerHTML, 10);

	Array.prototype.forEach.call(inputs, function (input) {
		spots = spots + parseFloat(input.value, 10);
	});

	if (spots > kids || spots < kids) {
		span.innerHTML = "<span style='color: red'>" + spots + "</span>";
	} else {
		span.innerHTML = spots;
	}
}

function fixdata(data) {
	var o = "",
		l = 0,
		w = 10240;
	for (; l < data.byteLength / w; ++l) {
		o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
	}
	o += String.fromCharCode.apply(null, new Uint8Array(data.slice(o.length)));
	return o;
}

function wb_data(workbook) {
	var result = {};
	workbook.SheetNames.forEach(function (sheetName) {
		var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
		if (roa.length > 0) {
			result[sheetName] = roa;
		}
	});
	return result;
}

function sheet_from_array_of_arrays(data) {
	var ws = {};
	var range = {
		s: {
			c: 10000000,
			r: 10000000
		},
		e: {
			c: 0,
			r: 0
		}
	};
	for (var R = 0; R != data.length; ++R) {
		for (var C = 0; C != data[R].length; ++C) {
			if (range.s.r > R) {
				range.s.r = R;
			}
			if (range.s.c > C) {
				range.s.c = C;
			}
			if (range.e.r < R) {
				range.e.r = R;
			}
			if (range.e.c < C) {
				range.e.c = C;
			}
			var cell = {
				v: data[R][C]
			};
			if (cell.v == null) continue;
			var cell_ref = XLSX.utils.encode_cell({
				c: C,
				r: R
			});

			if (typeof cell.v === 'number') {
				cell.t = 'n';
			} else if (typeof cell.v === 'boolean') {
				cell.t = 'b';
			} else if (cell.v instanceof Date) {
				cell.t = 'n';
				cell.z = XLSX.SSF._table[14];
				cell.v = datenum(cell.v);
			} else {
				cell.t = 's';
			}

			ws[cell_ref] = cell;
		}
	}
	if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
	return ws;
}

function process_wb(wb) {
	var fieldsets = document.querySelectorAll('#info-area fieldset');
	
	output = wb_data(wb);

	console.log("output:", output);

	// Make sortable arrays
	output[Object.keys(output)[0]].forEach(function (item, index) {
		var temp_arr = [],
			cloned_item = JSON.parse(JSON.stringify(item));

		delete cloned_item["First Name"];
		delete cloned_item["Last Name"];
		delete cloned_item["RC"];

		// console.log("item: ", cloned_item);

		for (var i = 0, temp = Object.keys(cloned_item); i < temp.length; i++) {
			temp_arr.push(cloned_item[temp[i]]);
		}

		if (temp_arr.length !== 0) {
			to_sort.push([index, temp_arr]);

			// Get the activities while we're here
			if (activity_names.length < Object.keys(cloned_item).length) {
				activity_names = Object.keys(cloned_item);
			}
		}
	});

	console.log("to_sort", to_sort);

	document.querySelector('#info-area h4 span').innerHTML = to_sort.length;

	activity_names.forEach(function (activity, index) {
		var fragment = document.createDocumentFragment(),
			temp = document.createElement('div');

		temp.innerHTML = '<div class="pure-control-group"><label for="activity">' + activity + '</label><input type="number" id="' + index + '" value="30"></div>';

		while (temp.firstChild) {
			fragment.appendChild(temp.firstChild);
		}

		if (index % 2 !== 0) {
			fieldsets[0].appendChild(fragment);
		} else {
			fieldsets[1].appendChild(fragment);
		}

	});

	count_spots();
	Array.prototype.forEach.call(document.querySelectorAll('#info-area input'), function (elem) {
		elem.addEventListener('onchange', count_spots);
	});

}

function make_wb() {

	var limits = [],
		final_sort = [],
		new_sheet_array = [];

	Array.prototype.forEach.call(document.querySelectorAll('#info-area input'), function (limit_el) {

		limits[limit_el.getAttribute('id')] = parseFloat(limit_el.value, 10);

	});

	console.log('Limits', limits);

	function sort_by_choice(a, b) {
		return a[1][k] - b[1][k];
	}
	
	// Sort said arrays
	for (var k = 0; k < limits.length; k++) {

		shuffle(to_sort);

		to_sort.sort(sort_by_choice);

		var temp_arr = [];

		while (temp_arr.length !== limits[k] && to_sort.length > 0) {
			temp_arr.push(to_sort.shift());
		}

		final_sort[k] = temp_arr;

	}

	console.log("Final sort: ", final_sort);

	// Make data ready to be put in worksheet
	for (var i = 0; i < final_sort.length; i++) {

		new_sheet_array.push([activity_names[i]]);
		new_sheet_array.push(["RC", "Last Name", "First Name"]);

		final_sort[i].forEach(function (item) {

			var person = output[Object.keys(output)[0]][item[0]];
			// console.log("Item:", item[0]);
			// console.log("person:", person);

			new_sheet_array.push([person["RC"], person["Last Name"], person["First Name"]]);

		});

		new_sheet_array.push([null]);

	}

	console.log("built sheet", new_sheet_array);

	function Workbook() {
		if (!(this instanceof Workbook)) {
			return new Workbook();
		}
		this.SheetNames = [];
		this.Sheets = {};
	}

	var ws_name = Object.keys(output)[0],
		sorted_wb = new Workbook(),
		sorted_ws = sheet_from_array_of_arrays(new_sheet_array);

	sorted_wb.SheetNames.push(ws_name);
	sorted_wb.Sheets[ws_name] = sorted_ws;

	var wbout = XLSX.write(sorted_wb, {
		bookType: 'xlsx',
		bookSST: true,
		type: 'binary'
	});

	function s2ab(s) {
		var buf = new ArrayBuffer(s.length);
		var view = new Uint8Array(buf);
		for (var i = 0; i != s.length; ++i) {
			view[i] = s.charCodeAt(i) & 0xFF;
		}
		return buf;
	}

	saveAs(
		new Blob(
				[s2ab(wbout)], {
				type: "application/octet-stream"
			}
		),
		"Sorted Activities.xlsx");

}

document.getElementById('save').addEventListener('click', make_wb);

function handleDrop(e) {
	e.stopPropagation();
	e.preventDefault();
	var files = e.dataTransfer.files;
	var i, f;
	for (i = 0, f = files[i]; i != files.length; ++i) {
		var reader = new FileReader();
		reader.onload = function (e) {
			var data = e.target.result,
				wb;
			if (rABS) {
				wb = XLSX.read(data, {
					type: 'binary'
				});
			} else {
				var arr = fixdata(data);
				wb = XLSX.read(btoa(arr), {
					type: 'base64'
				});
			}
			process_wb(wb);
		};
		if (rABS) {
			reader.readAsBinaryString(f);
		} else {
			reader.readAsArrayBuffer(f);
		}
	}
}

function handleDragover(e) {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
}
if (drop.addEventListener) {
	drop.addEventListener('dragenter', handleDragover, false);
	drop.addEventListener('dragover', handleDragover, false);
	drop.addEventListener('drop', handleDrop, false);
}