var fs = require('fs'), path = require('path'), URL = require('url');
var express = require('express'), app = express();
var sprintf = require('printj').sprintf;
var logit = require('./_logit');
var cors = require('./_cors');
var data = "a,b,c\n1,2,3".split("\n").map(function(x) { return x.split(","); });
var XLSX = require('xlsx');

/* helper to generate the workbook object */
function make_book() {
	var ws = XLSX.utils.aoa_to_sheet(data);
	var wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
	return wb;
}

function get_data(req, res, type) {
	var wb = make_book();
	/* send buffer back */
	res.status(200).send(XLSX.write(wb, {type:'buffer', bookType:type}));
}

function get_file(req, res, file) {
	var wb = make_book();
	/* write using XLSX.writeFile */
	XLSX.writeFile(wb, file);
	res.status(200).send("wrote to " + file + "\n");
}

function load_data(file) {
	var wb = XLSX.readFile(file);
	/* generate array of arrays */
	data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1});
	console.log(data);
}

function post_data(req, res) {
	var keys = Object.keys(req.files), k = keys[0];
	load_data(req.files[k].path);
	res.status(200).send("ok\n");
}

function post_file(req, res, file) {
	load_data(file);
	res.status(200).send("ok\n");
}
app.use(logit.mw);
app.use(cors.mw);
app.use(require('express-formidable')());
app.get('/', function(req, res, next) {
	var url = URL.parse(req.url, true);
	if(url.query.t) return get_data(req, res, url.query.t);
	else if(url.query.f) return get_file(req, res, url.query.f);
	res.status(403).end("Forbidden");
});
app.post('/', function(req, res, next) {
	var url = URL.parse(req.url, true);
	console.log('url:', url)
	if(url.query.f) return post_file(req, res, url.query.f);
	return post_data(req, res);
});

app.get('/createData', (req, res, next) => {
	var url = URL.parse(req.url, true);
	console.log('url:', url);
	var ws = XLSX.utils.json_to_sheet([
	  //{ S:1, h:2, e:3, e_1:4, t:5, J:6, S_1:7 },
	  { S:2, h:3, e:4, e_1:5, t:6, J:7, S_1:8 },
	  {hola: 100, adios: 200},
	  {...url.query}
	], {header:["S","h","e","e_1","t","J","S_1", "hola", "adios"]});

	var wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
	// XLSX.writeFile(wb, './out.xlsx');
	res.status(200).send(XLSX.write(wb, {type: 'buffer', bookType: 'xlsx'}));
})

app.post('/createCsv', (req, res, next) => {
	const body = req.body
	console.log('body:', body);

	// fs.writeFileSync(file, data[, options])

	const ws = XLSX.utils.json_to_sheet([
	  { employee_id: 'employee_id', clave: 'clave', valor: 'valor' },
	  {...body}
	], {header:[ "employee_id", "clave", "valor" ]});

	var wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
	// XLSX.writeFile(wb, './out.xlsx');
	res.status(200).send(XLSX.write(wb, {type: 'buffer', bookType: 'xlsx'}));
})

var port = +process.argv[2] || +process.env.PORT || 7262;
app.listen(port, function() { console.log('Serving HTTP on port ' + port); });
