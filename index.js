const express = require('express')
const app = express()
const port = 3000



app.listen(port, () => console.log(`Example app listening on port ${port}!`))


var XLSX = require('xlsx'), request = require('request');
request('https://xxx/testexcle.xlsx', {encoding: null}, function(err, res, data) {
	if(err || res.statusCode !== 200) return;

	/* data is a node Buffer that can be passed to XLSX.read */
	var wb = XLSX.read(data, {type:'buffer'});
	var ws = wb.Sheets[wb.SheetNames[0]];
	app.get('/', (req, res) => res.send(XLSX.utils.sheet_to_html(ws, {blankrows:false})))
	/* DO SOMETHING WITH workbook HERE */
});

app.use(express.json()) // for parsing application/json
app.use(express.urlencoded({ extended: true })) // for parsing application/x-www-form-urlencoded


app.post('/xlsx', function (req, res, next) {
	request(req.body.url, {encoding: null}, function(err, requestres, data) {
		if(err || requestres.statusCode !== 200) return res.json({'error': '请求链接无法访问'});
		try {
			var wb = XLSX.read(data, {type:'buffer'});
			var ws = wb.Sheets[wb.SheetNames[0]];
			res.send(XLSX.utils.sheet_to_html(ws, {blankrows:false}))
		} catch (error) {
			return res.json({'error': 'xlsx 解析失败'})
		}
	});
})

