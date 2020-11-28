const express = require('express')
const app = express()
const port = 3000

app.listen(port, () => console.log(`Example app listening on port ${port}!`))

app.use(express.json()) // for parsing application/json
app.use(express.urlencoded({ extended: true })) // for parsing application/x-www-form-urlencoded

app.post('/xlsx', function (req, res, next) {
	request(req.body.url, {encoding: null}, function(err, requestres, data) {
		if(err || requestres.statusCode !== 200) return res.json({'error': '请求链接无法访问'});
		try {
			var wb = XLSX.read(data, {type:'buffer'});
			var sendData = []
			if(wb.SheetNames.length > 0){
				for (let i = 0; i < wb.SheetNames.length; i++){
					let ws = wb.Sheets[wb.SheetNames[i]];
					let item = {
						title: wb.SheetNames[i],
						html: XLSX.utils.sheet_to_html(ws, {blankrows:false})
					}
					sendData.push(item)
				}
			}
			res.send(sendData)
		} catch (error) {
			return res.json({'error': 'xlsx 解析失败'})
		}
	});
})

