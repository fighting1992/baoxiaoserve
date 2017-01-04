var express = require('express');
var app = express();
var https = require('https');
var fs = require('fs');
var bodyParser = require('body-parser');
// var cheerio = require('cheerio');
var iconv = require('iconv-lite');
var xlsx = require('node-xlsx');

var Excel = require('exceljs');


app.use("/", express.static(__dirname));
// 创建服务端









var workbook = new Excel.Workbook();
workbook.xlsx.readFile('./a.xlsx')
.then(function(data) {
    global.data= data;
});


var exec = require('child_process').exec,child;



// app.use(bodyParser.json()); // for parsing application/json
app.use(bodyParser.urlencoded({ extended: true }));
app.post('/download',function(request,response){
    response.set("Access-Control-Allow-Origin", "*");
    response.set("Access-Control-Allow-Headers", "Content-Type,Content-Length, Authorization, Accept,X-Requested-With");
    response.set("Access-Control-Allow-Methods","PUT,POST,GET,DELETE,OPTIONS");
    // 获取日期json
    // getDateJson(request,response);
            // response.download('./b.xlsx')
    
    a(request,response)
    
    // 生成xlsx
    function createXlsx () {
        response.send(typeof json);
        /* 
        **定义arrcont
        * 发票序号为顺序
        * 加班日期前端返回
        * 用餐人员前端返回
        * 晚餐固定
        * 金额固定
        * 备注 前端返回
        */
        var arrcont = [
            ["加班餐费明细单"],
            ['发票序号' , '加班时间' , '用餐人员' , '中餐/晚餐', '金额' , '备注'],
            ['1','2016年9月1日','jiao','晚餐','35',''],
            ['1','2016年9月1日','jiao','晚餐','35',''],
        ];
        // arrcont为表格数据 为数组类型
        var arr = [
            {
                name : '报销单' , 
                data : arrcont
            }
        ];
        response.send(arrcont);
        var file = xlsx.build(arr);
        fs.writeFileSync('user11.xlsx', file, 'binary');
    }









    // createXlsx();
});
app.get('/download',function(request,response){
    response.download('./'+request.query.timeName+'.xlsx')
        app.delete('./'+request.query.timeName+'.xlsx');
        child = exec('rm -rf ./'+request.query.timeName+'.xlsx',function(err,out) { 
        });
    
})

app.listen('7000',function(){
    console.log('app is start!');
})
    // 表格
function a (req, res) {
        

            var reqData=JSON.parse(req.body.reqData)
            var i = 1;
            var sun = 0;
            reqData.day.forEach(function(value, index, array) {
                value.forEach(function(value,index){
                    if(value.isWork) {

                         var row = value.isHoliday? 
                         [, i, reqData.year+'年'+(Math.floor(reqData.month)+1)+'月'+value.date_str+'日', reqData.name, '午餐/晚餐', 70 ,value.mark]
                         : [, i, reqData.year+'年'+(Math.floor(reqData.month)+1)+'月'+value.date_str+'日', reqData.name, '晚餐', 35 ,value.mark]
                         sun += row[5];
                        data.eachSheet(function(worksheet, sheetId) {
                            worksheet.getRow(i+2).values = row;
                            worksheet.getCell("A"+(i+2)).style=worksheet.getCell('A2').style;
                            worksheet.getCell("B"+(i+2)).style=worksheet.getCell('B2').style;
                            worksheet.getCell("C"+(i+2)).style=worksheet.getCell('C2').style;
                            worksheet.getCell("D"+(i+2)).style=worksheet.getCell('D2').style;
                            worksheet.getCell("E"+(i+2)).style=worksheet.getCell('E2').style;
                            worksheet.getCell("F"+(i+2)).style=worksheet.getCell('F2').style;
                        });
                       i++;
                    }
                })
            });


                var row = [, '合计', , , , sun ,]
                data.eachSheet(function(worksheet, sheetId) {
                    worksheet.getRow(i+2).values = row;
                    worksheet.getCell("A"+(i+2)).style=worksheet.getCell('A2').style;
                    worksheet.getCell("B"+(i+2)).style=worksheet.getCell('B2').style;
                    worksheet.getCell("C"+(i+2)).style=worksheet.getCell('C2').style;
                    worksheet.getCell("D"+(i+2)).style=worksheet.getCell('D2').style;
                    worksheet.getCell("E"+(i+2)).style=worksheet.getCell('E2').style;
                    worksheet.getCell("F"+(i+2)).style=worksheet.getCell('F2').style;
                });
                var timeName = new Date().getTime()+"";
              data.xlsx.writeFile('./'+timeName+'.xlsx')
                        .then(function() {
                        res.send(timeName)
                            
                        });
                        workbook.xlsx.readFile('./a.xlsx')
                        .then(function(data) {
                            global.data = data;
                        });
           
        
        
        
    };




    function getDateJson (request,response) {
        https.get('https://sp0.baidu.com/8aQDcjqpAAV3otqbppnN2DJv/api.php?query=2016%E5%B9%B410%E6%9C%88&co=&resource_id=6018&t=1480061849715&ie=utf8&oe=gbk&cb=op_aladdin_callback&format=json&tn=baidu&cb=&_=1479871811296' , function(res){

            // res.setEncoding('GBK');    //正常可转码为utf8 不支持gbk
            var arrBuf = [];
            var bufLength = 0;

            res.on('data',function(d){
                arrBuf.push(d);
                bufLength += d.length;
            })
            .on('end',function(){
                // Buffer.concat(数组, 数组长度);
               var chunkAll = Buffer.concat(arrBuf, bufLength);   
               var strJson = iconv.decode(chunkAll,'gbk');
               // 返回json格式
               response.send(strJson);
            });
        }).on('error', function(e) {
            console.log("Got error: " + e.message);
        });
    }
