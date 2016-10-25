Array.prototype.unique = function()
{
    var n = {},r=[]; //n为hash表，r为临时数组
    for(var i = 0; i < this.length; i++) //遍历当前数组
    {
        if (!n[this[i]]) //如果hash表中没有当前项
        {
            n[this[i]] = true; //存入hash表
            r.push(this[i]); //把当前数组的当前项push到临时数组里面
        }
    }
    return r;
}
//保存通用函数
function datenum(v, date1904) {
    if(date1904) v+=1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data, opts) {
    var ws = {};
    console.log(data);
    var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
    for(var R = 0; R != data.length; ++R) {
        for(var C = 0; C != data[R].length; ++C) {
            if(range.s.r > R) range.s.r = R;
            if(range.s.c > C) range.s.c = C;
            if(range.e.r < R) range.e.r = R;
            if(range.e.c < C) range.e.c = C;
            var cell = {v: data[R][C] };
            if(cell.v == null) continue;
            var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
            
            if(typeof cell.v === 'number') cell.t = 'n';
            else if(typeof cell.v === 'boolean') cell.t = 'b';
            else if(cell.v instanceof Date) {
                cell.t = 'n'; cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            }
            else cell.t = 's';
            
            ws[cell_ref] = cell;
        }
    }
    if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
    return ws;
}
//创建workbook
function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

var XLSX = require('xlsx');
var FIELD = {
    useSheet: 'Sheet1',
    tableHead: [],
    workbook: null
};
function readFile(path,col,fileName){
    var workbook = XLSX.readFile(path);
    FIELD.workbook = workbook;
    var sheet_name_list = workbook.SheetNames;
    var first_sheet_name = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[first_sheet_name];
    var colorNumArr = [];
    //计算总行数,这儿实际上计算的是总行数*列数,不过不影响后面
    var length = 0;
    for(var z in worksheet){
        length++;
    }
    //提取表头
    var tableHead = [];
    for(var z in worksheet){
        if(z[0] === '!') continue;
        var colNum = z.substring(1);
        if(colNum > 1) continue;
        tableHead.push(worksheet[z].v);
    }
    FIELD.tableHead.push(tableHead);
    //console.log(length);
    //提取主题内容
    for(var i=2; i<=length; ++i){
        if(!worksheet[col+i] || !worksheet[col+i].v) continue;
        colorNumArr.push(worksheet[col+i].v);
        //if(worksheet[z].v !== '色号') continue;
    }
    var result = colorNumArr.unique().sort();
    //console.log(result);
    onlySheet(result,worksheet,'B',length,fileName);
}
//多sheet
function multiSheets(arr,worksheet,col,length){
    var wb = new Workbook();
    arr.map(function(val){
        var sameColor = [];
        //换行
        var rowBreak = 0;
        var temArr = [];
        for(var z in worksheet){
            if(z[0] === '!') continue;
            var colNum = z.substring(1);
            if(!worksheet[col+colNum] || !worksheet[col+colNum].v) continue;
            if(worksheet[col+colNum].v !== val) continue;

            if(rowBreak === 0){
                rowBreak = colNum;
                temArr.push(worksheet[z].v);
            }else{
                if(colNum !== rowBreak){
                    rowBreak = colNum;
                    sameColor.push(temArr);
                    temArr = [];
                    temArr.push(worksheet[z].v);
                }else{
                    temArr.push(worksheet[z].v);
                }
            }
        }
        //for in 循环后最后一次结果
        sameColor.push(temArr);
        setSheet(sameColor,val);
        // saveExcel(sameColor,val,'kerry-processed/4776尺码明细.xls');
    });  
    function setSheet(arr,sheetName){
        var data = FIELD.tableHead.concat(arr);
        var ws = sheet_from_array_of_arrays(data);
        var ws_name = sheetName;
        wb.SheetNames.push(ws_name);
        wb.Sheets[ws_name] = ws;
    }  
    XLSX.writeFile(wb, 'kerry-processed/4776尺码明细_处理.xls');
}
//单sheet
function onlySheet(arr,worksheet,col,length,fileName){
    var wb = new Workbook();
    var sameColor = [];
    var sumArr = [];
    arr.map(function(val){
        //换行
        var rowBreak = 0;
        var temArr = [];
        for(var z in worksheet){
            if(z[0] === '!') continue;
            var colNum = z.substring(1);
            if(!worksheet[col+colNum] || !worksheet[col+colNum].v) continue;
            if(worksheet[col+colNum].v !== val) continue;

            if(rowBreak === 0){
                //换行
                sameColor.push([null]);
                rowBreak = colNum;
                temArr.push(worksheet[z].v);
            }else{
                if(colNum !== rowBreak){
                    rowBreak = colNum;
                    sameColor.push(temArr);
                    sumArr.push(temArr);
                    temArr = [];
                    temArr.push(worksheet[z].v);
                }else{
                    temArr.push(worksheet[z].v);
                }
            }
        }
        //for in 循环后最后一次结果
        sameColor.push(temArr);
        sumArr.push(temArr);
        //求和
        var target = addArr(sumArr);
        sumArr.length = 0;
        sameColor.push(['sum','--'].concat(target));
    }); 
    //console.log(sameColor);
    saveExcel(sameColor,'处理','kerry-processed/' + '处理_' + fileName);
}
function addArr(arr){
    var result = [];
    if(arr.length === 0) return result;
    for(var i=2; i<10; ++i){
        var sum = 0;
        arr.map(function(val){
            sum += val[i];
        });
        result.push(sum);
    }
    return result;
}
function saveExcel(arr,sheetName,fileName){
    var data = FIELD.tableHead.concat(arr);

    var wb = new Workbook(), 
        ws = sheet_from_array_of_arrays(data);
        ws_name = sheetName;

    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;
    XLSX.writeFile(wb, fileName);
}
function readDir(path){
    var fs = require('fs');
    fs.readdir(path,function(err,files){
        files.map(function(fileName){
            readFile(path + '/' + fileName, 'B', fileName);
        });
    });
}
//读取单个文件
// readFile('kerry-origin/4776尺码明细.xls','B','处理_4776尺码明细.xls');
//读取文件夹
readDir('kerry-origin');
