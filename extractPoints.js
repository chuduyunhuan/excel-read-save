function sheet_from_array_of_arrays(data, opts) {
    var ws = {};
    //console.log(data);
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
//对象数组去重
Array.prototype.unique = function(field){
    var n = {},r=[]; //n为hash表，r为临时数组
    for(var i = 0, len = this.length; i < len; i++){
        if (!n[this[i][field]]){
            n[this[i][field]] = true; //存入hash表
            r.push(this[i][field]); //把当前数组的当前项push到临时数组里面
        }
    }
    return r;
};
//普通数组去重
Array.prototype.unique2 = function()
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
//创建workbook
function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}
var XLSX = require('xlsx');
var fs = require('fs'); 
//新版excel格式
var CONFIG_NEW = {    
    path: 'origin-files',
    fileName: '高铁.xlsx'
};
readFileNew(CONFIG_NEW);
// calculateCount(CONFIG_NEW);
function readFileNew(opts) {
    var path = opts.path || 'origin-files',
        fileName = opts.fileName;
    var filePath = path + '/' + fileName;
    //读
    var workbook = XLSX.readFile(filePath);
    //写
    var wb = new Workbook();
    var data = [];
    //取出所有sheet中的数据
    for(var sheet in workbook.Sheets) {
        var sheetName = workbook.Sheets[sheet];
        var fromto = sheetName['!ref'];
        data = data.concat(XLSX.utils.sheet_to_json(sheetName));
    }
    var result = data.map(function(obj) {
        var points = obj.lng.split(','),
            line = obj.line,
            line_section = obj.line_section;
        return {
            lng: points[0],
            lat: points[1],
            line: line,
            line_section: line_section
        };
    });
    saveJsonFile(result);
}
function saveJsonFile(data) {
    var path = 'processed-files';
    var name = 'railway.js';
    fs.writeFile(path+'/'+name, JSON.stringify(data), 'utf-8', (err) => {
        if (err) throw err;
        console.log('It\'s saved!');
    });
}
function calculateCount(opts) {
    var path = opts.path || 'processed-files',
        fileName = opts.fileName;
    var filePath = path + '/' + fileName;
    //读
    var workbook = XLSX.readFile(filePath);

    var data = [];
    //取出所有sheet中的数据
    for(var sheet in workbook.Sheets) {
        var sheetName = workbook.Sheets[sheet];
        var fromto = sheetName['!ref'];
        data = data.concat(XLSX.utils.sheet_to_json(sheetName));
    }
    return data.length;
}

