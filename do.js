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

//新版excel格式
var CONFIG_NEW = {
    tableHead: ['小区名称','类型','LAC','CI','经度','纬度','cell_name,cell_nt,lac,ci,lon,lat'],
    usedCols: {
        GSM: ['小区中文名','类型','LAC','CI','经度','纬度'],
        TD: ['小区中文名','类型','LAC','CI','经度','纬度'],
        LTE: ['小区名称','类型','eNodeBID','CI','经度','纬度']
    },
    '类型': {
        GSM: '2G',
        TD: '3G',
        LTE: '4G'
    },
    allDis: '区域',
    path: 'origin-files',
    fileName: '交通枢纽.xlsx'
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

    //获取所有保障区域名称
    var allDis = data.map(obj => {
        return obj[CONFIG_NEW.allDis];
    });
    var uniqueDis = allDis.unique2();

    //每个保障区域作为一个sheet
    uniqueDis.map(val => {
        var sheetArr = [];
        for(var sheet in workbook.Sheets) {
            var sheetName = workbook.Sheets[sheet];
            var fromto = sheetName['!ref'];
            var inuse = CONFIG_NEW.usedCols[sheet];
            var sheetData = XLSX.utils.sheet_to_json(sheetName);
            sheetData.map(obj => {
                var allDis = obj[CONFIG_NEW.allDis];
                if(allDis != val) return;
                var temArr = [];
                var temStr = '';
                inuse.map(useVal => {
                    if(!obj[useVal]) {
                        if(useVal != '类型') {
                            temArr.push(' ');
                            temStr += '|';
                        }else{
                            temArr.push(CONFIG_NEW['类型'][sheet]);
                            temStr += CONFIG_NEW['类型'][sheet] + '|';
                        }
                    }else{
                        temArr.push(obj[useVal]);
                        temStr += obj[useVal] + '|';
                    }
                });
                temStr = temStr.replace(/\|$/,'');//替换最后一个|
                temArr.push(temStr);
                sheetArr.push(temArr);
            });
        }
        setSheetNew(sheetArr,val);
    });

    function setSheetNew(arr,sheetName){
        //设置表头
        var tableHead = CONFIG_NEW.tableHead;
        var sheetData = [tableHead].concat(arr);

        var ws = sheet_from_array_of_arrays(sheetData);
        var ws_name = sheetName;
        wb.SheetNames.push(ws_name);
        wb.Sheets[ws_name] = ws;
    }
    XLSX.writeFile(wb, 'processed-files/' + '处理_' + fileName);
    console.log('save done!');
    var originTotalCount = calculateCount(opts),
        processedTotalCount = calculateCount({fileName: '处理_' + fileName});
    console.log('原始数据共 ' + originTotalCount + ' 条');
    console.log('处理后数据共 ' + processedTotalCount + ' 条');
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

