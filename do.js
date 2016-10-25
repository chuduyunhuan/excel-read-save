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
//创建workbook
function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}
var XLSX = require('xlsx');
//配置文件
var FILED_HIERARCHY = {
    hieraarchyField: '物理热点(需与小区表中热点名称对应)',
    disField: '物理热点（名称不可重复）',
    hotSpotField: '层次关系',
    detialedInfo: {
        GSM: '2G小区',
        TD: '3G小区',
        LTE: '4G小区'
    },
    filterField: 'J',
    doubleLine: 'E',
    notLine: 'F',
    fieldObjArr: [
        {'物理热点（名称不可重复）': 'J'},
        {'小区名称': 'A'},
        {'小区类型': ' '},
        {'LAC': 'B'},
        {'CI': 'C'},
        {'经度': 'G'},
        {'纬度': 'H'},
        {'方向角': 'E'},
        {'载频数': 'I'},
        {'站类型（室内/宏站/应急车）': 'F'}
    ],
    disType: '',
    fieldArr: ['物理热点（名称不可重复）','小区名称',' 小区类型','LAC','CI','经度','纬度','方向角','载频数','基站类型（室内/宏站/应急车）','拼接后'],
};
//层次关系
function readHieraarchy(opts){
    var path = opts.path || 'origin-files',
        fileName = opts.fileName;
        col = opts.colTarget || 'D',
        rowBegin = opts.rowBegin || 3,
        rowEnd = opts.rowEnd || 14;

    var filePath = path + '/' + fileName;
    var hotSpotArr = [];
    var workbook = XLSX.readFile(filePath);
    var hieraarchySheet = workbook.Sheets[FILED_HIERARCHY.hotSpotField];
    for(var i=rowBegin; i<rowEnd; ++i){
        if(!hieraarchySheet[col+i]) continue;
        var hotSpot = hieraarchySheet[col+i].v;
        if(!hotSpot) continue;
        hotSpotArr.push(hotSpot);
    }
    readDetailedInfo(workbook,hotSpotArr,fileName);
}
//获取物理热点下的小区详情数据
function readDetailedInfo(workbook,arr,fileName){
    var wb = new Workbook();
    arr.map(function(val){
        var resultArr = [];
        var disTypeObj = FILED_HIERARCHY.detialedInfo;
        for(var name in disTypeObj){
            var disType = disTypeObj[name];
            var worksheet = workbook.Sheets[disType];
            FILED_HIERARCHY.fieldObjArr[2]['小区类型'] = disType.substring(0,2);
            var length = 0;
            var rowBreak = 0;
            for(var z in worksheet){
                var rowNum = z.substring(1);
                if(rowNum !== rowBreak){
                    length++;
                    rowBreak = rowNum;
                }
            }
            //console.log(FILED_HIERARCHY.fieldObjArr[2],length);
            //取列号
            var colNumArr = FILED_HIERARCHY.fieldObjArr.map(function(obj){
                for(var name in obj){
                    return obj[name];
                }
            });
            //console.log(colNumArr);
            //取数据
            for(var i=2; i<length; ++i){
                var temArr = [];
                var temStr = '';
                var done = false;
                colNumArr.map(function(field){
                    if(!worksheet[FILED_HIERARCHY.filterField + i]) return;
                    //console.log(val,worksheet[FILED_HIERARCHY.filterField + i].v);
                    if(worksheet[FILED_HIERARCHY.filterField + i].v !== val) return;
                    if(field.indexOf('2G') !== -1 || field.indexOf('3G') !== -1 ||field.indexOf('4G') !== -1){
                        temStr += FILED_HIERARCHY.fieldObjArr[2]['小区类型'] + '|';
                        // if(worksheet[field+i]){
                            temArr.push(FILED_HIERARCHY.fieldObjArr[2]['小区类型']);
                        // }
                        return;
                    }
                    if(!worksheet[field+i]) {
                        return;
                    }

                    if(field !== FILED_HIERARCHY.filterField){
                        temStr += worksheet[field+i].v + '|';
                        if(field === FILED_HIERARCHY.doubleLine){
                            temStr += '|';
                        }
                        if(field === FILED_HIERARCHY.notLine){
                            temStr = temStr.substring(0,temStr.length-1);
                        }
                    }
                    temArr.push(worksheet[field+i].v);
                    done = true;
                });
                if(done){
                    temArr.push(temStr);
                    resultArr.push(temArr);
                }
            }
        }
        setSheet(resultArr,val);
    });
    function setSheet(arr,sheetName){
        //设置表头
        var tableHead = FILED_HIERARCHY.fieldArr;
        var data = [tableHead].concat(arr);
        var ws = sheet_from_array_of_arrays(data);
        var ws_name = sheetName;
        wb.SheetNames.push(ws_name);
        wb.Sheets[ws_name] = ws;
    }
    XLSX.writeFile(wb, 'processed-files/' + '处理_' + fileName);
}


var path = 'origin-files',
    fileName = '-集中性能区域 保障热点小区资源信息表_20160930.xlsx',
    colTarget = 'D',
    rowBegin = 3,
    rowEnd = 14;
var OPTIONS_CONFIG = {
    path: path,
    fileName: fileName,
    // colTarget: colTarget,
    // rowBegin: rowBegin,
    // rowEnd: rowEnd
};
readHieraarchy(OPTIONS_CONFIG);
