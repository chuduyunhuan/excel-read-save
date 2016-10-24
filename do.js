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
    commonField: {
        name: '小区名称',
        type: '',
        lac: 'LAC',
        ci: 'CI',
        lng: '经度',
        lat: '纬度',
        direction: '方向角',
        frequency: '载频数',
        cellType: '基站类型（室内/宏站/应急车）'
    },
    '层次关系': {},
    '2G小区': {},
    '3G小区': {},
    '4G小区': {}
};
//层次关系
function readHieraarchy(path,col,rowBegin,rowEnd){
    var hotSpotArr = [];
    var workbook = XLSX.readFile(path);
    var hieraarchySheet = workbook.Sheets[FILED_HIERARCHY.hotSpotField];
    for(var i=rowBegin; i<rowEnd; ++i){
        if(!hieraarchySheet[col+i]) continue;
        var hotSpot = hieraarchySheet[col+i].v;
        if(!hotSpot) continue;
        hotSpotArr.push(hotSpot);
    }
    readDetailedInfo(workbook,hotSpotArr);
}
//获取详情数据
function readDetailedInfo(workbook,arr){    
    arr.map(function(val){   
        var resultArr = [];     
        for(var name in FILED_HIERARCHY.detialedInfo){
            FILED_HIERARCHY.commonField.type = FILED_HIERARCHY.detialedInfo[name].substring(0,2);
            var worksheet = workbook.Sheets[FILED_HIERARCHY.detialedInfo[name]];
            //获取最终保存所需字段的列号
            var colArr = []; 
            //过滤字段提取,这儿可以优化
            var colDisField = '';
            var length = 0;
            for (var z in worksheet) {
                length++;
                if(z[0] === '!') continue;
                var temObj = FILED_HIERARCHY.commonField;
                for(var field in temObj) {
                    if(worksheet[z].v === FILED_HIERARCHY.disField){
                        colDisField = z[0];
                        continue;
                    }
                    if(worksheet[z].v !== temObj[field]) continue;
                    colArr.push(z[0]);
                }
            }
             //console.log(colArr);
            // console.log(colDisField);
            resultArr.push(extractFieldVal(colArr,colDisField,val,worksheet,length));
        }
        //console.log(resultArr);
    });    
}
//根据行列号提取数据
function extractFieldVal(colArr,colDisField,val,worksheet,length){
    var result = [];
    for(var i=2; i<length; ++i){
        if(!worksheet[colDisField+i]) continue;
        var hotSpot = worksheet[colDisField+i].v;
        if(!hotSpot || hotSpot !== val) continue;
        //构造所需字符串格式
        var temStr = '';
        for(var j=0,lenJ=colArr.length; j<lenJ; ++j){
            if(!worksheet[colArr[j] + i]) continue;
            var temVal = worksheet[colArr[j] + i].v;
            temStr += temVal + '|';
        }
        result.push(temStr);
    }
    console.log(result);
    // for (var z in worksheet) {
    //     if(z[0] === '!') continue;
    //     var temObj = FILED_HIERARCHY.commonField;
    //     for(var field in temObj) {
    //         if(worksheet[z].v === temObj[field]) continue;
    //         if()
    //     }
    // }
}
function readExcel1(path){
    var workbook = XLSX.readFile(path);
    var sheet_name_list = workbook.SheetNames;
    // console.log(sheet_name_list);
    var hieraarchySheet = workbook.Sheets['层次关系'];
    console.log(hieraarchySheet['D34'].v);
    // for(var name in hieraarchySheet){
    //     if(name[0] === '!') continue;
    //     //找到物理热点对应的列编号
    //     var hotSpotName = hieraarchySheet[name].v;
    //     console.log(hotSpotName);
    //     if(hotSpotName !== FILED_HIERARCHY.hieraarchyField) continue;
    //     console.log(name[0]);
    //     //console.log(JSON.stringify(hieraarchySheet[name].v));
    // }
    // console.log(hieraarchySheet);
    // sheet_name_list.forEach(function(y) {
    //     // console.log(y);
    //   // iterate through sheets 
    //   var worksheet = workbook.Sheets[y];
    //   // for (z in worksheet) {
    //   //   // all keys that do not begin with "!" correspond to cell addresses 
    //   //   if(z[0] === '!') continue;
    //   //   console.log(y + "!" + z + "=" + JSON.stringify(worksheet[z].v));
    //   // }
    // });
}
//保存
function saveExcel(){

}
//readExcel('origin-files/test.xlsx');
readHieraarchy('origin-files/test.xlsx','D',3,14);
