var XLSX = require('xlsx');
var fileName = 'oct17.xlsx';

var workbook = XLSX.readFile(fileName);

var sheets = workbook.Sheets,
  sheetKey,
cellKey;

var indexName = "bank_details";

function addDocument(document, id, callback) {
  console.log("INSERT INTO table(name, address, phone) VALUES ('"+(document.name||'')+"', '"+(document.address||'')+"', "+(document.phone||null)+");");
 callback();
}

var dataToIndex = [];


for(sheetKey in sheets){
  var sheet = sheets[sheetKey];
  var dataKeys = {};
  var cellObj = undefined;
  var cellStart = undefined,
    cellEnd = undefined,
    cellMatch;
  for(cellKey in sheet){
    var cell = sheet[cellKey];
    if(cellMatch = cellKey.match(/^([a-zA-Z]+)1$/)){
      dataKeys[cellKey] = cell['v'];
      if(!cellStart){
        cellStart = cellMatch[1];
      }
      cellEnd = cellMatch[1];
    }else{
      var indexMatch = cellKey.match(/^([a-zA-Z]+)[0-9]+$/);
      if(indexMatch){
        var index = indexMatch[1] + '1';
        if(indexMatch[1] === cellStart){
          cellObj = {};
        }
        var value = cell['v'] || '';
        if(cell['v'].replace){
          cell['v'].replace("'","''");
        }
        cellObj[dataKeys[index].toLowerCase()] = value;
        if(indexMatch[1] === cellEnd){
          dataToIndex.push([cellObj['id'], cellObj]);
        }
      }
    }
  }
}


function iterateAndIndex(){
  var obj = dataToIndex.pop();
  if(!obj){
    return; 
  }
  addDocument(obj[1], obj[0], iterateAndIndex);
};
iterateAndIndex();
