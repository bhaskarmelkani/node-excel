var XLSX = require('xlsx');
var elasticsearch = require('elasticsearch');
var fileName = ''; // Enter file name.

var workbook = XLSX.readFile(fileName);

var sheets = workbook.Sheets,
  sheetKey,
cellKey;

var elasticClient = new elasticsearch.Client({
  host: 'localhost:9200',
  log: 'info'
});

var indexName = "";  // Name for index.
var id = '';  // Id for the document.

function addDocument(document, id, callback) {
  try{
  if(id){
    return elasticClient.index({
      index: indexName,
      type: "spreadsheet",
      id: id,
      body: document
    }, function(){
      if(callback){
        callback(); 
      }
    });
  }else{
    console.log('Document: ' + JSON.stringify(document));
    return elasticClient.index({
      index: indexName,
      type: "spreadsheet",
      body: document
    }, function(){
      if(callback){
        callback(); 
      }
    });
  }
  }catch(e){
    console.log('Exception');
    console.log(e);
  }

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
        cellObj[dataKeys[index]] = cell['v'];
        if(indexMatch[1] === cellEnd){
          dataToIndex.push([cellObj[id], cellObj]);
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
