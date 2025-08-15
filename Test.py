javascript:(function(){
function downloadXLSX(data){
 if(window.XLSX){
  var ws = XLSX.utils.aoa_to_sheet(data);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  var wbout = XLSX.write(wb, {bookType:"xlsx", type:"array"});
  var blob = new Blob([wbout], {type:"application/octet-stream"});
  var url = URL.createObjectURL(blob);
  var a = document.createElement("a");
  a.href = url; a.download = "pvalue.xlsx";
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
 } else {
  alert("SheetJS not found. Loading it now, click the bookmarklet again in 2 seconds...");
  var s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  document.head.appendChild(s);
 }
}
function scrapeTable(){
 var data=[];
 var table=document.querySelector("table");
 if(table){
  var rows=table.querySelectorAll("tr");
  rows.forEach(function(r){
   var cells=r.querySelectorAll("th,td");
   var row=[]; cells.forEach(function(c){row.push(c.innerText.trim())});
   data.push(row);
  });
  return data;
 }
 var grid=document.querySelector("[role='grid']");
 if(grid){
  var headers=[...grid.querySelectorAll("[role='columnheader']")].map(h=>h.innerText.trim());
  if(headers.length) data.push(headers);
  var rows=[...grid.querySelectorAll("[role='row']")].filter(r=>!r.querySelector("[role='columnheader']"));
  rows.forEach(function(r){
    var cells=[...r.querySelectorAll("[role='gridcell'],[role='cell']")].map(c=>c.innerText.trim());
    data.push(cells);
  });
  return data;
 }
 return null;
}
var data=scrapeTable();
if(!data||!data.length){alert("No table or grid found!");return;}
downloadXLSX(data);
})();
