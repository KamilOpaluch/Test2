javascript:(function(){
function downloadXLSX(data){
 if(window.XLSX){
  // Remove repeated headers (pattern: each header appears 5 times)
  if(data.length>0){
    let header = data[0];
    let blockSize = Math.floor(header.length / (header.filter(h=>h.trim()!=="").length / (new Set(header).size)));
    // Simpler: assume each unique header repeats exactly 5 times
    const repeatCount = 5;
    let newHeader = [];
    for(let i=0; i<header.length; i+=repeatCount){
      newHeader.push(header[i]);
    }
    data[0] = newHeader;

    // Trim all rows to same width as new header
    data = data.map(r => r.slice(0,newHeader.length));
  }

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
