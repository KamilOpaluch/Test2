javascript:(function(){
if(window.__TABLE_EXPORTING__)return;
window.__TABLE_EXPORTING__=true;
function T(s){return s==null?"":String(s).replace(/\s+/g," ").trim()}
function dl(b,n){var a=document.createElement("a");a.href=URL.createObjectURL(b);a.download=n;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),800)}
function toCSV(A){return A.map(r=>r.map(v=>'"'+String(v??"").replace(/"/g,'""')+'"').join(",")).join("\r\n")}
function uniq(h){var seen={};return h.map(x=>{x=x||"col";if(!seen[x]){seen[x]=1;return x}seen[x]+=1;return x+"("+seen[x]+")"})}
function modeLen(rows){var m={},best=0,val=0;rows.forEach(r=>{var L=r.length;if(L>1){m[L]=(m[L]||0)+1;if(m[L]>best){best=m[L];val=L}}});return val||Math.max(0,...rows.map(r=>r.length))}
function compressHeader(A){
 if(!A||A.length<2)return A;
 var header=A[0].slice();
 var body=A.slice(1).filter(r=>r.length>1);
 var w=modeLen(body);
 if(w && header.length>=w*2) header=header.slice(0,w); // remove repeats (e.g. 5x)
 if(w && header.length<w) header=header.concat(Array(w-header.length).fill(""));
 A[0]=uniq(header);
 var width=w||A[0].length;
 for(var i=1;i<A.length;i++){
   if(A[i].length<width)A[i]=A[i].concat(Array(width-A[i].length).fill(""));
   else if(A[i].length>width)A[i]=A[i].slice(0,width);
 }
 return A;
}
function pickBiggestContainer(rows){
 if(rows.length===0)return null;
 let best=null,bestCount=0;
 rows.forEach(r=>{
   let p=r.parentElement;
   while(p&&p!==document.body){
     const c=p.querySelectorAll('[role="row"]').length;
     if(c>bestCount){best=p;bestCount=c}
     p=p.parentElement;
   }
 });
 return best||rows[0].parentElement;
}
function collect(){
 var t=document.querySelector("table");
 if(t){
   var data=[];
   t.querySelectorAll("tr").forEach(tr=>{
     data.push(Array.from(tr.querySelectorAll("th,td")).map(td=>T(td.innerText)));
   });
   return data.filter(r=>r.length);
 }
 var grid=document.querySelector('[role="grid"]');
 if(grid){
   var data2=[],heads=Array.from(grid.querySelectorAll('[role="columnheader"]')).map(h=>T(h.innerText));
   if(heads.length)data2.push(heads);
   Array.from(grid.querySelectorAll('[role="row"]')).forEach(r=>{
     if(r.querySelector('[role="columnheader"]'))return;
     var cells=Array.from(r.querySelectorAll('[role="gridcell"],[role="cell"]')).map(c=>T(c.innerText));
     if(cells.length)data2.push(cells);
   });
   if(data2.length>0)return data2;
 }
 var rows=Array.from(document.querySelectorAll('[role="row"]'));
 if(rows.length){
   var container=pickBiggestContainer(rows);
   if(container){
     var hdr=Array.from(container.querySelectorAll('[role="row"]')).find(r=>r.querySelector('[role="columnheader"]'));
     var data3=[];
     if(hdr){
       var h=Array.from(hdr.querySelectorAll('[role="columnheader"],div,span')).map(x=>T(x.innerText));
       if(h.filter(Boolean).length)data3.push(h);
     }
     Array.from(container.querySelectorAll('[role="row"]')).forEach(r=>{
       if(r===hdr)return;
       var cells=Array.from(r.querySelectorAll('[role="gridcell"],[role="cell"],div[tabindex]')).map(c=>T(c.innerText));
       if(cells.some(v=>v!==""))data3.push(cells);
     });
     if(data3.length>0)return data3;
   }
 }
 return null;
}
function save(A){
 if(!A||!A.length){alert("No table-like data found.");window.__TABLE_EXPORTING__=false;return;}
 A=compressHeader(A);
 function doX(){
   try{
     var ws=XLSX.utils.aoa_to_sheet(A),wb=XLSX.utils.book_new();
     XLSX.utils.book_append_sheet(wb,ws,"Sheet1");
     var out=XLSX.write(wb,{bookType:"xlsx",type:"array"});
     dl(new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),"pvalue.xlsx");
   }catch(e){
     console.warn("XLSX failed, CSV fallback:",e);
     dl(new Blob([toCSV(A)],{type:"text/csv;charset=utf-8;"}),"pvalue.csv");
   }
   window.__TABLE_EXPORTING__=false;
 }
 if(window.XLSX){doX();}
 else{
   if(window.__XLSX_LOADING__){setTimeout(doX,1000);}
   else{
     window.__XLSX_LOADING__=true;
     var s=document.createElement("script");
     s.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
     s.onload=doX;
     s.onerror=function(){
       dl(new Blob([toCSV(A)],{type:"text/csv;charset=utf-8;"}),"pvalue.csv");
       window.__TABLE_EXPORTING__=false;
     };
     document.head.appendChild(s);
   }
 }
}
var data=collect();
save(data);
})();
