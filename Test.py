(async function(){
  try {
    // ---------- helpers ----------
    const S = s => (s==null ? "" : String(s).replace(/\s+/g," ").trim());
    const dl = (blob, name) => {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(()=>URL.revokeObjectURL(a.href), 500);
    };
    const toCSV = A => A.map(r => r.map(v => '"'+String(v??"").replace(/"/g,'""')+'"').join(",")).join("\r\n");
    const pickBiggestContainer = rows => {
      if(!rows.length) return null;
      let best=null, bestCount=0;
      rows.forEach(r=>{
        let p=r.parentElement;
        while(p && p!==document.body){
          const c=p.querySelectorAll('[role="row"]').length;
          if(c>bestCount){best=p;bestCount=c}
          p=p.parentElement;
        }
      });
      return best || rows[0].parentElement;
    };

    // Gather same-origin documents (page + iframes we can access)
    const getDocs = () => {
      const docs=[document];
      const frames = Array.from(window.frames||[]);
      for (const f of frames) {
        try { if (f.document) docs.push(f.document); } catch(e) { /* cross-origin; skip */ }
      }
      return docs;
    };

    // Try to collect table-like data from a document
    const collectFromDoc = (doc) => {
      // 1) Native <table>
      const t = doc.querySelector("table");
      if (t) {
        const data=[];
        t.querySelectorAll("tr").forEach(tr=>{
          data.push(Array.from(tr.querySelectorAll("th,td")).map(td=>S(td.innerText)));
        });
        const out = data.filter(r=>r.length);
        if (out.length) return out;
      }
      // 2) ARIA grid
      const grid = doc.querySelector("[role='grid']");
      if (grid) {
        const data2=[];
        const heads = Array.from(grid.querySelectorAll("[role='columnheader']")).map(h=>S(h.innerText));
        if (heads.length) data2.push(heads);
        Array.from(grid.querySelectorAll("[role='row']")).forEach(r=>{
          if (r.querySelector("[role='columnheader']")) return;
          const cells = Array.from(r.querySelectorAll("[role='gridcell'],[role='cell']")).map(c=>S(c.innerText));
          if (cells.length) data2.push(cells);
        });
        if (data2.length) return data2;
      }
      // 3) Generic div grid (role=row, tabindex, etc.)
      const rows = Array.from(doc.querySelectorAll("[role='row']"));
      if (rows.length) {
        const container = pickBiggestContainer(rows);
        if (container) {
          const hdr = Array.from(container.querySelectorAll("[role='row']")).find(r=>r.querySelector("[role='columnheader']"));
          const data3=[];
          if (hdr) {
            const h = Array.from(hdr.querySelectorAll("[role='columnheader'],div,span")).map(x=>S(x.innerText)).filter(Boolean);
            if (h.length) data3.push(h);
          }
          Array.from(container.querySelectorAll("[role='row']")).forEach(r=>{
            if (r===hdr) return;
            const cells = Array.from(r.querySelectorAll("[role='gridcell'],[role='cell'],div[tabindex]")).map(c=>S(c.innerText));
            if (cells.some(v=>v!=="")) data3.push(cells);
          });
          if (data3.length) return data3;
        }
      }
      // 4) Last resort: dense divs under a scroller
      const scrollers = Array.from(doc.querySelectorAll("div,section")).filter(el=>{
        const st=getComputedStyle(el);
        return /(auto|scroll)/.test(st.overflow + st.overflowY + st.overflowX);
      });
      for (const sc of scrollers) {
        const rows2 = Array.from(sc.querySelectorAll("div[tabindex], [data-row-index]"));
        if (rows2.length > 5) {
          const data4=[];
          rows2.forEach(r=>{
            let cells = Array.from(r.children).map(c=>S(c.innerText)).filter(Boolean);
            if (!cells.length) cells = Array.from(r.querySelectorAll("div,span")).map(c=>S(c.innerText)).filter(Boolean);
            if (cells.length) data4.push(cells);
          });
          if (data4.length) return data4;
        }
      }
      return null;
    };

    // Collect from page and any same-origin iframes; choose the largest result
    const collect = () => {
      const docs = getDocs();
      let best=null;
      for (const d of docs) {
        const A = collectFromDoc(d);
        if (A && (!best || A.length > best.length)) best = A;
      }
      return best;
    };

    // Header fix modes
    const headerBlankDuplicates = (A) => {
      if (!A || !A.length) return A;
      const H = A[0].slice();
      // Auto-detect runs of identical labels and blank the duplicates in each run
      for (let i=0; i<H.length; ) {
        const cur = H[i];
        let j = i+1;
        while (j < H.length && H[j] === cur) j++;
        for (let k=i+1; k<j; k++) H[k] = ""; // only blank duplicates
        i = j;
      }
      A[0] = H;
      return A;
    };

    const headerCycleUnique = (A) => {
      if (!A || !A.length) return A;
      const H = A[0].map(S);
      // Build base list of unique headers from runs
      const base=[];
      for (let i=0; i<H.length; ) {
        const label = H[i];
        if (label!=="") base.push(label);
        let j=i+1;
        while (j<H.length && S(H[j])===S(label)) j++;
        i=j;
      }
      if (!base.length) return A;
      // Cycle base across full width so headers align with data columns (no data edits)
      const fixed = H.map((_, idx) => base[idx % base.length] || "");
      A[0] = fixed;
      return A;
    };

    // -------- run --------
    const A = collect();
    if (!A || !A.length) {
      alert("No table-like data found. Try scrolling the grid into view or opening the data view directly (not inside a cross-origin iframe).");
      return;
    }

    const mode = (prompt("Header fix mode:\n1 = Blank duplicates in header only (data untouched)\n2 = Cycle unique headers across all columns (data untouched)\nCancel = export as-is", "2") || "").trim();
    if (mode === "1") {
      headerBlankDuplicates(A);
    } else if (mode === "2") {
      headerCycleUnique(A);
    } // else leave as-is

    // Normalize to rectangular (pads shorter rows with blanks for Excel)
    const width = Math.max(...A.map(r=>r.length));
    for (let i=0;i<A.length;i++){
      if (A[i].length < width) A[i] = A[i].concat(Array(width - A[i].length).fill(""));
    }

    // Ask for filename
    const fname = (prompt("File name (without extension):", "pvalue") || "pvalue").replace(/[^\w.-]+/g,"_");
    const csv = toCSV(A);
    dl(new Blob([csv], {type:"text/csv;charset=utf-8;"}), fname + ".csv");
    alert("Done. Downloaded " + fname + ".csv (" + A.length + " rows).");

  } catch (err) {
    console.error(err);
    alert("Export error: " + (err && err.message ? err.message : err));
  }
})();
