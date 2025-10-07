 document.addEventListener("contextmenu", e => e.preventDefault());
    document.addEventListener("keydown", e => {
      if (123 == e.keyCode || (e.ctrlKey && e.shiftKey && [73,74].includes(e.keyCode)) || (e.ctrlKey && [85,83].includes(e.keyCode))) {
        e.preventDefault(); return false;
      }
    });

    const calcBtn = document.querySelector("button.calc");
    const excelBtn = document.querySelector("button.excel");
    const calcView = document.getElementById("calc-view");

    // Excel Export FIXED FUNCTION
    function exportToExcel(elementId, filename = "export_data.xls") {
      const container = document.getElementById(elementId);
      const table = container.querySelector("table");
      if (!table) return alert("No table found to export!");

      const html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" 
              xmlns:x="urn:schemas-microsoft-com:office:excel" 
              xmlns="http://www.w3.org/TR/REC-html40">
        <head><meta charset="UTF-8"></head><body>${table.outerHTML}</body></html>
      `;

      const blob = new Blob([html], { type: "application/vnd.ms-excel" });
      const url = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }

    excelBtn.addEventListener("click", () => exportToExcel("calc-view", "PE_Valuation_Calculations.xls"));

    const n = {
      x:600,
      y:["15/07/2003","07/09/2003","15/01/2004","23/02/2004","31/03/2004","03/04/2004","07/05/2004","01/10/2004","15/01/2005","20/07/2005","01/10/2005","03/11/2005","30/12/2005"],
      z:[100,27,100,0,137.5,0,0,37.5,128,0,70,0,0],
      A:[0,0,0,50,0,45,40,0,0,40,0,25,0],
      B:[0,0,0,0,0,0,0,0,0,0,0,0,0],
      C:[0,0,0,0,0,0,0,0,0,0,0,0,590]
    };

    const l = [
      {date:"2003-07-15",value:100},{date:"2003-09-07",value:27},
      {date:"2004-01-15",value:100},{date:"2004-02-23",value:-50},
      {date:"2004-03-31",value:137.5},{date:"2004-04-03",value:-45},
      {date:"2004-05-07",value:-40},{date:"2004-10-01",value:37.5},
      {date:"2005-01-15",value:128},{date:"2005-07-20",value:-40},
      {date:"2005-10-01",value:70},{date:"2005-11-03",value:-25},
      {date:"2005-12-30",value:-547}
    ];

    function c(e,t){return e.slice(0,t+1).reduce((a,b)=>a+b,0);}
    function r(e,t,a,n){return c(t,n)/(c(e,n)+c(a,n));}
    function d(e,t,a,n,l){return n[l]/(c(e,l)+c(a,l));}
    function o(e,t,a,n){return c(e,n)-c(t,n)+a[n];}
    function i(e,t,a,n,l){return o(t,a,n,l)/(c(e,l)+c(a,l));}
    function s(e,t,a,n,l){return o(a,t,e,l)/n;}
    function u(e,t){return typeof t==="number"&&isFinite(t)?t.toFixed(e):"";}
    function h(e,t){
      const a=(new Date(t[0].date)).getTime();
      return t.reduce((sum,{date:n,value:l})=>{
        const c=((new Date(n)).getTime()-a)/315576e5;
        return sum+l/Math.pow(1+e,c);
      },0);
    }
    function f(e,t=.1){
      let a=t;
      for(let i=0;i<100;i++){
        const val=h(a,e), n=1e-4*a||1e-4, l=h(a+n,e), slope=(l-val)/n;
        if(Math.abs(slope)<1e-8)break;
        const newA=a-val/slope;
        if(Math.abs(newA-a)<1e-6){a=newA;break;}
        a=newA;
      }
      return a;
    }

    function m(e,t){
      const a=[];
      for(let e=0;e<n.z.length;e++)
        a.push([
          n.y[e],n.z[e],c(n.z,e),n.A[e],c(n.A,e),n.B[e],
          n.z[e]-n.A[e]-n.C[e],n.C[e],
          o(n.A,n.B,n.C,e),
          r(n.z,n.A,n.B,e),
          i(n.z,n.A,n.B,n.C,e),
          d(n.z,n.A,n.B,n.C,e),
          s(n.C,n.B,n.A,n.x,e),
          t[e]
        ]);
      return a.unshift(e),a[0].map((_,col)=>a.map(row=>row[col]));
    }

    function v(e,t){
      const a=document.getElementById(t);
      a.innerHTML="";
      const n=document.createElement("table");
      const l=document.createElement("tbody");
      const c=e.length-1;
      e.forEach((row,rowIndex)=>{
        const tr=document.createElement("tr");
        row.forEach((val,colIndex)=>{
          const td=document.createElement("td");
          const header=rowIndex===0;
          if(header || colIndex===0){td.textContent=val;}
          else if(rowIndex===c){
            td.textContent=(typeof val==="number"&&isFinite(val)?(100*val).toFixed(2)+"%":"-");
          }else{
            td.textContent=typeof val==="number"?u(2,val):val;
          }
          tr.appendChild(td);
        });
        l.appendChild(tr);
      });
      n.appendChild(l);
      a.appendChild(n);
    }

    function g(){
      const headers=["Date","Paid In Capital","Cumulative Paid In Cap","Distributions","Cumulative Distributions","Recalled Cap","Contrib/Dist","Residual Val","Total Val","DPI","TVPI","RVPI","MOIC","NET IRR"];
      const irr=f(l);
      const irrArray=n.y.map((_,i)=>i===n.y.length-1?irr:NaN);
      const matrix=m(headers,irrArray);
      v(matrix,"calc-view");
      calcBtn.disabled=true;
    }

    calcBtn.addEventListener("click", g);
