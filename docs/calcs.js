
document.addEventListener("DOMContentLoaded", function() {
  /*************************************
   * TVPI, DPI + Single-Series IRR Script
   *************************************/

  // Query the buttons and display elements from the DOM
  let btn = document.querySelector("button.calc");
  let excelBtn = document.querySelector("button.excel");
  let calcHolder = document.getElementById("calc-view");
  let irrHolder = document.getElementById("irr");

  // Button listener for Compute Calcs
  btn.addEventListener("click", computeCalcs);

  // Button listener for Export to Excel
  excelBtn.addEventListener("click", () =>
    exportTableToExcel("calc-view", "PE_Valuation_Calculations")
  );

  let inputs = {
    commitedCapital: 600,
    periods: [
      "15/07/2003",
      "07/09/2003",
      "15/01/2004",
      "23/02/2004",
      "31/03/2004",
      "03/04/2004",
      "07/05/2004",
      "01/10/2004",
      "15/01/2005",
      "20/07/2005",
      "01/10/2005",
      "03/11/2005",
      "30/12/2005"
    ],
    paidInCap: [100, 27, 100, 0, 137.5, 0, 0, 37.5, 128, 0, 70, 0, 0],
    dist: [0, 0, 0, 50, 0, 45, 40, 0, 0, 40, 0, 25, 0],
    recalledCap: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    residualVal: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 590]
  };

  // --- "Raw" cash flows for full-series IRR computation ---
  const rawCashFlows = [
    { date: "2003-07-15", value: 100.0 },
    { date: "2003-09-07", value: 27.0 },
    { date: "2004-01-15", value: 100.0 },
    { date: "2004-02-23", value: -50.0 },
    { date: "2004-03-31", value: 137.5 },
    { date: "2004-04-03", value: -45.0 },
    { date: "2004-05-07", value: -40.0 },
    { date: "2004-10-01", value: 37.5 },
    { date: "2005-01-15", value: 128.0 },
    { date: "2005-07-20", value: -40.0 },
    { date: "2005-10-01", value: 70.0 },
    { date: "2005-11-03", value: -25.0 },
    { date: "2005-12-30", value: -547.0 }
  ];

  // --- Helper Functions ---
  function cumAsOfYear(flow, idx) {
    return flow.slice(0, idx + 1).reduce((sum, v) => sum + v, 0);
  }

  function dpi(paid, dist, rec, i) {
    return cumAsOfYear(dist, i) / (cumAsOfYear(paid, i) + cumAsOfYear(rec, i));
  }

  function rvpi(paid, dist, rec, resid, i) {
    return resid[i] / (cumAsOfYear(paid, i) + cumAsOfYear(rec, i));
  }

  function totVal(dist, rec, resid, i) {
    return cumAsOfYear(dist, i) - cumAsOfYear(rec, i) + resid[i];
  }

  function tvpi(paid, dist, rec, resid, i) {
    return (
      totVal(dist, rec, resid, i) / (cumAsOfYear(paid, i) + cumAsOfYear(rec, i))
    );
  }

  function moic(resid, rec, dist, comm, i) {
    return totVal(dist, rec, resid, i) / comm;
  }

  function limDigits(d, v) {
    return typeof v === "number" && isFinite(v) ? v.toFixed(d) : "";
  }

  // --- XNPV / XIRR Implementation ---
  function xnpv(rate, cfs) {
    const t0 = new Date(cfs[0].date).getTime();
    return cfs.reduce((sum, { date, value }) => {
      const dt = (new Date(date).getTime() - t0) / (1000 * 60 * 60 * 24 * 365);
      return sum + value / Math.pow(1 + rate, dt);
    }, 0);
  }

  function xirr(cfs, guess = 0.1) {
    let rate = guess;
    for (let i = 0; i < 100; i++) {
      const f0 = xnpv(rate, cfs);
      const dx = rate * 1e-4 || 1e-4;
      const f1 = xnpv(rate + dx, cfs);
      const deriv = (f1 - f0) / dx;
      if (Math.abs(deriv) < 1e-8) break;
      const next = rate - f0 / deriv;
      if (Math.abs(next - rate) < 1e-6) {
        rate = next;
        break;
      }
      rate = next;
    }
    return rate;
  }

  // --- Table Assembly & Rendering ---
  function getArrayToDisplay(headers, irrArr) {
    const rows = [];
    for (let i = 0; i < inputs.paidInCap.length; i++) {
      rows.push([
        inputs.periods[i],
        inputs.paidInCap[i],
        cumAsOfYear(inputs.paidInCap, i),
        inputs.dist[i],
        cumAsOfYear(inputs.dist, i),
        inputs.recalledCap[i],
        inputs.paidInCap[i] - inputs.dist[i] - inputs.residualVal[i],
        inputs.residualVal[i],
        totVal(inputs.dist, inputs.recalledCap, inputs.residualVal, i),
        dpi(inputs.paidInCap, inputs.dist, inputs.recalledCap, i),
        tvpi(
          inputs.paidInCap,
          inputs.dist,
          inputs.recalledCap,
          inputs.residualVal,
          i
        ),
        rvpi(
          inputs.paidInCap,
          inputs.dist,
          inputs.recalledCap,
          inputs.residualVal,
          i
        ),
        moic(
          inputs.residualVal,
          inputs.recalledCap,
          inputs.dist,
          inputs.commitedCapital,
          i
        ),
        irrArr[i]
      ]);
    }
    rows.unshift(headers);
    return rows[0].map((_, ci) => rows.map((r) => r[ci]));
  }

  function createTable(tableData, elementId) {
    const container = document.getElementById(elementId);
    container.innerHTML = "";

    const tbl = document.createElement("table");
    const tbody = document.createElement("tbody");

    const irrRowIndex = tableData.length - 1;

    tableData.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      row.forEach((cell, cellIndex) => {
        const td = document.createElement("td");

        // Check if this is the header row (first row after transpose)
        const isHeaderRow = rowIndex === 0;
        const isFirstColumn = cellIndex === 0;
        
        if (isHeaderRow || isFirstColumn) {
          // Header row or first column - just display as-is (labels)
          td.textContent = cell;
        } else if (rowIndex === irrRowIndex) {
          // IRR row data cells (not first column)
          if (typeof cell === "number" && isFinite(cell)) {
            // Format IRR as percentage
            td.textContent = (cell * 100).toFixed(2) + "%";
          } else {
            // Show dash for empty/NaN values
            td.textContent = "-";
          }
        } else if (typeof cell === "number") {
          td.textContent = limDigits(2, cell);
        } else {
          td.textContent = cell;
        }

        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    tbl.appendChild(tbody);
    container.appendChild(tbl);
  }

  // --- Main Calculation Function ---
  function computeCalcs() {
    const rowHeaders = [
      "Date",
      "Paid In Capital",
      "Cumulative Paid In Cap",
      "Distributions",
      "Cumulative Distributions",
      "Recalled Cap",
      "Contrib/Dist",
      "Residual Val",
      "Total Val",
      "DPI",
      "TVPI",
      "RVPI",
      "MOIC",
      "NET IRR"
    ];

    // compute single full-series IRR
    const fullIrr = xirr(rawCashFlows);

    // build IRR array with value only at last index
    const irrValues = inputs.periods.map((_, idx) =>
      idx === inputs.periods.length - 1 ? fullIrr : NaN
    );

    // assemble + render
    const arr = getArrayToDisplay(rowHeaders, irrValues);
    createTable(arr, "calc-view");

    // optional separate IRR display
    if (irrHolder) {
      irrHolder.innerText = `Overall IRR: ${limDigits(2, fullIrr)}`;
    }

    btn.disabled = true;
    console.log("Rendered with full-series IRR:", fullIrr);
  }

  // --- Export to Excel Functionality ---
  function exportTableToExcel(containerID, filename = "") {
    let downloadLink;
    const dataType = "application/vnd.ms-excel";
    const tableContainer = document.getElementById(containerID);
    const tableSelect = tableContainer.querySelector("table");
    if (!tableSelect) {
      alert("No table found to export!");
      return;
    }
    let tableHTML = tableSelect.outerHTML.replace(/ /g, "%20");
    filename = filename ? filename + ".xls" : "export_data.xls";
    downloadLink = document.createElement("a");
    if (navigator.msSaveOrOpenBlob) {
      const blob = new Blob(["\ufeff", tableHTML], { type: dataType });
      navigator.msSaveOrOpenBlob(blob, filename);
    } else {
      downloadLink.href = "data:" + dataType + ", " + tableHTML;
      downloadLink.download = filename;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);
    }
  }
});
