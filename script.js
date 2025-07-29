document.addEventListener("DOMContentLoaded", function () {
  let consolidatedData = {};

  function handleFiles() {
    const input = document.getElementById("fileUpload");
    const files = input.files;
    if (!files.length) return alert("Please upload one or more Excel files.");

    consolidatedData = {};
    let filesRead = 0;

    for (let file of files) {
      const reader = new FileReader();
      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        extractSalesData(json, file.name);
        filesRead++;

        if (filesRead === files.length) {
          displayTable();
          createChart();
          document.getElementById("downloadBtn").style.display = "inline";
        }
      };
      reader.readAsArrayBuffer(file);
    }
  }

  function extractSalesData(data, filename) {
    const month = filename.match(/POS_(\w+)25/i)?.[1] || filename;
    for (let row of data) {
      for (let i = 0; i < row.length - 1; i += 2) {
        let sku = row[i];
        let sales = parseFloat(row[i + 1]);
        if (sku && !isNaN(sales)) {
          sku = sku.toString().trim();
          if (!consolidatedData[sku]) consolidatedData[sku] = 0;
          consolidatedData[sku] += sales;
        }
      }
    }
  }

  function displayTable() {
    const tableContainer = document.getElementById("tableContainer");
    const sortedEntries = Object.entries(consolidatedData).sort(
      (a, b) => b[1] - a[1]
    );

    let html =
      "<table><thead><tr><th>SKU</th><th>Sales</th></tr></thead><tbody>";
    for (let [sku, sales] of sortedEntries) {
      html += `<tr><td>${sku}</td><td>${sales}</td></tr>`;
    }
    html += "</tbody></table>";
    tableContainer.innerHTML = html;
  }

  function createChart() {
    const sorted = Object.entries(consolidatedData)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);
    const labels = sorted.map((e) => e[0]);
    const data = sorted.map((e) => e[1]);
    const ctx = document.getElementById("salesChart").getContext("2d");
    document.getElementById("salesChart").style.display = "block";

    const gradient = ctx.createLinearGradient(0, 0, 0, 400);
    gradient.addColorStop(0, "#00f0ff");
    gradient.addColorStop(1, "#00bfff");

    Chart.defaults.color = "#ffffff"; // sets global font colour
    Chart.defaults.font.family = "Inter, sans-serif"; // optional styling

    new Chart(ctx, {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Top 10 SKU Sales",
            data,
            backgroundColor: gradient,
            borderRadius: 8,
            borderWidth: 2,
            borderColor: "#00ffff",
            hoverBackgroundColor: "#00ffff",
          },
        ],
      },
      options: {
        responsive: true,
        plugins: {
          legend: {
            labels: {
              color: "#ffffff",
              font: { size: 14 },
            },
          },
          tooltip: {
            backgroundColor: "#0a0f1a",
            titleColor: "#00ffff",
            bodyColor: "#ffffff",
            borderColor: "#00ffff",
            borderWidth: 1,
            titleFont: { size: 14, weight: "bold" },
            bodyFont: { size: 13 },
          },
        },
        scales: {
          x: {
            ticks: { color: "#ffffff" },
            grid: { color: "rgba(255, 255, 255, 0.1)" },
          },
          y: {
            ticks: { color: "#ffffff" },
            grid: { color: "rgba(255, 255, 255, 0.1)" },
          },
        },
      },
    });
  }

  document.getElementById("downloadBtn").onclick = function () {
    const worksheet = XLSX.utils.aoa_to_sheet([
      ["SKU", "Sales"],
      ...Object.entries(consolidatedData),
    ]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Consolidated");
    XLSX.writeFile(workbook, "Consolidated_SKU_Sales.xlsx");
  };

  window.handleFiles = handleFiles; // expose to HTML button onclick
});
