document.addEventListener("DOMContentLoaded", function () {
  let consolidatedData = {};

  const dropZone = document.getElementById("dropZone");
  const fileInput = document.getElementById("fileUpload");

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    fileInput.files = e.dataTransfer.files;
    handleFiles();
  });

  dropZone.addEventListener("click", () => fileInput.click());

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

    const barCanvas = document.getElementById("salesChart");
    const pieCanvas = document.getElementById("pieChart");

    const dpr = window.devicePixelRatio || 2;
    const displayWidth = barCanvas.parentElement.offsetWidth || 2000;
    const rowHeight = 70;
    const displayHeight = Math.max(400, labels.length * rowHeight);

    barCanvas.style.width = `${displayWidth}px`;
    barCanvas.style.height = `${displayHeight}px`;
    barCanvas.width = displayWidth * dpr;
    barCanvas.height = displayHeight * dpr;

    const barCtx = barCanvas.getContext("2d");
    barCtx.scale(dpr, dpr);

    barCanvas.style.display = "block";

    new Chart(barCtx, {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Top 10 SKU Sales",
            data,
            backgroundColor: "#007bff",
            borderColor: "#0056b3",
            hoverBackgroundColor: "#3399ff",
          },
        ],
      },
      options: {
        responsive: false,
        maintainAspectRatio: false,
        devicePixelRatio: dpr,
        plugins: {
          legend: {
            labels: {
              color: "#000",
              font: { size: 14 },
            },
          },
          tooltip: {
            backgroundColor: "#f0f0f0",
            titleColor: "#000",
            bodyColor: "#000",
            borderColor: "#ccc",
            borderWidth: 1,
          },
        },
        scales: {
          x: {
            ticks: { color: "#000" },
            grid: { color: "rgba(0, 0, 0, 0.1)" },
          },
          y: {
            ticks: { color: "#000" },
            grid: { color: "rgba(0, 0, 0, 0.1)" },
          },
        },
      },
    });

    pieCanvas.style.width = `${displayWidth}px`;
    pieCanvas.style.height = `${displayHeight}px`;
    pieCanvas.width = displayWidth * dpr;
    pieCanvas.height = displayHeight * dpr;

    const pieCtx = pieCanvas.getContext("2d", { willReadFrequently: true });

    new Chart(pieCtx, {
      type: "pie",
      data: {
        labels,
        datasets: [
          {
            label: "Sales Share",
            data,
            backgroundColor: [
              "#00c6ff",
              "#0072ff",
              "#00ffcc",
              "#ffaa00",
              "#ff6384",
              "#36a2eb",
              "#cc65fe",
              "#ffce56",
              "#4bc0c0",
              "#9966ff",
            ],
            borderColor: "#ffffff",
            borderWidth: 1,
          },
        ],
      },
      options: {
        responsive: false,
        maintainAspectRatio: false,
        devicePixelRatio: dpr,
        layout: {
          padding: { top: 20, bottom: 20 },
        },
        plugins: {
          legend: {
            labels: {
              color: "#000000",
              font: { size: 14 },
            },
          },
          tooltip: {
            backgroundColor: "#f0f0f0",
            titleColor: "#000000",
            bodyColor: "#000000",
            borderColor: "#cccccc",
            borderWidth: 1,
          },
        },
      },
    });

    document.getElementById("salesChart").style.display = "block";
    document.getElementById("pieChart").style.display = "block";
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

  function downloadCanvasAsPDF(canvasId, filename) {
    const canvas = document.getElementById(canvasId);

    html2canvas(canvas, { scale: 2 }).then((canvasImage) => {
      const imgData = canvasImage.toDataURL("image/png");
      const pdf = new jspdf.jsPDF({
        orientation: "landscape",
        unit: "px",
        format: [canvasImage.width, canvasImage.height],
      });

      pdf.addImage(imgData, "PNG", 0, 0, canvasImage.width, canvasImage.height);
      pdf.save(filename);
    });
  }

  function downloadCanvasAsPNG(canvasId, filename) {
    const canvas = document.getElementById(canvasId);
    const link = document.createElement("a");
    link.download = filename;
    link.href = canvas.toDataURL("image/png");
    link.click();
  }

  document.getElementById("downloadBar").onclick = function () {
    downloadCanvasAsPNG("salesChart", "Top_10_Bar_Chart.png");
  };

  document.getElementById("downloadPie").onclick = function () {
    downloadCanvasAsPNG("pieChart", "Top_10_Pie_Chart.png");
  };

  window.handleFiles = handleFiles;
});
