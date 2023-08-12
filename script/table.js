// document.addEventListener('DOMContentLoaded', () => {
//   const fileInput = document.createElement('input');
//   fileInput.type = 'file';
//   fileInput.accept = '.xlsx';
//   fileInput.addEventListener('change', handleFile, false);
//   document.body.appendChild(fileInput);
// });

// function handleFile(e) {
//   const files = e.target.files;
//   const file = files[0];

//   const reader = new FileReader();
//   reader.onload = function (e) {
//     const data = new Uint8Array(e.target.result);
//     const workbook = XLSX.read(data, { type: 'array' });
//     const sheetName = workbook.SheetNames[0];
//     const worksheet = workbook.Sheets[sheetName];
//     const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//     const tableBody = document.querySelector('#data-table tbody');
//     tableBody.innerHTML = '';

//     for (let i = 1; i < jsonData.length; i++) {
//       const row = jsonData[i];
//       const tr = document.createElement('tr');

//       for (let j = 0; j < row.length; j++) {
//         const td = document.createElement('td');
//         td.textContent = row[j];
//         tr.appendChild(td);
//       }

//       tableBody.appendChild(tr);
//     }
//   };

//   reader.readAsArrayBuffer(file);
// }

// //////////////////////////////////////////////////////////////////////////////////

// document.addEventListener("DOMContentLoaded", () => {
//   loadExcelData("../tables/Indicators-TC-Belgrade.xlsx");
// });

// function loadExcelData(filePath) {
//   const xhr = new XMLHttpRequest();
//   xhr.open("GET", filePath, true);
//   xhr.responseType = "arraybuffer";

//   xhr.onload = function (e) {
//     const data = new Uint8Array(xhr.response);
//     const workbook = XLSX.read(data, { type: "array" });
//     const sheetName = workbook.SheetNames[0];
//     const worksheet = workbook.Sheets[sheetName];
//     const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//     const tableBody = document.querySelector("#data-table tbody");
//     tableBody.innerHTML = "";

//     for (let i = 1; i < jsonData.length; i++) {
//       const row = jsonData[i];
//       const tr = document.createElement("tr");

//       for (let j = 0; j < row.length; j++) {
//         const td = document.createElement("td");
//         td.textContent = row[j];
//         tr.appendChild(td);
//       }

//       tableBody.appendChild(tr);
//     }
//   };

//   xhr.send();
// }

// function openPage(pageUrl) {
//   window.open(pageUrl);
// }

/////////////////////////////////////////////////
document.addEventListener("DOMContentLoaded", () => {
  loadExcelData("../tables/Indicators-TC-Belgrade.xlsx");
});

function loadExcelData(filePath) {
  const xhr = new XMLHttpRequest();
  xhr.open("GET", filePath, true);
  xhr.responseType = "arraybuffer";

  xhr.onload = function (e) {
    const data = new Uint8Array(xhr.response);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const tableBody = document.querySelector("#data-table tbody");
    tableBody.innerHTML = "";
    let maxRow = 0;
    for (let i = 2; i < jsonData.length; i++) {
      const row = jsonData[i];
      maxRow = Math.max(maxRow, row.length);
      const tr = document.createElement("tr");
      for (let j = 0; j < maxRow; j++) {
        const td = document.createElement("td");

        // Check if the cell is part of a merged cell range
        const cellAddress = XLSX.utils.encode_cell({ r: i, c: j });
        const mergedCell =
          worksheet["!merges"] &&
          worksheet["!merges"].find(
            (merge) =>
              merge.s.c <= j &&
              merge.e.c >= j &&
              merge.s.r <= i &&
              merge.e.r >= i
          );

        if (mergedCell) {
          // Get the value from the top-left cell of the merged range
          const topLeftCell = XLSX.utils.decode_cell(mergedCell.s);
          td.textContent = getMergedCellValue(
            topLeftCell.r,
            topLeftCell.c,
            worksheet,
            jsonData
          );
        } else {
          td.textContent = row[j];
        }

        tr.appendChild(td);
      }

      tableBody.appendChild(tr);
    }
  };

  xhr.send();
}

function getMergedCellValue(row, col, worksheet, jsonData) {
  for (const merge of worksheet["!merges"]) {
    const { s, e } = merge;
    if (row >= s.r && row <= e.r && col >= s.c && col <= e.c) {
      const topLeftCell = XLSX.utils.decode_cell(s);
      return jsonData[topLeftCell.r][topLeftCell.c];
    }
  }
  return null;
}

function openPage(pageUrl) {
  window.open(pageUrl);
}
