document.addEventListener("DOMContentLoaded", () => {
  loadExcelData("../tables/humanCapital.xlsx");
});

function loadExcelData(filePath) {
  const xhr = new XMLHttpRequest();
  xhr.open("GET", filePath, true);
  xhr.responseType = "arraybuffer";

  xhr.onload = function (e) {
    const data = new Uint8Array(xhr.response);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    console.log(workbook.SheetNames);
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    console.log(jsonData);
    const tableBody = document.querySelector("#data-table tbody");
    tableBody.innerHTML = "";
    let maxRow = 1;

const mergedCellAll = 
  worksheet["!merges"];
    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      maxRow = Math.max(maxRow, row.length);
      const tr = document.createElement("tr");
      for (let j = 0; j < maxRow; j++) {

        const td = document.createElement("td");

        if(mergedCellAll != undefined){
          mergedCellAll.forEach(element => {
            if(i == element['s']['r'] && j == element['s']['c']) {
              td.rowSpan = element['e']['r'] - element['s']['r'] +1;
          }
            td.textContent = row[j];
            
          });
          if ((row[j] === undefined)) {
            if (j != mergedCellAll[0]['s']['c'] && j!= mergedCellAll[1]['s']['c'] && j != mergedCellAll[2]['s']['c'] && j!= mergedCellAll[3]['s']['c']) {
              tr.appendChild(td);
            }
          } else {
            tr.appendChild(td)
          }
        }
        else {
          td.textContent = row[j];
          tr.appendChild(td)
        }
      }

      tableBody.appendChild(tr);
    }
  };

  xhr.send();
}

function openPage(pageUrl) {
  window.open(pageUrl);
}
