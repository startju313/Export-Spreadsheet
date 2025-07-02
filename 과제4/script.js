const sheetContainer = document.querySelector("#spreadsheet-container");
const exportButton = document.querySelector("#export-btn");
const numRows = 10; 
const numCols = 10; 
const sheetData = []; 
const colLabels = [ 
    "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
];

class Cell {
    constructor(isHeader, isDisabled, cellVal, rowIdx, colIdx, dispRowName, dispColName, active = false) {
        this.isHeader = isHeader;
        this.isDisabled = isDisabled;
        this.cellVal = cellVal;
        this.rowIdx = rowIdx;
        this.colIdx = colIdx;
        this.dispRowName = dispRowName;
        this.dispColName = dispColName;
        this.active = active;
    }
}

exportButton.addEventListener('click', function (e) { 
    let csv = "";
    for (let i = 0; i < sheetData.length; i++) { 
        if (i === 0) continue; // 헤더 행 건너뛰기
        csv +=
            sheetData[i] 
                .filter(item => !item.isHeader)
                .map(item => item.cellVal) 
                .join(',') + "\r\n";
    }

    const csvObj = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const csvUrl = URL.createObjectURL(csvObj);

    const a = document.createElement("a");
    a.href = csvUrl;
    a.download = 'spreadsheet.csv'; 
    a.click();
    URL.revokeObjectURL(csvUrl);
});

initSpreadsheet();

function initSpreadsheet() {
    for (let i = 0; i < numRows; i++) { 
        let spreadsheetRow = [];
        for (let j = 0; j < numCols; j++) { 
            let initialCellValue = ''; 
            let isHeaderCell = false;
            let disableCell = false; 

            // 첫 번째 열에 행 번호 넣기
            if (j === 0) {
                initialCellValue = i;
                isHeaderCell = true;
                disableCell = true;
            }

            // 첫 번째 행에 알파벳 컬럼명 넣기
            if (i === 0) {
                // (0,0) 셀은 비워두거나 원하는 값으로 설정
                initialCellValue = colLabels[j - 1]; 
                isHeaderCell = true;
                disableCell = true;
            }

            // 첫 번째 row의 첫 번째 컬럼은 비움
            if (i === 0 && j === 0) {
                initialCellValue = "";
            }

            const currentRowName = i; 
            const currentColName = colLabels[j - 1]; 

            const cell = new Cell(isHeaderCell, disableCell, initialCellValue, i, j, currentRowName, currentColName, false);
            spreadsheetRow.push(cell);
        }
        sheetData.push(spreadsheetRow); 
    }
    drawSheet();
}


function createCellElement(cell) { 
    const cellElem = document.createElement('input'); 
    cellElem.className = 'cell';
    cellElem.id = 'cell_' + cell.rowIdx + cell.colIdx; 
    cellElem.value = cell.cellVal; 
    cellElem.disabled = cell.isDisabled; 

    if (cell.isHeader) {
        cellElem.classList.add("header");
    }

    cellElem.onclick = () => onCellClick(cell); 
    cellElem.onchange = (e) => updateCellValue(e.target.value, cell); 
    return cellElem;
}

function updateCellValue(value, cell) { 
    cell.cellVal = value; 
}

function onCellClick(cell) { 
    clearHeaderActiveStates();
    
    const clickedCellElement = getElementFromRowCol(cell.rowIdx, cell.colIdx);
    if (clickedCellElement) {
        clickedCellElement.classList.add('selected');
    }

    const columnHeader = sheetData[0][cell.colIdx]; 
    const rowHeader = sheetData[cell.rowIdx][0]; 
    const columnHeaderEl = getElementFromRowCol(columnHeader.rowIdx, columnHeader.colIdx); 
    const rowHeaderEl = getElementFromRowCol(rowHeader.rowIdx, rowHeader.colIdx); 
    
    columnHeaderEl.classList.add('active');
    rowHeaderEl.classList.add('active');
    
    document.querySelector("#cell-status").innerHTML = cell.dispColName + cell.dispRowName; 
}

function clearHeaderActiveStates() {
    const headers = document.querySelectorAll('.header');
    headers.forEach((header) => {
        header.classList.remove('active');
    });
    
    const selectedCells = document.querySelectorAll('.cell.selected');
    selectedCells.forEach((selectedCell) => {
        selectedCell.classList.remove('selected');
    });
}

function getElementFromRowCol(row, col) { 
    return document.querySelector("#cell_" + row + col);
}

function drawSheet() {
    for (let i = 0; i < sheetData.length; i++) { 
        const rowContainerEl = document.createElement("div");
        rowContainerEl.className = "cell-row";

        for (let j = 0; j < sheetData[i].length; j++) { 
            const cell = sheetData[i][j]; 
            rowContainerEl.append(createCellElement(cell)); 
        }
        sheetContainer.append(rowContainerEl); 
    }
}