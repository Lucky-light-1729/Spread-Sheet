let sheetCount = 0;
let cellData = {};
let currentCell = null;

function createSpreadSheet(containerId) 
{
    const container = document.getElementById(containerId);
    const spreadSheet = document.createElement('div');
    spreadSheet.className = 'spreadSheet';
    spreadSheet.dataset.sheetId = sheetCount;

    spreadSheet.appendChild(document.createElement('div'));

    for (let col = 0; col < 20; col++) 
    {
        const colHeader = document.createElement('div');
        colHeader.className = 'header';
        colHeader.textContent = String.fromCharCode(65 + col);
        spreadSheet.appendChild(colHeader);
    }

    for (let row = 1; row < 25; row++) 
    {
        const rowHeader = document.createElement('div');
        rowHeader.className = 'row-header';
        rowHeader.textContent = row;
        spreadSheet.appendChild(rowHeader);

        for (let col = 0; col < 20; col++) 
        {
            const cell = document.createElement('div');
            cell.className = 'cell';
            cell.contentEditable = true;
            const cellId = `S${sheetCount}_${String.fromCharCode(65 + col)}${row}`;
            cell.dataset.id = cellId;

            cell.addEventListener('click', () => currentCell = cell);
            cell.addEventListener('input', () => {
                cellData[cellId] = {
                    value: cell.innerText,
                    color: cell.style.color,
                    bg: cell.style.backgroundColor
                };
            });
            spreadSheet.appendChild(cell);
        }
    }

    container.appendChild(spreadSheet);
}

function format(command) 
{
    document.execCommand(command, false, null);
}

function applyColor(color, type) 
{
    if (currentCell) 
    {
        currentCell.style[type] = color;
        const cellId = currentCell.dataset.id;
        cellData[cellId] = cellData[cellId] || {};
        cellData[cellId][type === 'color' ? 'color' : 'bg'] = color;
    }
}

document.getElementById('font-select').addEventListener('change', function () {
    document.execCommand('fontName', false, this.value);
});

function addNewSheet() 
{
    createSpreadSheet('sheet-container');
    sheetCount++;
}

function exportToCSV() 
{
    let csv = '';
    const rows = 25;
    const cols = 20;

    for (let row = 1; row <= rows; row++) 
    {
        let rowData = [];
        for (let col = 0; col < cols; col++) 
        {
            let cellId = `S${sheetCount - 1}_${String.fromCharCode(65 + col)}${row}`;
            rowData.push((cellData[cellId]?.value || '').replace(/,/g, ''));
        }
        csv += rowData.join(',') + '\n';
    }

    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'spreadsheet.csv';
    a.click();
    URL.revokeObjectURL(url);
}

window.onload = () => {
    addNewSheet();
};
