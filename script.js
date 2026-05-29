document.addEventListener("DOMContentLoaded", () => {
    // Automatically calculate the correct path on GitHub Pages
    const excelUrl = window.location.pathname.endsWith('/') 
        ? 'Data/itinerary.xlsx' 
        : './Data/itinerary.xlsx';

    fetch(excelUrl)
        .then(response => {
            if (!response.ok) throw new Error("Excel file not found at " + excelUrl);
            return response.arrayBuffer();
        })
        .then(buffer => {
            const data = new Uint8Array(buffer);
            // Read file using the locally hosted library
            const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: true });
            
            parseSummaryTab(workbook);
            parseHorizontalTab(workbook, 'Transportation', 'table-transportation', 8);
            parseHorizontalTab(workbook, 'Accommodation', 'table-accommodation', 6);
            parseHorizontalTab(workbook, 'Activities', 'table-activities', 6);
            parseHorizontalTab(workbook, 'Contact', 'table-contact', 7);
            parseOtherInfoTab(workbook);
        })
        .catch(err => {
            console.error("Initialization Error:", err);
            document.body.insertAdjacentHTML('afterbegin', `<div style="background:#ffdddd; color:#990000; padding:15px; text-align:center; font-weight:bold; margin-bottom:10px;">Loading Failed: ${err.message}</div>`);
        });
});

function parseSummaryTab(workbook) {
    const sheet = workbook.Sheets['Summary'];
    if (!sheet) return;
    const getVal = (cellName) => (sheet[cellName] && sheet[cellName].v !== undefined) ? sheet[cellName].v : '-';
    
    document.getElementById('summary-trip-type').innerText = getVal('B3');
    document.getElementById('summary-dest').innerText = getVal('B4');
    document.getElementById('summary-start').innerText = formatCellText(sheet['B5']);
    document.getElementById('summary-end').innerText = formatCellText(sheet['B6']);
}

function parseHorizontalTab(workbook, sheetName, tableId, totalColumns) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;

    const tbody = document.querySelector(`#${tableId} tbody`);
    tbody.innerHTML = "";

    let rowIndex = 4; 
    while (true) {
        let primaryKey = `A${rowIndex}`;
        if (!sheet[primaryKey] || sheet[primaryKey].v === undefined || String(sheet[primaryKey].v).trim() === '') break; 

        let rowHtml = "<tr>";
        for (let colIndex = 0; colIndex < totalColumns; colIndex++) {
            let colLetter = XLSX.utils.encode_col(colIndex);
            let cell = sheet[`${colLetter}${rowIndex}`];
            let val = formatCellText(cell);

            if (colIndex === (totalColumns - 1)) {
                if (val && val !== '-') {
                    rowHtml += `<td><a href="info.html?id=${encodeURIComponent(val)}" class="info-btn">More Info</a></td>`;
                } else {
                    rowHtml += `<td>-</td>`;
                }
            } else {
                rowHtml += `<td>${val}</td>`;
            }
        }
        rowHtml += "</tr>";
        tbody.innerHTML += rowHtml;
        rowIndex++;
    }
}

function parseOtherInfoTab(workbook) {
    const sheet = workbook.Sheets['Other Info'];
    if (sheet && sheet['A4'] && sheet['A4'].v !== undefined) {
        document.getElementById('other-info-content').innerText = sheet['A4'].v;
    }
}

function formatCellText(cell) {
    if (!cell || cell.v === undefined || cell.v === '') return '-';
    if (cell.v instanceof Date) {
        try {
            return cell.v.toISOString().split('T')[0];
        } catch(e) {
            return cell.w ? cell.w : String(cell.v);
        }
    }
    return cell.w ? cell.w : cell.v;
}
