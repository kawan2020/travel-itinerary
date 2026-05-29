document.addEventListener("DOMContentLoaded", () => {
    const excelUrl = './Data/itinerary.xlsx';

    fetch(excelUrl)
        .then(response => {
            if (!response.ok) throw new Error("Excel file not found");
            return response.arrayBuffer();
        })
        .then(buffer => {
            const data = new Uint8Array(buffer);
            const workbook = XLSX.read(data, { type: 'array', cellDates: false });
            
            parseSummaryTab(workbook);
            
            // Transportation now reads 9 columns (A to I)
            parseHorizontalTab(workbook, 'Transportation', 'table-transportation', 9);
            
            // Accommodation stays standard at 6 columns (A to F)
            parseHorizontalTab(workbook, 'Accommodation', 'table-accommodation', 6);
            
            // Activities now reads 7 columns (A to G)
            parseHorizontalTab(workbook, 'Activities', 'table-activities', 7);
            
            // Contact stays standard at 7 columns (A to G)
            parseHorizontalTab(workbook, 'Contact', 'table-contact', 7);
            
            parseOtherInfoTab(workbook);
        })
        .catch(err => {
            console.error("Error:", err);
        });
});

function parseSummaryTab(workbook) {
    const sheet = workbook.Sheets['Summary'];
    if (!sheet) return;
    document.getElementById('summary-trip-type').innerText = sheet['B3']?.w || sheet['B3']?.v || '-';
    document.getElementById('summary-dest').innerText = sheet['B4']?.w || sheet['B4']?.v || '-';
    document.getElementById('summary-start').innerText = sheet['B5']?.w || sheet['B5']?.v || '-';
    document.getElementById('summary-end').innerText = sheet['B6']?.w || sheet['B6']?.v || '-';
}

function parseHorizontalTab(workbook, sheetName, tableId, totalColumns) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;

    const tbody = document.querySelector(`#${tableId} tbody`);
    tbody.innerHTML = "";

    let rowIndex = 4; 
    while (true) {
        let primaryKey = `A${rowIndex}`;
        if (!sheet[primaryKey] || sheet[primaryKey].v === undefined) break; 

        let rowHtml = "<tr>";
        
        // 🔍 DYNAMIC HIDDEN ADDRESS CAPTURE
        let hiddenAddressVal = '';
        if (sheetName === 'Transportation') {
            hiddenAddressVal = sheet[`I${rowIndex}`]?.w || sheet[`I${rowIndex}`]?.v || '';
        } else if (sheetName === 'Activities') {
            hiddenAddressVal = sheet[`G${rowIndex}`]?.w || sheet[`G${rowIndex}`]?.v || '';
        }

        for (let colIndex = 0; colIndex < totalColumns; colIndex++) {
            let colLetter = XLSX.utils.encode_col(colIndex);
            let cell = sheet[`${colLetter}${rowIndex}`];
            let val = cell?.w || cell?.v || '-';

            // Skip rendering the hidden address columns onto the homepage table layout
            if (sheetName === 'Transportation' && colIndex === 8) continue; // Skip Column I
            if (sheetName === 'Activities' && colIndex === 6) continue;     // Skip Column G

            // Generate clean button parameters matching the respective sheets
            if (colIndex === (totalColumns - 1) || (sheetName === 'Transportation' && colIndex === 7)) {
                if (val && val !== '-') {
                    rowHtml += `<td><a href="info.html?id=${encodeURIComponent(val)}&address=${encodeURIComponent(hiddenAddressVal)}" class="info-btn">More Info</a></td>`;
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
    if (!sheet) return;

    let contentContainer = document.getElementById('other-info-content');
    contentContainer.innerHTML = "";

    let rowIndex = 4;
    while (true) {
        let cellRef = `A${rowIndex}`;
        if (!sheet[cellRef] || sheet[cellRef].v === undefined) break;

        let lineText = sheet[cellRef].w || sheet[cellRef].v || '';
        contentContainer.innerHTML += `<div>${lineText}</div>`;
        rowIndex++;
    }
}
