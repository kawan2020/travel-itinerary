document.addEventListener("DOMContentLoaded", () => {
    // 1. Fetch your modified horizontal Excel file from GitHub
    fetch('Data/itinerary.xlsx')
        .then(response => {
            if (!response.ok) throw new Error("Excel file not found in Data folder");
            return response.arrayBuffer();
        })
        .then(buffer => {
            const data = new Uint8Array(buffer);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 2. Parse the respective sheets with accurate cell coordinates
            parseSummaryTab(workbook);
            parseHorizontalTab(workbook, 'Transportation', 'table-transportation', 8);
            parseHorizontalTab(workbook, 'Accommodation', 'table-accommodation', 6);
            parseHorizontalTab(workbook, 'Activities', 'table-activities', 6);
            parseHorizontalTab(workbook, 'Contact', 'table-contact', 7);
            parseOtherInfoTab(workbook);
        })
        .catch(err => console.error("Error updating web app itinerary:", err));
});

// Process the metadata summary fields based on Rows 3 to 6
function parseSummaryTab(workbook) {
    const sheet = workbook.Sheets['Summary'];
    if (!sheet) return;
    // Remapped to point precisely at your new row layout coordinates
    document.getElementById('summary-trip-type').innerText = sheet['B3'] ? sheet['B3'].v : '-';
    document.getElementById('summary-dest').innerText = sheet['B4'] ? sheet['B4'].v : '-';
    document.getElementById('summary-start').innerText = sheet['B5'] ? sheet['B5'].v : '-';
    document.getElementById('summary-end').innerText = sheet['B6'] ? sheet['B6'].v : '-';
}

// Dynamically scale data mapping for infinite row iterations
function parseHorizontalTab(workbook, sheetName, tableId, totalColumns) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;

    const tbody = document.querySelector(`#${tableId} tbody`);
    tbody.innerHTML = "";

    // Convert rows dynamically (Headers are on Row 3, data rows start on Row 4)
    let rowIndex = 4; 
    while (true) {
        let cellRef = `A${rowIndex}`;
        // Break loop safely if the date or name column row cell is completely empty
        if (!sheet[cellRef] || sheet[cellRef].v === undefined || sheet[cellRef].v === '') break; 

        let rowHtml = "<tr>";
        for (let colIndex = 0; colIndex < totalColumns; colIndex++) {
            let colLetter = XLSX.utils.encode_col(colIndex);
            let cell = sheet[`${colLetter}${rowIndex}`];
            
            // Format raw values or date strings smoothly
            let val = '';
            if (cell) {
                if (cell.w) {
                    val = cell.w; // Use formatted text version if available (keeps dates looking clean)
                } else if (cell.v !== undefined) {
                    val = cell.v;
                }
            }

            // Map the link reference targets cleanly into actionable links
            if (colIndex === (totalColumns - 1) && val) {
                rowHtml += `<td><a href="info.html?id=${encodeURIComponent(val)}" class="info-btn">More Info</a></td>`;
            } else {
                rowHtml += `<td>${val}</td>`;
            }
        }
        rowHtml += "</tr>";
        tbody.innerHTML += rowHtml;
        rowIndex++;
    }
}

// Render unstructured open text field inputs from cell A4 downwards
function parseOtherInfoTab(workbook) {
    const sheet = workbook.Sheets['Other Info'];
    if (sheet && sheet['A4']) {
        document.getElementById('other-info-content').innerText = sheet['A4'].v;
    }
}

