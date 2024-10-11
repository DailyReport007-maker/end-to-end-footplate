document.addEventListener('DOMContentLoaded', function () {
    const cliNameSelect = document.getElementById('cli-name');
    const periodSelect = document.getElementById('period');
    const fromDateInput = document.getElementById('from-date');
    const toDateInput = document.getElementById('to-date');
    const quarterCheckboxes = document.querySelectorAll('.quarter-checkbox');
    const analyzeBtn = document.getElementById('analyze-btn');
    const reportDiv = document.getElementById('report');
    const detailsDiv = document.getElementById('details');
    const detailsTableBody = document.querySelector('#details-table tbody');
    const downloadExcelBtn = document.getElementById('download-excel-btn');

    // Google Sheet Details
    const sheetId = '1wb-xTJB4uM85A7V7aHoWJuXbRn0obxnqaFrvQJpPZi4';
    const apiKey = 'AIzaSyAw23pJz0K9fZb2rRRAe2C2cJDilRc0Kac';
    const sheetName = 'ALP END TO END';

    let cliData = {};
    let reportData = [];

    // Load CLI options and details from CSV file
    fetch('CLIALP.csv')
        .then(response => response.text())
        .then(csvText => {
            const rows = csvText.split('\n');
            const uniqueCliNames = new Set();

            rows.forEach((row, index) => {
                if (index > 0 && row) { // Skip the header row
                    const columns = row.split(',');
                    const cliName = columns[0].trim(); 
                    const lpId = columns[1].trim();
                    const lpName = columns[2].trim();
                    const desg = columns[3].trim();
                    const hq = columns[4].trim();

                    // Store LP details under the corresponding CLI Name
                    if (!cliData[cliName]) {
                        cliData[cliName] = [];
                    }

                    cliData[cliName].push({ cliName, lpId, lpName, desg, hq });
                    uniqueCliNames.add(cliName);
                }
            });

            // Populate the CLI dropdown with options
            uniqueCliNames.forEach(cliName => {
                const option = document.createElement('option');
                option.value = cliName;
                option.textContent = cliName;
                cliNameSelect.appendChild(option);
            });

            // Add 'ALL' option
            const allOption = document.createElement('option');
            allOption.value = 'ALL';
            allOption.textContent = 'ALL';
            cliNameSelect.appendChild(allOption);
        });

    // Automatically update TO DATE based on selected PERIOD and FROM DATE
    periodSelect.addEventListener('change', updateToDate);
    fromDateInput.addEventListener('change', updateToDate);
    quarterCheckboxes.forEach(checkbox => checkbox.addEventListener('change', updateToDate));

    function updateToDate() {
        const fromDate = new Date(fromDateInput.value);
        let toDate;

        if (periodSelect.value === 'QUARTERLY') {
            const selectedQuarters = Array.from(quarterCheckboxes)
                .filter(checkbox => checkbox.checked)
                .map(checkbox => parseInt(checkbox.value, 10));

            if (selectedQuarters.length > 0) {
                const firstQuarter = Math.min(...selectedQuarters);
                const lastQuarter = Math.max(...selectedQuarters);

                const quarterStartDates = {
                    1: '2024-01-01',
                    2: '2024-04-01',
                    3: '2024-07-01',
                    4: '2024-10-01',
                    5: '2025-01-01'
                };

                const quarterEndDates = {
                    1: '2024-03-31',
                    2: '2024-06-30',
                    3: '2024-10-07',
                    4: '2024-12-31',
                    5: '2025-03-31'
                };

                fromDateInput.value = quarterStartDates[firstQuarter];
                toDateInput.value = quarterEndDates[lastQuarter];
            }
        } else {
            switch (periodSelect.value) {
                case 'MONTHLY':
                    toDate = new Date(fromDate.getFullYear(), fromDate.getMonth() + 1, fromDate.getDate() - 1);
                    break;
                case 'HALFYEARLY':
                    toDate = new Date(fromDate.getFullYear(), fromDate.getMonth() + 6, fromDate.getDate() - 1);
                    break;
                case 'YEARLY':
                    toDate = new Date(fromDate.getFullYear() + 1, fromDate.getMonth(), fromDate.getDate() - 1);
                    break;
                default:
                    toDate = null;
            }

            if (toDate) {
                toDateInput.value = toDate.toISOString().split('T')[0];
            }
        }
    }

    // Handle Analysis button click
    analyzeBtn.addEventListener('click', function () {
        const cliName = cliNameSelect.value;
        const period = periodSelect.value;
        const fromDate = fromDateInput.value;
        const toDate = toDateInput.value;

        if (cliName && period && fromDate && toDate) {
            analyzeData(cliName, fromDate, toDate);
        } else {
            alert('Please fill all fields.');
        }
    });
    // Function to format date from yyyy-mm-dd to dd-mm-yyyy
function formatDate(dateString) {
    if (!dateString || dateString === 'N/A') return 'N/A';
    const dateParts = dateString.split('-');
    return `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`; // Converts to dd-mm-yyyy
}

function analyzeData(cliName, fromDate, toDate) {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${sheetName}?key=${apiKey}`;

    fetch(url)
        .then(response => response.json())
        .then(data => {
            const rows = data.values;
            let cliLpIds = cliName === 'ALL' ? [].concat(...Object.values(cliData)) : cliData[cliName] || [];

            const latestData = {}; // Store the latest entry for each LP ID

            // Step 1: Iterate over the entire sheet to find the latest entry for each LP ID
            rows.forEach((row, index) => {
                if (index > 0 && row.length > 0) { // Skip the header row
                    const sheetCliName = row[0]?.trim();
                    const lpId = row[1]?.trim();
                    const lastFootplateDate = row[6]?.trim(); // LAST FOOTPLATE DONE DATE
                    const beat = row[5]?.trim();

                    if (lpId && (cliName === 'ALL' || sheetCliName === cliName)) {
                        // Keep only the latest entry for each LP ID
                        if (!latestData[lpId] || new Date(lastFootplateDate) > new Date(latestData[lpId].lastFootplateDate)) {
                            latestData[lpId] = { lastFootplateDate, beat };
                        }

                        // Step 2: Remove LP IDs that have data within the selected period
                        if (new Date(lastFootplateDate) >= new Date(fromDate) &&
                            new Date(lastFootplateDate) <= new Date(toDate)) {
                            cliLpIds = cliLpIds.filter(lp => lp.lpId !== lpId);
                        }
                    }
                }
            });

            // Step 3: Prepare the report data for LP IDs not having data in the selected period
            const lpDetails = cliLpIds.map(lp => {
                const sheetEntry = latestData[lp.lpId] || {}; // Get latest entry, or empty if no data exists

                let lastFootplateDone = sheetEntry.lastFootplateDate ? formatDate(sheetEntry.lastFootplateDate) : 'N/A';
                let beat = sheetEntry.beat || 'N/A';
                let dueDate = 'N/A';

                // For all entries, calculate 30-day due date from the last footplate date
                if (lastFootplateDone !== 'N/A') {
                    dueDate = calculateDueDate(lastFootplateDone);
                }

                return {
                    ...lp,  // Include lpName, desg, hq, cliName from CSV
                    lastFootplateDone,
                    beat,
                    dueDate
                };
            });

            const lpNotDoneCount = lpDetails.length;

            generateReport(cliName, lpNotDoneCount, lpDetails);
            reportData = lpDetails; // Store the report data for Excel generation
        })
        .catch(error => {
            console.error('Error fetching data:', error);
        });
}

function calculateDueDate(lastFootplateDate) {
    const dateParts = lastFootplateDate.split('-'); // Expecting dd-mm-yyyy
    const lastDate = new Date(`${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`); // Convert to yyyy-mm-dd
    if (isNaN(lastDate.getTime())) {
        console.error('Invalid last footplate date:', lastFootplateDate);
        return 'N/A'; // Return N/A if the date is invalid
    }

    // Always calculate 30 days for all entries
    const dueDate = new Date(lastDate.setDate(lastDate.getDate() + 30));

    // Return the due date in dd-mm-yyyy format
    return formatDate(`${dueDate.getDate().toString().padStart(2, '0')}-${(dueDate.getMonth() + 1).toString().padStart(2, '0')}-${dueDate.getFullYear()}`);
}

    
    
    function generateReport(cliName, lpNotDoneCount, lpDetails) {
        reportDiv.innerHTML = `
            <h2>Report for ${cliName}</h2>
            <p>Number of LPs not done End to End Footplate: <span id="lp-count" class="highlighted">${lpNotDoneCount}</span></p>
            <div class="legend">
                <p><span class="legend-box purple"></span> Due in 10 days or less</p>
                <p><span class="legend-box red"></span> Overdue</p>
            </div>
        `;
    
        document.getElementById('lp-count').addEventListener('click', function () {
            populateDetailsTable(lpDetails);
        });
    
        reportDiv.style.display = 'block';
        downloadExcelBtn.style.display = 'block';  // Show the download button after report is generated
    }
    
    
    function populateDetailsTable(lpDetails) {
        detailsTableBody.innerHTML = ''; // Clear previous details
    
        lpDetails.forEach(lp => {
            const row = document.createElement('tr');
            const dueDate = lp.dueDate !== 'N/A' ? new Date(lp.dueDate) : null; // Directly parse yyyy-mm-dd format
            const today = new Date(); // Get today's date
    
            if (dueDate) {
                const timeDiff = dueDate.getTime() - today.getTime(); // Time difference in milliseconds
                const daysDiff = timeDiff / (1000 * 3600 * 24); // Convert milliseconds to days
    
                // Highlight the row based on days difference
                if (daysDiff <= 0) {
                    row.classList.add('overdue-row'); // Add red color for overdue
                } else if (daysDiff <= 10) {
                    row.classList.add('warning-row'); // Add purple color for 10 days before due date
                }
            }
    
            // Insert data into the row
            row.innerHTML = `
                <td>${lp.lpId}</td>
                <td>${lp.lpName}</td>
                <td>${lp.desg}</td>
                <td>${lp.hq}</td>
                <td>${lp.cliName}</td>
                <td>${lp.lastFootplateDone}</td>
                <td>${lp.beat}</td>
                <td>${lp.dueDate}</td>
            `;
            detailsTableBody.appendChild(row);
        });
    
        detailsDiv.classList.remove('hidden');
    }
    
    
    downloadExcelBtn.addEventListener('click', function () {
        // Prepare data for Excel
        const worksheetData = [
            ['CLI Name', 'ALP ID', 'ALP Name', 'Designation', 'HQ', 'Last Footplate Done', 'Beat', 'Due Date'],  // Header row
            ...reportData.map(lp => [
                lp.cliName, lp.lpId, lp.lpName, lp.desg, lp.hq, 
                lp.lastFootplateDone, lp.beat, lp.dueDate
            ])
        ];
    
        // Create a new workbook and worksheet
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(worksheetData);
    
        // Loop through reportData to highlight due dates
        reportData.forEach((lp, index) => {
            const dueDateCell = `H${index + 2}`; // Assuming the Due Date is in column H (8th column)
            const dueDateParts = lp.dueDate.split('-'); // Split the due date (yyyy-mm-dd)
            const dueDate = new Date(dueDateParts[0], dueDateParts[1] - 1, dueDateParts[2]); // Create date object
            const today = new Date();
            const tenDaysBeforeDueDate = new Date(dueDate);
            tenDaysBeforeDueDate.setDate(tenDaysBeforeDueDate.getDate() - 10);
    
            // Apply highlighting based on the due date
            if (dueDate < today) {
                ws[dueDateCell].s = {
                    fill: {
                        fgColor: { rgb: "FFCCCB" }, // Red background for overdue
                    },
                };
            } else if (today >= tenDaysBeforeDueDate && today < dueDate) {
                ws[dueDateCell].s = {
                    fill: {
                        fgColor: { rgb: "E6E6FA" }, // Purple background for 10 days before due date
                    },
                };
            }
        });
    
        // Append the worksheet to the workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Report');
    
        // Generate Excel file and trigger download
        XLSX.writeFile(wb, 'report.xlsx');
    });
    
});
