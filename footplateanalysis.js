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

    // Google Sheet Details (If applicable)
    const sheetId = '1fHSnNcPxryFY1JP2or3ZzqC4tO2qe6E1-VJ2-UX_nrQ';
    const apiKey = 'AIzaSyAw23pJz0K9fZb2rRRAe2C2cJDilRc0Kac';
    const sheetName = 'END TO END FOOTPLATE';

    let cliData = {};
    let reportData = [];

    // Load CLI options and details from CSV file
    fetch('CLI.csv')
        .then(response => response.text())
        .then(csvText => {
            const rows = csvText.split('\n');
            const uniqueCliNames = new Set();

            rows.forEach((row, index) => {
                if (index > 0 && row) { // Skip the header row
                    const columns = row.split(',');
                    const cliName = columns[0].trim();  // Fetch CLI Name
                    const lpId = columns[1].trim();
                    const lpName = columns[2].trim();
                    const desg = columns[3].trim();
                    const hq = columns[4].trim();

                    // Store LP details under the corresponding CLI Name
                    if (!cliData[cliName]) {
                        cliData[cliName] = [];
                    }

                    cliData[cliName].push({ cliName, lpId, lpName, desg, hq });

                    if (cliName) {
                        uniqueCliNames.add(cliName);
                    }
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

    function analyzeData(cliName, fromDate, toDate) {
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${sheetName}?key=${apiKey}`;

        fetch(url)
            .then(response => response.json())
            .then(data => {
                const rows = data.values;
                let cliLpIds = cliName === 'ALL' ? [].concat(...Object.values(cliData)) : cliData[cliName] || [];

                rows.forEach((row, index) => {
                    if (index > 0 && row.length > 0) { // Skip the header row
                        const sheetCliName = row[0].trim();
                        const lpId = row[1].trim();
                        const date = row[6].trim();

                        if ((cliName === 'ALL' || sheetCliName === cliName) &&
                            new Date(date) >= new Date(fromDate) &&
                            new Date(date) <= new Date(toDate)) {
                            cliLpIds = cliLpIds.filter(lp => lp.lpId !== lpId);
                        }
                    }
                });

                const lpNotDoneCount = cliLpIds.length;

                generateReport(cliName, lpNotDoneCount, cliLpIds);
                reportData = cliLpIds; // Store the report data for Excel generation
            })
            .catch(error => {
                console.error('Error fetching data:', error);
            });
    }

    function generateReport(cliName, lpNotDoneCount, lpDetails) {
        reportDiv.innerHTML = `
            <h2>Report for ${cliName}</h2>
            <p>Number of LPs not done End to End Footplate: <span id="lp-count">${lpNotDoneCount}</span></p>
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
            row.innerHTML = `
                <td>${lp.lpId}</td>
                <td>${lp.lpName}</td>
                <td>${lp.desg}</td>
                <td>${lp.hq}</td>
                <td>${lp.cliName}</td>  <!-- CLI Name from CSV -->
            `;
            detailsTableBody.appendChild(row);
        });

        detailsDiv.classList.remove('hidden');
    }

    // Handle Excel download
    downloadExcelBtn.addEventListener('click', function () {
        if (reportData.length > 0) {
            const worksheet = XLSX.utils.json_to_sheet(reportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Report');

            XLSX.writeFile(workbook, 'Report.xlsx');
        } else {
            alert('No data available to download.');
        }
    });
});
