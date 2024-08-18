let barChart, pieChart;

document.addEventListener('DOMContentLoaded', function() {
    populateDropdowns();
    document.getElementById('printButton').addEventListener('click', printReport);
    document.getElementById('pdfButton').addEventListener('click', downloadPDF);
});

async function populateDropdowns() {
    const response = await fetch('https://sheets.googleapis.com/v4/spreadsheets/1Pn0sosMuxX2XZYRe9qIvkGM06GcyIWdEeaD0gfPeRZU/values/DAILY REPORT!A:AE?key=AIzaSyAw23pJz0K9fZb2rRRAe2C2cJDilRc0Kac');
    const data = await response.json();
    
    const cliNames = new Set();
    const cliLobbies = new Set();

    data.values.forEach(row => {
        cliNames.add(row[0]); // CLI NAME (Column A)
        cliLobbies.add(row[1]); // CLI LOBBY (Column B)
    });

    // Convert sets to sorted arrays
    const sortedCliNames = Array.from(cliNames).sort();
    const sortedCliLobbies = Array.from(cliLobbies).sort();

    // Populate CLI NAME dropdown
    const cliNameSelect = document.getElementById('cliName');
    sortedCliNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        cliNameSelect.appendChild(option);
    });

    // Add 'ALL' option
    const allOptionCliName = document.createElement('option');
    allOptionCliName.value = 'ALL';
    allOptionCliName.textContent = 'ALL';
    cliNameSelect.appendChild(allOptionCliName);

    // Populate CLI LOBBY dropdown
    const cliLobbySelect = document.getElementById('cliLobby');
    sortedCliLobbies.forEach(lobby => {
        const option = document.createElement('option');
        option.value = lobby;
        option.textContent = lobby;
        cliLobbySelect.appendChild(option);
    });

    // Add 'ALL' option
    const allOptionCliLobby = document.createElement('option');
    allOptionCliLobby.value = 'ALL';
    allOptionCliLobby.textContent = 'ALL';
    cliLobbySelect.appendChild(allOptionCliLobby);
}

document.getElementById('dutyForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent form submission
    
    const cliName = document.getElementById('cliName').value;
    const cliLobby = document.getElementById('cliLobby').value;
    const fromDate = document.getElementById('fromDate').value;
    const toDate = document.getElementById('toDate').value;

    fetchData(cliName, cliLobby, fromDate, toDate);
});

async function fetchData(cliName, cliLobby, fromDate, toDate) {
    const response = await fetch('https://sheets.googleapis.com/v4/spreadsheets/1Pn0sosMuxX2XZYRe9qIvkGM06GcyIWdEeaD0gfPeRZU/values/DAILY REPORT!A:AE?key=AIzaSyAw23pJz0K9fZb2rRRAe2C2cJDilRc0Kac');
    const data = await response.json();
    
    const filteredData = data.values.filter(row => {
        const cliNameMatches = cliName === 'ALL' || row[0] === cliName;
        const cliLobbyMatches = cliLobby === 'ALL' || row[1] === cliLobby;
        const dateMatches = new Date(row[2]) >= new Date(fromDate) && new Date(row[2]) <= new Date(toDate);
        return cliNameMatches && cliLobbyMatches && dateMatches;
    });

    displayData(filteredData);
}

function displayData(filteredData) {
    const tableBody = document.getElementById('reportFormBody');
    tableBody.innerHTML = '';

    const parameters = {
        'FOOTPLATE DURING WEE HOURS': 3,
        'FOOTPLATE EXCLUDING WEE HOURS': 5,
        'SPECIAL FOOTPLATE': 7,
        'AUTO AND IB AMBUSH(NO OF TRAINS)': 8,
        'LEVEL CROSSING AMBUSH(NO OF TRAINS)': 10,
        'CD ZONES AMBUSH(NO OF TRAINS)': 12,
        'BA AMBUSH(NO OF CREWS)': 14,
        'SPEED GUN CHECKING(NO OF TRAINS)': 16,
        'COUNCELLING AT LOBBY CONCERNED TO SAFETY(NO OF CREWS)': 18,
        'CVVRS ANALYSIS(NO OF TRAINS)': 19,
        'SPM ANALYSIS(NO OF TRAINS)': 21,
        'SHUNTING PROCEDURE & RULES INVOLVED AS PER GR 5.13': 23,
        'COUNSELLING FOR PERSONAL SAFETY OF CREW': 24,
        'LONG HOURS DUTY OF CREW': 25,
        'FAMILY/SAFETY SEMINAR': 26,
        'RUNNING ROOM INSPECTION': 27,
        'LOBBY INSPECTION': 29
    };

    const abnormalityColumns = {
        'FOOTPLATE DURING WEE HOURS': 4,
        'FOOTPLATE EXCLUDING WEE HOURS': 6,
        'AUTO AND IB AMBUSH(NO OF TRAINS)': 9,
        'LEVEL CROSSING AMBUSH(NO OF TRAINS)': 11,
        'CD ZONES AMBUSH(NO OF TRAINS)': 13,
        'BA AMBUSH(NO OF CREWS)': 15,
        'SPEED GUN CHECKING(NO OF TRAINS)': 17,
        'CVVRS ANALYSIS(NO OF TRAINS)': 20,
        'SPM ANALYSIS(NO OF TRAINS)': 22,
        'RUNNING ROOM INSPECTION': 28,
        'LOBBY INSPECTION': 30
    };

    const total = {};
    const abnormalities = {};

    filteredData.forEach(row => {
        Object.keys(parameters).forEach(param => {
            const index = parameters[param];
            const value = parseInt(row[index], 10) || 0;
            if (!total[param]) total[param] = 0;
            total[param] += value;
        });

        Object.keys(abnormalityColumns).forEach(param => {
            const index = abnormalityColumns[param];
            const value = parseInt(row[index], 10) || 0;
            if (!abnormalities[param]) abnormalities[param] = 0;
            abnormalities[param] += value;
        });
    });

    Object.keys(parameters).forEach(param => {
        const newRow = document.createElement('tr');

        const paramCell = document.createElement('td');
        paramCell.textContent = param;
        newRow.appendChild(paramCell);

        const doneCell = document.createElement('td');
        doneCell.textContent = total[param] || '0';
        newRow.appendChild(doneCell);

        const abnormalityCell = document.createElement('td');
        abnormalityCell.textContent = abnormalities[param] || '0';
        newRow.appendChild(abnormalityCell);

        tableBody.appendChild(newRow);
    });

    renderBarChart(total);
    renderPieChart(abnormalities);
}

function renderBarChart(total) {
    const ctx = document.getElementById('barGraphCanvas').getContext('2d');
    if (barChart) barChart.destroy(); // Destroy previous chart instance if it exists

    barChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: Object.keys(total),
            datasets: [{
                label: 'Total',
                data: Object.values(total),
                backgroundColor: 'rgba(75, 192, 192, 0.2)',
                borderColor: 'rgba(75, 192, 192, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

function renderPieChart(abnormalities) {
    const ctx = document.getElementById('pieChartCanvas').getContext('2d');
    if (pieChart) pieChart.destroy(); // Destroy previous chart instance if it exists

    pieChart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: Object.keys(abnormalities),
            datasets: [{
                label: 'Abnormalities',
                data: Object.values(abnormalities),
                backgroundColor: [
                    'rgba(255, 99, 132, 0.2)',
                    'rgba(54, 162, 235, 0.2)',
                    'rgba(255, 206, 86, 0.2)',
                    'rgba(75, 192, 192, 0.2)',
                    'rgba(153, 102, 255, 0.2)',
                    'rgba(255, 159, 64, 0.2)'
                ],
                borderColor: [
                    'rgba(255, 99, 132, 1)',
                    'rgba(54, 162, 235, 1)',
                    'rgba(255, 206, 86, 1)',
                    'rgba(75, 192, 192, 1)',
                    'rgba(153, 102, 255, 1)',
                    'rgba(255, 159, 64, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true
        }
    });
}

function printReport() {
    const reportForm = document.getElementById('reportForm');
    const originalContent = document.body.innerHTML;

    const newWindow = window.open('', '', 'width=800,height=600');
    newWindow.document.open();
    newWindow.document.write(`
        <html>
            <head>
                <title>Print Report</title>
                <style>
                    @media print {
                        .no-print { display: none; }
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                    }
                    table, th, td {
                        border: 1px solid black;
                    }
                    th, td {
                        padding: 8px;
                        text-align: left;
                    }
                </style>
            </head>
            <body>
                <h1>DAILY REPORT ANALYSIS of ${document.getElementById('cliName').value} from ${document.getElementById('fromDate').value} to ${document.getElementById('toDate').value}</h1>
                ${reportForm.outerHTML}
            </body>
        </html>
    `);
    newWindow.document.close();

    newWindow.onload = function() {
        newWindow.print();
    };
}

function downloadPDF() {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF();
    
    // Set up the PDF heading
    const heading = `DAILY REPORT ANALYSIS of ${document.getElementById('cliName').value} from ${document.getElementById('fromDate').value} to ${document.getElementById('toDate').value}`;
    pdf.setFontSize(16);
    pdf.text(heading, 10, 10);
    
    // Setup table data
    const reportTable = document.getElementById('reportForm');
    const rows = reportTable.querySelectorAll('tr');
    
    let y = 30;
    const margin = 10;
    const lineHeight = 10;
    const colWidths = [130, 30, 30]; // Adjust column widths
    
    // Table headers
    const headers = Array.from(rows[0].children).map(cell => cell.textContent);
    pdf.setFontSize(10);
    
    // Draw headers
    pdf.rect(margin, y - lineHeight, colWidths[0], lineHeight);
    pdf.text(headers[0], margin + 2, y - 2);
    
    pdf.rect(margin + colWidths[0], y - lineHeight, colWidths[1], lineHeight);
    pdf.text(headers[1], margin + colWidths[0] + 2, y - 2);
    
    pdf.rect(margin + colWidths[0] + colWidths[1], y - lineHeight, colWidths[2], lineHeight);
    pdf.text(headers[2], margin + colWidths[0] + colWidths[1] + 2, y - 2);
    
    // Draw rows
    rows.forEach((row, rowIndex) => {
        if (rowIndex === 0) return; // Skip header row
        const cells = Array.from(row.children).map(cell => cell.textContent);
        y += lineHeight;
        
        // Draw and fill each cell
        pdf.rect(margin, y - lineHeight, colWidths[0], lineHeight);
        pdf.text(cells[0], margin + 2, y - 2);
        
        pdf.rect(margin + colWidths[0], y - lineHeight, colWidths[1], lineHeight);
        pdf.text(cells[1], margin + colWidths[0] + 2, y - 2);
        
        pdf.rect(margin + colWidths[0] + colWidths[1], y - lineHeight, colWidths[2], lineHeight);
        pdf.text(cells[2], margin + colWidths[0] + colWidths[1] + 2, y - 2);
    });
    
    // Save the PDF
    pdf.save('report.pdf');
}

