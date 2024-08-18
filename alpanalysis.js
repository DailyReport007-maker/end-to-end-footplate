document.addEventListener('DOMContentLoaded', function() {
    populateDropdowns(); // Populate dropdowns when the page loads
});

async function populateDropdowns() {
    const response = await fetch('https://sheets.googleapis.com/v4/spreadsheets/1wb-xTJB4uM85A7V7aHoWJuXbRn0obxnqaFrvQJpPZi4/values/ALP END TO END!A:G?key=AIzaSyAw23pJz0K9fZb2rRRAe2C2cJDilRc0Kac');
    const data = await response.json();
    
    const cliNames = new Set();
    const alpIds = new Set();

    data.values.forEach(row => {
        cliNames.add(row[0]); // CLI NAME (Column A)
        alpIds.add(row[1]); // ALP ID (Column B)
    });

    // Convert sets to sorted arrays
    const sortedCliNames = Array.from(cliNames).sort();
    const sortedAlpIds = Array.from(alpIds).sort();

    // Populate CLI NAME datalist
    const cliNameList = document.getElementById('cliNameList');
    sortedCliNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        cliNameList.appendChild(option);
    });

    // Add 'ALL' option
    const allOptionCliName = document.createElement('option');
    allOptionCliName.value = 'ALL';
    allOptionCliName.textContent = 'ALL';
    cliNameList.appendChild(allOptionCliName);

    // Populate ALP ID datalist
    const alpIdList = document.getElementById('alpIdList');
    sortedAlpIds.forEach(id => {
        const option = document.createElement('option');
        option.value = id;
        alpIdList.appendChild(option);
    });

    // Add 'ALL' option
    const allOptionAlpId = document.createElement('option');
    allOptionAlpId.value = 'ALL';
    allOptionAlpId.textContent = 'ALL';
    alpIdList.appendChild(allOptionAlpId);

    // Initialize autocomplete for CLI NAME and ALP ID
    $('#cliName').autocomplete({
        source: sortedCliNames
    });

    $('#alpId').autocomplete({
        source: sortedAlpIds
    });
}

document.getElementById('dutyForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent form submission
    
    const cliName = document.getElementById('cliName').value;
    const alpId = document.getElementById('alpId').value;
    const fromDate = document.getElementById('fromDate').value;
    const toDate = document.getElementById('toDate').value;

    fetchData(cliName, alpId, fromDate, toDate);
});

async function fetchData(cliName, alpId, fromDate, toDate) {
    const response = await fetch('https://sheets.googleapis.com/v4/spreadsheets/1wb-xTJB4uM85A7V7aHoWJuXbRn0obxnqaFrvQJpPZi4/values/ALP END TO END!A:G?key=AIzaSyAw23pJz0K9fZb2rRRAe2C2cJDilRc0Kac');
    const data = await response.json();
    
    const filteredData = data.values.filter(row => {
        const cliNameMatches = cliName === 'ALL' || row[0] === cliName;
        const alpIdMatches = alpId === 'ALL' || row[1] === alpId;
        const dateMatches = new Date(row[6]) >= new Date(fromDate) && new Date(row[6]) <= new Date(toDate);
        return cliNameMatches && alpIdMatches && dateMatches;
    });

    displayData(filteredData);
}

function displayData(filteredData) {
    const tableBody = document.getElementById('reportFormBody');
    tableBody.innerHTML = '';

    filteredData.forEach(row => {
        const newRow = document.createElement('tr');

        row.forEach((cell, index) => {
            const cellElement = document.createElement('td');
            cellElement.textContent = cell || 'N/A';
            newRow.appendChild(cellElement);
        });

        tableBody.appendChild(newRow);
    });
}

function printReport() {
    // Hide the form
    document.getElementById('dutyForm').style.display = 'none';
    document.getElementById('printReport').style.display = 'none';

    // Print the report
    window.print();

    // Show the form again after printing
    document.getElementById('dutyForm').style.display = 'block';
    document.getElementById('printReport').style.display = 'block';
}
