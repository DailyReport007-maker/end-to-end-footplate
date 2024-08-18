document.addEventListener('DOMContentLoaded', function() {
    populateDropdowns(); // Populate dropdowns when the page loads
});

async function populateDropdowns() {
    const response = await fetch('https://sheets.googleapis.com/v4/spreadsheets/1fHSnNcPxryFY1JP2or3ZzqC4tO2qe6E1-VJ2-UX_nrQ/values/END TO END FOOTPLATE!A:G?key=AIzaSyAw23pJz0K9fZb2rRRAe2C2cJDilRc0Kac');
    const data = await response.json();
    
    const cliNames = new Set();
    const lpIds = new Set();

    data.values.forEach(row => {
        cliNames.add(row[0]); // CLI NAME (Column A)
        lpIds.add(row[1]); // LP ID (Column B)
    });

    // Convert sets to sorted arrays
    const sortedCliNames = Array.from(cliNames).sort();
    const sortedLpIds = Array.from(lpIds).sort();

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

    // Populate LP ID datalist
    const lpIdList = document.getElementById('lpIdList');
    sortedLpIds.forEach(id => {
        const option = document.createElement('option');
        option.value = id;
        lpIdList.appendChild(option);
    });

    // Add 'ALL' option
    const allOptionLpId = document.createElement('option');
    allOptionLpId.value = 'ALL';
    allOptionLpId.textContent = 'ALL';
    lpIdList.appendChild(allOptionLpId);

    // Initialize autocomplete for CLI NAME and LP ID
    $('#cliName').autocomplete({
        source: sortedCliNames
    });

    $('#lpId').autocomplete({
        source: sortedLpIds
    });
}

document.getElementById('dutyForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent form submission
    
    const cliName = document.getElementById('cliName').value;
    const lpId = document.getElementById('lpId').value;
    const fromDate = document.getElementById('fromDate').value;
    const toDate = document.getElementById('toDate').value;

    fetchData(cliName, lpId, fromDate, toDate);
});

async function fetchData(cliName, lpId, fromDate, toDate) {
    const response = await fetch('https://sheets.googleapis.com/v4/spreadsheets/1fHSnNcPxryFY1JP2or3ZzqC4tO2qe6E1-VJ2-UX_nrQ/values/END TO END FOOTPLATE!A:G?key=AIzaSyAw23pJz0K9fZb2rRRAe2C2cJDilRc0Kac');
    const data = await response.json();
    
    const filteredData = data.values.filter(row => {
        const cliNameMatches = cliName === 'ALL' || row[0] === cliName;
        const lpIdMatches = lpId === 'ALL' || row[1] === lpId;
        const dateMatches = new Date(row[6]) >= new Date(fromDate) && new Date(row[6]) <= new Date(toDate);
        return cliNameMatches && lpIdMatches && dateMatches;
    });

    displayData(filteredData);
}

function displayData(filteredData) {
    const tableBody = document.getElementById('reportFormBody');
    tableBody.innerHTML = '';

    filteredData.forEach(row => {
        const newRow = document.createElement('tr');

        row.forEach(cell => {
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
