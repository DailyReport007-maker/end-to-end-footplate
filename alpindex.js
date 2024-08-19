// Fetch the CSV file and parse it using PapaParse
fetch('CLIALP.csv')
    .then(response => response.text())
    .then(csvText => {
        const data = Papa.parse(csvText, {
            header: true,
            skipEmptyLines: true
        }).data;

        populateCliNames(data);
    })
    .catch(error => {
        console.error('Error fetching the CSV file:', error);
    });

// Populate the CLI Names dropdown with unique values from the CSV data
function populateCliNames(data) {
    const cliNames = [...new Set(data.map(item => item['CLI NAME']))]; // Get unique CLI NAMEs
    const cliNameDropdown = document.getElementById('cliName');

    cliNames.forEach(cli => {
        const option = document.createElement('option');
        option.value = cli;
        option.textContent = cli;
        cliNameDropdown.appendChild(option);
    });

    cliNameDropdown.addEventListener('change', function() {
        populateLpIds(data, this.value);
    });
}

// Populate the ALP ID dropdown based on the selected CLI Name
function populateLpIds(data, selectedCliName) {
    const filteredData = data.filter(item => item['CLI NAME'] === selectedCliName);
    const lpIdDropdown = document.getElementById('alpId');

    lpIdDropdown.innerHTML = '<option value="">Select ALP ID</option>'; // Reset ALP ID dropdown

    filteredData.forEach(item => {
        const option = document.createElement('option');
        option.value = item['ALP ID'];
        option.textContent = item['ALP ID'];
        lpIdDropdown.appendChild(option);
    });

    lpIdDropdown.addEventListener('change', function() {
        autoFillDetails(filteredData, this.value);
    });
}

// Auto-fill the form details based on the selected ALP ID
function autoFillDetails(data, selectedLpId) {
    const selectedData = data.find(item => item['ALP ID'] === selectedLpId);

    document.getElementById('alpName').value = selectedData['ALP NAME'] || '';
    document.getElementById('desg').value = selectedData['DESG'] || '';
    document.getElementById('hq').value = selectedData['HQ'] || '';
}

// Handle the BEAT field, including the "OTHER" option
document.getElementById('BEAT').addEventListener('change', function() {
    var otherBeatInput = document.getElementById('otherBeatInput');
    if (this.value === 'OTHER') {
        otherBeatInput.style.display = 'block';
        otherBeatInput.required = true;
    } else {
        otherBeatInput.style.display = 'none';
        otherBeatInput.required = false;
    }
});

// Handle form submission
document.getElementById('footplateForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent the default form submission

    const submitButton = this.querySelector('button[type="submit"]');
    submitButton.disabled = true; // Disable the submit button to prevent multiple submissions

    // Handle the BEAT field if "OTHER" is selected
    var selectBeat = document.getElementById('BEAT');
    var otherBeatInput = document.getElementById('otherBeatInput');
    
    // If "OTHER" is selected, replace the BEAT value with the custom input value
    if (selectBeat.value === 'OTHER') {
        selectBeat.setAttribute('name', 'BEAT_OTHER');
        otherBeatInput.setAttribute('name', 'BEAT');
    }

    var formData = new FormData(this);
    var data = {};
    formData.forEach((value, key) => data[key] = value);

    console.log('Final Data being submitted:', data); // Debugging line

    fetch('https://script.google.com/macros/s/AKfycbwxnVSd3eGB4SvHVz2IINLEat031qB4xvNKbe8BcdPyg5BQpL0xkbQuAyu_B-fTdIqB0A/exec', { // Replace with your actual Web App URL
        method: 'POST',
        mode: 'no-cors',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: new URLSearchParams(data)
    })
    .then(response => {
        alert('Form data submitted successfully!');
        document.getElementById('footplateForm').reset(); // Reset the form after submission

        // Reset the BEAT field handling
        selectBeat.setAttribute('name', 'BEAT');
        otherBeatInput.setAttribute('name', 'BEAT_OTHER');
    })
    .catch(error => {
        console.error('Error:', error);
    })
    .finally(() => {
        submitButton.disabled = false; // Re-enable the submit button after submission is complete
    });
});
