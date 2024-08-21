document.getElementById('dutyType').addEventListener('change', function() {
    var otherDutyField = document.getElementById('otherDutiesInput');
    
    if (this.value === 'OTHER DUTY') {
        otherDutyField.style.display = 'block';
        otherDutyField.required = true;  // Make the field required
    } else {
        otherDutyField.style.display = 'none';
        otherDutyField.required = false;  // Make the field optional
        otherDutyField.value = '';  // Clear the field if hidden
    }
});

document.getElementById('otherDutiesForm').addEventListener('submit', function(e) {
    e.preventDefault(); // Prevent the default form submission

    var form = this; // Reference to the form element
    var submitButton = document.getElementById('submitButton'); // Reference to the submit button
    var loadingMessage = document.getElementById('loadingMessage'); // Reference to the loading message element

    // Disable the submit button to prevent multiple submissions
    submitButton.disabled = true;
    submitButton.classList.add('loading'); // Add class to change button color
    loadingMessage.style.display = 'block'; // Show loading message

    var formData = new FormData(form);
    var dutyTypeSelect = document.getElementById('dutyType');
    var otherDutyField = document.getElementById('otherDutiesInput');
    
    // Include the text from 'OTHER DUTY' field if visible
    if (dutyTypeSelect.value === 'OTHER DUTY') {
        formData.set('DUTY TYPE', otherDutyField.value); // Update the 'DUTY TYPE' field with the text box value
    }

    // Debugging: Check form data
    for (var pair of formData.entries()) {
        console.log(pair[0] + ': ' + pair[1]);
    }

    // AJAX request to send data to Google Sheets
    fetch(form.action, {
        method: 'POST',
        body: formData
    })
    .then(response => response.text())
    .then(result => {
        console.log('Success:', result);
        alert('Form submitted successfully!');
        form.reset(); // Reset the form after user clicks OK
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred while submitting the form.');
    })
    .finally(() => {
        submitButton.disabled = false; // Enable the button again after submission
        submitButton.classList.remove('loading'); // Remove class to revert button color
        loadingMessage.style.display = 'none'; // Hide loading message
    });
});
