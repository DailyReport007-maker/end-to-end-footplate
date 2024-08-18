document.getElementById('abnormalityReportForm').addEventListener('submit', function(e) {
    e.preventDefault(); // Prevent the default form submission

    var form = this; // Reference to the form element
    var submitButton = document.getElementById('submitButton'); // Reference to the submit button

    // Disable the submit button to prevent multiple submissions
    submitButton.disabled = true;

    var formData = new FormData(form);

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
        form.reset(); // Reset the form
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred while submitting the form.');
    })
    .finally(() => {
        submitButton.disabled = false; // Enable the button again after submission
    });
});