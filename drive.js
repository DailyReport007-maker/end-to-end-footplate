$(document).ready(function() {
    let crewData = [];

    // Load CSV data
    $.get('crewData.csv', function(data) {
        let rows = data.split('\n');
        rows.forEach(row => {
            let [crewId, crewName, desg, hq] = row.split(',');
            if (crewId && crewName && desg && hq) {
                crewData.push({
                    id: crewId,
                    text: crewId,
                    crewName: crewName,
                    desg: desg,
                    hq: hq
                });
            }
        });

        // Initialize Select2 for crewId with dynamic searching
        $('#crewId').select2({
            placeholder: 'Select CREW ID',
            data: crewData,
            matcher: function(params, data) {
                if ($.trim(params.term) === '') {
                    return data;
                }
                if (data.text.toLowerCase().indexOf(params.term.toLowerCase()) !== -1) {
                    return $.extend({}, data, true);
                }
                return null;
            }
        });
    });

    // On selection change, autofill crew details
    $('#crewId').on('select2:select', function(e) {
        let data = e.params.data;
        $('#crewName').val(data.crewName);
        $('#desg').val(data.desg);
        $('#hq').val(data.hq);
    });

    // Initialize Select2 for other selects if needed
    $('#cliName, #cliLobby').select2();

    // Ensure the form and submit button are present
    var form = document.getElementById('driveform');
    var submitButton = document.getElementById('submitButton');

    if (form && submitButton) {
        // Form submission
        form.addEventListener('submit', function(e) {
            e.preventDefault(); // Prevent the default form submission

            // Disable the submit button to prevent multiple submissions
            submitButton.disabled = true;

            var formData = new FormData(form);

            // Debugging: Check form data
            for (var pair of formData.entries()) {
                console.log(pair[0] + ': ' + pair[1]);
            }

            // AJAX request to send data to Google Sheets
            fetch('https://script.google.com/macros/s/AKfycby9ffTOr386IQpBR9MV6VjVYtPeBVJjYwb2b_lhEcvtgLjEb10102dk6QGqUbmH6fJQ/exec', {
                method: 'POST',
                body: formData
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error('Network response was not ok ' + response.statusText);
                }
                return response.json();
            })
            .then(result => {
                if (result.result === 'success') {
                    console.log('Success:', result);
                    alert('Form submitted successfully!');
                    // Only reset the form if needed
                    // form.reset(); 
                } else {
                    throw new Error(result.error);
                }
            })
            .catch(error => {
                console.error('Error:', error);
                alert('An error occurred while submitting the form.');
            })
            .finally(() => {
                submitButton.disabled = false; // Enable the button again after submission
            });
        });
    } else {
        console.error('Form or submit button not found.');
    }
});
