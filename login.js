// Wait for the DOM to be fully loaded
document.addEventListener('DOMContentLoaded', function() {
    // Get the form element
    var form = document.getElementById('loginForm');

    function checkValidation(){
        var emailInput = document.getElementById('email');
        var passwordInput = document.getElementById('password');
        var isValid = true;

         // Reset custom validation messages
         emailInput.setCustomValidity('');
         passwordInput.setCustomValidity('');
 
         // Check if username and password match
         if (emailInput.value !== 'joydonato1' || passwordInput.value !== '123') {
             emailInput.setCustomValidity('Invalid username or password');
             isValid = false;
         }
 

        if (emailInput == "" || passwordInput == ""){
            alert("Username and Password must be filled out");
        }

    }
    /*
    // Add a submit event listener to the form
    form.addEventListener('submit', function(event) {
        // Get the email and password input elements
        var emailInput = document.getElementById('email');
        var passwordInput = document.getElementById('password');
        var isValid = true;

        // Reset custom validation messages
        emailInput.setCustomValidity('');
        passwordInput.setCustomValidity('');

        // Check if username and password match
        if (emailInput.value !== 'joydonato1' || passwordInput.value !== '123') {
            emailInput.setCustomValidity('Invalid username or password');
            isValid = false;
        }

        // If the form is invalid, prevent submission
        if (!isValid) {
            event.preventDefault();
            event.stopPropagation();
        } else {
            // Redirect to the "hello world" page if login is successful
            window.location.href = 'hello.html';
        }
        // Add Bootstrap's was-validated class to show validation feedback
        form.classList.add('was-validated');
    }, false);
    */
});
