<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

echo '<script>alert("Email sent successfully!");</script>';

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $to = "agustinhumpreyzoe@gmail.com"; // Change this to your email address
    $subject = "Password Reset Request";
    $email = $_POST['email'];
    $message = "A password reset request has been made for the email: $email";
    $headers = "agustinhumpreyzoe@gmail.com"; // Change this to your email or website name

    if (mail($to, $subject, $message, $headers)) {
        echo '<script>alert("Email sent successfully!");</script>';
    } else {
        echo '<script>alert("Email sending failed!");</script>';
    }
}
?>
