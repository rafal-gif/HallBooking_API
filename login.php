<?php
session_start();
include "config.php";

if (isset($_POST["submit"])) {
    $username = $_POST["username"];
    $password = $_POST["password"];

    $sql = "SELECT * FROM users WHERE username='$username' AND password='$password'";
    $result = mysqli_query($conn, $sql);

    if (mysqli_num_rows($result) > 0) {
        $_SESSION["username"] = $username;
        header("Location: home.php");
        exit;
    } else {
        $error = "اسم المستخدم أو كلمة المرور غير صحيحة";
    }
}
?>
<html>
<head>
    <title>تسجيل الدخول</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <h2>تسجيل الدخول</h2>

    <?php if (isset($error)) { echo "<p class='error'>$error</p>"; } ?>

    <form method="POST" action="login.php" onsubmit="return checkFields()">
        اسم المستخدم: <input type="text" name="username" id="username"><br><br>
        كلمة المرور: <input type="password" name="password" id="password"><br><br>
        <input type="submit" name="submit" value="الحفظ">
    </form>

    <br>
    <a href="register.php">مستخدم جديد؟ التسجيل</a>

    <script>
        function checkFields() {
            var username = document.getElementById("username").value;
            var password = document.getElementById("password").value;

            if (username === "" || password === "") {
                alert("الرجاء تعبئة جميع الحقول");
                return false;
            }
            return true;
        }
    </script>
</body>
</html>
