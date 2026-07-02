<?php
include "config.php";

if (isset($_POST["submit"])) {
    $username = $_POST["username"];
    $password = $_POST["password"];
    $confirm_password = $_POST["confirm_password"];

    if ($password == $confirm_password) {
        $sql = "INSERT INTO users (username, password) VALUES ('$username', '$password')";
        mysqli_query($conn, $sql);
        header("Location: login.php");
        exit;
    } else {
        $error = "كلمة المرور وتأكيدها غير متطابقين";
    }
}
?>
<html>
<head>
    <title>تسجيل مستخدم جديد</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <h2>تسجيل مستخدم جديد</h2>

    <?php if (isset($error)) { echo "<p class='error'>$error</p>"; } ?>

    <form method="POST" action="register.php" onsubmit="return checkPassword()">
        اسم المستخدم: <input type="text" name="username"><br><br>
        كلمة المرور: <input type="password" name="password" id="password"><br><br>
        تأكيد كلمة المرور: <input type="password" name="confirm_password" id="confirm_password"><br><br>
        <input type="submit" name="submit" value="التسجيل">
    </form>

    <br>
    <a href="login.php">لديك حساب؟ تسجيل الدخول</a>

    <script>
        function checkPassword() {
            var password = document.getElementById("password").value;
            var confirm = document.getElementById("confirm_password").value;

            if (password !== confirm) {
                alert("كلمة المرور وتأكيدها غير متطابقين");
                return false;
            }
            return true;
        }
    </script>
</body>
</html>
