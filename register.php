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
</head>
<body>
    <h2>تسجيل مستخدم جديد</h2>

    <?php if (isset($error)) { echo $error; } ?>

    <form method="POST" action="register.php">
        اسم المستخدم: <input type="text" name="username"><br><br>
        كلمة المرور: <input type="password" name="password"><br><br>
        تأكيد كلمة المرور: <input type="password" name="confirm_password"><br><br>
        <input type="submit" name="submit" value="التسجيل">
    </form>

    <a href="login.php">لديك حساب؟ تسجيل الدخول</a>
</body>
</html>
