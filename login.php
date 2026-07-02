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
</head>
<body>
    <h2>تسجيل الدخول</h2>

    <?php if (isset($error)) { echo $error; } ?>

    <form method="POST" action="login.php">
        اسم المستخدم: <input type="text" name="username"><br><br>
        كلمة المرور: <input type="password" name="password"><br><br>
        <input type="submit" name="submit" value="الحفظ">
    </form>

    <a href="register.php">مستخدم جديد؟ التسجيل</a>
</body>
</html>
