<?php
session_start();

if (!isset($_SESSION["user_id"])) {
    header("Location: login.php");
    exit;
}
?>
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>الصفحة الرئيسية</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <div class="card home-box">
        <h2>الصفحة الرئيسية</h2>
        <p>أهلاً بك، <span class="username"><?= htmlspecialchars($_SESSION["username"]) ?></span></p>
        <a href="logout.php">تسجيل الخروج</a>
    </div>
</body>
</html>
