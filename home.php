<?php
session_start();

if (!isset($_SESSION["username"])) {
    header("Location: login.php");
    exit;
}
?>
<html>
<head>
    <title>الصفحة الرئيسية</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <h2>الصفحة الرئيسية</h2>
    <p>أهلاً بك، <?php echo $_SESSION["username"]; ?></p>
    <a href="logout.php" onclick="return confirm('هل تريد تسجيل الخروج؟')">تسجيل الخروج</a>
</body>
</html>
