<?php
session_start();
require "config.php";

$error = "";
$success = $_SESSION["success"] ?? "";
unset($_SESSION["success"]);

if ($_SERVER["REQUEST_METHOD"] === "POST") {
    $username = trim($_POST["username"] ?? "");
    $password = $_POST["password"] ?? "";

    if ($username === "" || $password === "") {
        $error = "الرجاء إدخال اسم المستخدم وكلمة المرور";
    } else {
        $stmt = $pdo->prepare("SELECT id, username, password FROM users WHERE username = ?");
        $stmt->execute([$username]);
        $user = $stmt->fetch();

        if ($user && password_verify($password, $user["password"])) {
            $_SESSION["user_id"] = $user["id"];
            $_SESSION["username"] = $user["username"];
            header("Location: home.php");
            exit;
        } else {
            $error = "اسم المستخدم أو كلمة المرور غير صحيحة";
        }
    }
}
?>
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>تسجيل الدخول</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <div class="card">
        <h2>تسجيل الدخول</h2>

        <?php if ($error): ?>
            <div class="msg error"><?= htmlspecialchars($error) ?></div>
        <?php endif; ?>

        <?php if ($success): ?>
            <div class="msg success"><?= htmlspecialchars($success) ?></div>
        <?php endif; ?>

        <form method="POST" action="login.php">
            <input type="text" name="username" placeholder="اسم المستخدم" value="<?= htmlspecialchars($_POST['username'] ?? '') ?>" required>
            <input type="password" name="password" placeholder="كلمة المرور" required>
            <button type="submit">الحفظ</button>
        </form>

        <a href="register.php">مستخدم جديد؟ التسجيل</a>
    </div>
</body>
</html>
