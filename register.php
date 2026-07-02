<?php
session_start();
require "config.php";

$error = "";

if ($_SERVER["REQUEST_METHOD"] === "POST") {
    $username = trim($_POST["username"] ?? "");
    $password = $_POST["password"] ?? "";
    $confirm_password = $_POST["confirm_password"] ?? "";

    if ($username === "" || $password === "" || $confirm_password === "") {
        $error = "الرجاء تعبئة جميع الحقول";
    } elseif ($password !== $confirm_password) {
        $error = "كلمة المرور وتأكيدها غير متطابقين";
    } else {
        $stmt = $pdo->prepare("SELECT id FROM users WHERE username = ?");
        $stmt->execute([$username]);

        if ($stmt->fetch()) {
            $error = "اسم المستخدم موجود مسبقًا";
        } else {
            $hashed_password = password_hash($password, PASSWORD_DEFAULT);
            $stmt = $pdo->prepare("INSERT INTO users (username, password) VALUES (?, ?)");
            $stmt->execute([$username, $hashed_password]);

            $_SESSION["success"] = "تم التسجيل بنجاح، الرجاء تسجيل الدخول";
            header("Location: login.php");
            exit;
        }
    }
}
?>
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>تسجيل مستخدم جديد</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <div class="card">
        <h2>تسجيل مستخدم جديد</h2>

        <?php if ($error): ?>
            <div class="msg error"><?= htmlspecialchars($error) ?></div>
        <?php endif; ?>

        <form method="POST" action="register.php">
            <input type="text" name="username" placeholder="اسم المستخدم" value="<?= htmlspecialchars($_POST['username'] ?? '') ?>" required>
            <input type="password" name="password" placeholder="كلمة المرور" required>
            <input type="password" name="confirm_password" placeholder="تأكيد كلمة المرور" required>
            <button type="submit">التسجيل</button>
        </form>

        <a href="login.php">لديك حساب بالفعل؟ تسجيل الدخول</a>
    </div>
</body>
</html>
