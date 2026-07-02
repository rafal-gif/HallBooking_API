<?php
session_start();
header("Content-Type: application/json; charset=utf-8");
require "config.php";

$data = json_decode(file_get_contents("php://input"), true);
$username = trim($data["username"] ?? "");
$password = $data["password"] ?? "";
$confirm_password = $data["confirm_password"] ?? "";

if ($username === "" || $password === "" || $confirm_password === "") {
    echo json_encode(["success" => false, "message" => "الرجاء تعبئة جميع الحقول"]);
    exit;
}

if ($password !== $confirm_password) {
    echo json_encode(["success" => false, "message" => "كلمة المرور وتأكيدها غير متطابقين"]);
    exit;
}

$stmt = $pdo->prepare("SELECT id FROM users WHERE username = ?");
$stmt->execute([$username]);

if ($stmt->fetch()) {
    echo json_encode(["success" => false, "message" => "اسم المستخدم موجود مسبقًا"]);
    exit;
}

$hashed_password = password_hash($password, PASSWORD_DEFAULT);
$stmt = $pdo->prepare("INSERT INTO users (username, password) VALUES (?, ?)");
$stmt->execute([$username, $hashed_password]);

echo json_encode(["success" => true, "message" => "تم التسجيل بنجاح"]);
