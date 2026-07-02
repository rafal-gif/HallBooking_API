<?php
session_start();
header("Content-Type: application/json; charset=utf-8");
require "config.php";

$data = json_decode(file_get_contents("php://input"), true);
$username = trim($data["username"] ?? "");
$password = $data["password"] ?? "";

if ($username === "" || $password === "") {
    echo json_encode(["success" => false, "message" => "الرجاء إدخال اسم المستخدم وكلمة المرور"]);
    exit;
}

$stmt = $pdo->prepare("SELECT id, username, password FROM users WHERE username = ?");
$stmt->execute([$username]);
$user = $stmt->fetch();

if ($user && password_verify($password, $user["password"])) {
    $_SESSION["user_id"] = $user["id"];
    $_SESSION["username"] = $user["username"];
    echo json_encode(["success" => true]);
} else {
    echo json_encode(["success" => false, "message" => "اسم المستخدم أو كلمة المرور غير صحيحة"]);
}
