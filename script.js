document.addEventListener("DOMContentLoaded", () => {
    const loginForm = document.getElementById("loginForm");
    const registerForm = document.getElementById("registerForm");
    const usernameDisplay = document.getElementById("usernameDisplay");
    const logoutBtn = document.getElementById("logoutBtn");
    const messageBox = document.getElementById("message");

    function showMessage(text, type) {
        if (!messageBox) return;
        messageBox.textContent = text;
        messageBox.className = "msg " + type;
    }

    if (loginForm) {
        const params = new URLSearchParams(window.location.search);
        if (params.get("registered") === "1") {
            showMessage("تم التسجيل بنجاح، الرجاء تسجيل الدخول", "success");
        }

        loginForm.addEventListener("submit", async (e) => {
            e.preventDefault();
            const username = document.getElementById("username").value.trim();
            const password = document.getElementById("password").value;

            const res = await fetch("login_handler.php", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                credentials: "same-origin",
                body: JSON.stringify({ username, password }),
            });
            const data = await res.json();

            if (data.success) {
                window.location.href = "home.html";
            } else {
                showMessage(data.message, "error");
            }
        });
    }

    if (registerForm) {
        registerForm.addEventListener("submit", async (e) => {
            e.preventDefault();
            const username = document.getElementById("username").value.trim();
            const password = document.getElementById("password").value;
            const confirm_password = document.getElementById("confirm_password").value;

            const res = await fetch("register_handler.php", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                credentials: "same-origin",
                body: JSON.stringify({ username, password, confirm_password }),
            });
            const data = await res.json();

            if (data.success) {
                window.location.href = "login.html?registered=1";
            } else {
                showMessage(data.message, "error");
            }
        });
    }

    if (usernameDisplay) {
        fetch("session_check.php", { credentials: "same-origin" })
            .then((res) => res.json())
            .then((data) => {
                if (!data.loggedIn) {
                    window.location.href = "login.html";
                } else {
                    usernameDisplay.textContent = data.username;
                }
            });
    }

    if (logoutBtn) {
        logoutBtn.addEventListener("click", async (e) => {
            e.preventDefault();
            await fetch("logout_handler.php", { credentials: "same-origin" });
            window.location.href = "login.html";
        });
    }
});
