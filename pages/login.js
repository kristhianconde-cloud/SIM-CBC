/* ==========================================================================
   SIM-CBC - PÁGINA DE LOGIN Y REGISTRO (RBAC + Firebase Auth)
   ========================================================================== */

import { db } from "../services/db.js?v=20260717.2";
import {
  loginWithFirebase,
  registerWithFirebase,
  sendPasswordResetEmailFirebase,
  determineUserRole
} from "../services/firebase-auth.js?v=20261011.1";

export function renderLoginPage(containerId, onLoginSuccess) {
  const container = document.getElementById(containerId);
  if (!container) return;

  container.innerHTML = `
    <div class="login-container">
      <div class="login-background-glow"></div>
      <div class="login-background-glow-2"></div>
      
      <div class="login-card" id="auth-card">
        <!-- VISTA DE LOGIN -->
        <div id="login-view">
          <div class="login-header">
            <div class="login-logo">
              <i data-lucide="shield-check" style="width: 50px; height: 50px;"></i>
            </div>
            <h1 class="login-title">SIM-CBC</h1>
            <p class="login-subtitle">Sistema de Información y Monitoreo de las Condiciones Básicas de Calidad</p>
          </div>

          <div id="login-message" class="alert-message" style="display: none; margin-bottom: 1rem; padding: 0.75rem; border-radius: var(--radius-sm); font-size: 0.85rem;"></div>
          
          <form id="login-form" style="display: flex; flex-direction: column; gap: 1rem;">
            <div class="form-group">
              <label for="login-email">Correo Institucional</label>
              <input type="email" id="login-email" class="input-control" placeholder="usuario@institucion.edu.pe" required autocomplete="email">
            </div>
            
            <div class="form-group">
              <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.35rem;">
                <label for="login-password" style="margin-bottom: 0;">Contraseña</label>
                <a href="#" id="forgot-password-link" style="color: var(--accent-light); font-size: 0.75rem; text-decoration: none;">¿Olvidó su contraseña?</a>
              </div>
              <input type="password" id="login-password" class="input-control" placeholder="••••••••" required autocomplete="current-password">
            </div>

            <div class="form-group">
              <label for="login-role">Rol para Acceder</label>
              <select id="login-role" class="input-control" style="background-color: var(--bg-dark-900); color: var(--text-100);" required>
                <option value="Administrador">Administrador</option>
                <option value="Usuario" selected>Usuario (Monitoreo y Lectura)</option>
              </select>
            </div>
            
            <button type="submit" id="btn-submit-login" class="btn btn-primary" style="margin-top: 0.5rem; padding: 0.8rem;">
              <span>Iniciar Sesión</span>
              <i data-lucide="arrow-right" style="width: 18px; height: 18px;"></i>
            </button>
          </form>

          <div style="text-align: center; margin-top: 1.5rem;">
            <a href="#" id="go-to-register" style="color: var(--accent-light); font-size: 0.85rem; text-decoration: none; font-weight: 500; transition: color var(--transition-fast);">
              ¿No tienes una cuenta? Regístrate aquí
            </a>
          </div>
        </div>

        <!-- VISTA DE REGISTRO -->
        <div id="register-view" style="display: none;">
          <div class="login-header" style="margin-bottom: 1.5rem;">
            <div class="login-logo" style="margin-bottom: 0.5rem;">
              <i data-lucide="user-plus" style="width: 45px; height: 45px;"></i>
            </div>
            <h1 class="login-title" style="font-size: 1.5rem;">Crear Cuenta</h1>
            <p class="login-subtitle">Registro en el Sistema de Gestión de Calidad (SIM-CBC)</p>
          </div>

          <div id="register-message" class="alert-message" style="display: none; margin-bottom: 1rem; padding: 0.75rem; border-radius: var(--radius-sm); font-size: 0.85rem;"></div>

          <form id="register-form" style="display: flex; flex-direction: column; gap: 0.75rem;">
            <div class="form-group">
              <label for="reg-nombre">Nombre Completo</label>
              <input type="text" id="reg-nombre" class="input-control" placeholder="Ej. Dr. Carlos Mendoza" required>
            </div>

            <div class="form-group">
              <label for="reg-email">Correo Institucional</label>
              <input type="email" id="reg-email" class="input-control" placeholder="usuario@institucion.edu.pe" required autocomplete="email">
            </div>

            <div class="form-group">
              <label for="reg-password">Contraseña</label>
              <input type="password" id="reg-password" class="input-control" placeholder="Mínimo 6 caracteres" required minlength="6" autocomplete="new-password">
            </div>

            <div class="form-group">
              <label for="reg-rol">Tipo de Cuenta Solicitada</label>
              <select id="reg-rol" class="input-control" style="background-color: var(--bg-dark-900); color: var(--text-100);" required>
                <option value="Usuario" selected>Usuario (Monitoreo y Lectura SIM-CBC)</option>
              </select>
            </div>

            <button type="submit" id="btn-submit-register" class="btn btn-primary" style="margin-top: 0.5rem; padding: 0.8rem;">
              <span>Registrarse</span>
              <i data-lucide="user-check" style="width: 18px; height: 18px;"></i>
            </button>
          </form>

          <div style="text-align: center; margin-top: 1rem;">
            <a href="#" id="go-to-login" style="color: var(--accent-light); font-size: 0.85rem; text-decoration: none; font-weight: 500; transition: color var(--transition-fast);">
              ¿Ya tienes una cuenta? Inicia sesión
            </a>
          </div>
        </div>
      </div>
    </div>
  `;

  // Renderizar iconos de Lucide
  if (window.lucide) {
    window.lucide.createIcons();
  }

  // Elementos del DOM
  const loginView = document.getElementById("login-view");
  const registerView = document.getElementById("register-view");
  const goToRegister = document.getElementById("go-to-register");
  const goToLogin = document.getElementById("go-to-login");
  const forgotPasswordLink = document.getElementById("forgot-password-link");

  const loginForm = document.getElementById("login-form");
  const loginMsg = document.getElementById("login-message");
  const btnSubmitLogin = document.getElementById("btn-submit-login");

  const registerForm = document.getElementById("register-form");
  const registerMsg = document.getElementById("register-message");
  const btnSubmitRegister = document.getElementById("btn-submit-register");

  // Mostrar mensaje de alerta
  function showMessage(element, text, type = "error") {
    element.textContent = text;
    element.style.display = "block";
    if (type === "success") {
      element.style.backgroundColor = "var(--color-cumple-bg)";
      element.style.color = "var(--color-cumple)";
      element.style.border = "1px solid rgba(16, 185, 129, 0.2)";
    } else {
      element.style.backgroundColor = "var(--color-nocumple-bg)";
      element.style.color = "var(--color-nocumple)";
      element.style.border = "1px solid rgba(239, 68, 68, 0.2)";
    }
  }

  // Navegación entre vistas
  goToRegister.onclick = (e) => {
    e.preventDefault();
    loginView.style.display = "none";
    registerView.style.display = "block";
    registerForm.reset();
    registerMsg.style.display = "none";
    if (window.lucide) window.lucide.createIcons();
  };

  goToLogin.onclick = (e) => {
    e.preventDefault();
    registerView.style.display = "none";
    loginView.style.display = "block";
    loginForm.reset();
    loginMsg.style.display = "none";
    if (window.lucide) window.lucide.createIcons();
  };

  // Restablecer contraseña con Firebase
  if (forgotPasswordLink) {
    forgotPasswordLink.onclick = async (e) => {
      e.preventDefault();
      const email = document.getElementById("login-email").value.trim();
      if (!email) {
        showMessage(loginMsg, "Por favor ingrese su correo institucional en el campo de correo para enviarle el enlace de recuperación.", "error");
        return;
      }
      try {
        await sendPasswordResetEmailFirebase(email);
        showMessage(loginMsg, `Se ha enviado un correo a "${email}" con las instrucciones para restablecer su contraseña.`, "success");
      } catch (err) {
        showMessage(loginMsg, err.message, "error");
      }
    };
  }

  // Procesar Login Form
  loginForm.onsubmit = async (e) => {
    e.preventDefault();
    loginMsg.style.display = "none";

    const email = document.getElementById("login-email").value.trim();
    const password = document.getElementById("login-password").value;
    const selectedRole = document.getElementById("login-role").value;

    const originalBtnText = btnSubmitLogin.innerHTML;
    btnSubmitLogin.disabled = true;
    btnSubmitLogin.innerHTML = `<span>Autenticando...</span>`;

    try {
      let session = null;

      // Intentar autenticación con Firebase Auth
      try {
        session = await loginWithFirebase(email, password);
      } catch (fbErr) {
        // Si Firebase falló porque no hay internet o credencial local, intentar fallback con db local
        console.warn("[SIM-CBC] Firebase Auth falló, verificando base de datos local:", fbErr.message);
        try {
          session = await db.authenticateUser(email, password);
        } catch (dbErr) {
          // Si ambos fallaron, lanzar el mensaje de Firebase (o el de DB)
          throw fbErr;
        }
      }

      // Validar coherencia con el rol seleccionado
      if (selectedRole === "Administrador" && session.rol !== "Administrador") {
        throw new Error("Esta cuenta no tiene permisos de Administrador. Por favor ingrese seleccionando el rol 'Usuario'.");
      }

      // Si el rol seleccionado es Usuario y el usuario es Administrador, se le permite el acceso en modo monitoreo
      if (selectedRole === "Usuario" && session.rol === "Administrador") {
        session.rolVisual = "Usuario";
      }

      localStorage.setItem("sigeca_current_user", JSON.stringify(session));
      onLoginSuccess(session);
    } catch (err) {
      showMessage(loginMsg, err.message, "error");
    } finally {
      btnSubmitLogin.disabled = false;
      btnSubmitLogin.innerHTML = originalBtnText;
      if (window.lucide) window.lucide.createIcons();
    }
  };

  // Procesar Registro Form
  registerForm.onsubmit = async (e) => {
    e.preventDefault();
    registerMsg.style.display = "none";

    const nombre = document.getElementById("reg-nombre").value.trim();
    const email = document.getElementById("reg-email").value.trim();
    const password = document.getElementById("reg-password").value;

    // Validar correo
    const emailLower = email.toLowerCase();
    const isAllowedEmail = emailLower.endsWith("@upt.pe") ||
                           emailLower.endsWith("@upt.edu.pe") ||
                           emailLower.endsWith("@institucion.edu.pe");

    if (!isAllowedEmail) {
      showMessage(registerMsg, "El registro está restringido a correos institucionales (*@upt.pe, *@upt.edu.pe, *@institucion.edu.pe).", "error");
      return;
    }

    if (password.length < 6) {
      showMessage(registerMsg, "La contraseña debe contener al menos 6 caracteres.", "error");
      return;
    }

    const originalBtnText = btnSubmitRegister.innerHTML;
    btnSubmitRegister.disabled = true;
    btnSubmitRegister.innerHTML = `<span>Creando cuenta en Firebase...</span>`;

    try {
      // 1. Crear en Firebase Auth
      await registerWithFirebase(email, password, nombre);

      // 2. Registrar en db local para compatibilidad total
      try {
        await db.registerUser({
          nombre,
          email,
          password,
          rol: "Usuario",
          subtipo: "N/A",
          areas: ["Licenciamiento"]
        });
      } catch (localErr) {
        console.warn("[SIM-CBC] Registro local redundante:", localErr);
      }

      showMessage(
        registerMsg,
        "¡Cuenta creada exitosamente en Firebase! Ya puede iniciar sesión con su correo y contraseña.",
        "success"
      );

      // Redirigir a login después de 2 segundos
      setTimeout(() => {
        registerView.style.display = "none";
        loginView.style.display = "block";
        loginForm.reset();
        loginMsg.style.display = "none";
        document.getElementById("login-email").value = email;
        document.getElementById("login-password").value = "";
        document.getElementById("login-role").value = "Usuario";
        if (window.lucide) window.lucide.createIcons();
      }, 2000);

    } catch (err) {
      showMessage(registerMsg, err.message, "error");
    } finally {
      btnSubmitRegister.disabled = false;
      btnSubmitRegister.innerHTML = originalBtnText;
      if (window.lucide) window.lucide.createIcons();
    }
  };
}
