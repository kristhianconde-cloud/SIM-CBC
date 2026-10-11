/* ==========================================================================
   SIM-CBC - SERVICIO DE AUTENTICACIÓN FIREBASE (Firebase Auth Integration)
   ========================================================================== */

/**
 * Configuración oficial del proyecto Firebase 'sim-cbc'
 */
export const firebaseConfig = {
  apiKey: "AIzaSyD-jRnIBiS8yPolwv77IRc_qSTHDzM3S20",
  authDomain: "sim-cbc.firebaseapp.com",
  projectId: "sim-cbc",
  storageBucket: "sim-cbc.firebasestorage.app",
  messagingSenderId: "272745598225",
  appId: "1:272745598225:web:c27z50d4b71922f20bb640",
  measurementId: "G-29BBIBDSYW"
};

/**
 * Lista de correos institucionales designados como Administrador con permisos totales
 */
export const ADMIN_EMAILS = [
  "licencia_institucional@upt.pe",
  "admin@institucion.edu.pe",
  "admin@upt.pe"
];

let authInstance = null;

/**
 * Inicializa y retorna la instancia de Firebase Auth
 */
export function initFirebaseAuth() {
  if (authInstance) return authInstance;

  if (typeof window !== "undefined" && window.firebase) {
    try {
      if (!window.firebase.apps || !window.firebase.apps.length) {
        window.firebase.initializeApp(firebaseConfig);
      }
      authInstance = window.firebase.auth();
      return authInstance;
    } catch (err) {
      console.error("[SIM-CBC] Error al inicializar Firebase Auth:", err);
      return null;
    }
  }
  return null;
}

/**
 * Determina si el rol del usuario autenticado es 'Administrador' o 'Usuario' (solo lectura/monitoreo)
 * @param {string} email
 * @returns {"Administrador" | "Usuario"}
 */
export function determineUserRole(email) {
  const emailLower = (email || "").toLowerCase().trim();

  // 1. Verificar lista de correos administradores oficiales
  if (ADMIN_EMAILS.includes(emailLower)) {
    return "Administrador";
  }

  // 2. Verificar si en la base local de usuarios tiene rol Administrador
  try {
    const localUsers = JSON.parse(localStorage.getItem("sigeca_usuarios") || "[]");
    const found = localUsers.find(u => (u.email || "").toLowerCase() === emailLower);
    if (found && found.rol === "Administrador") {
      return "Administrador";
    }
  } catch (e) {
    console.warn("[SIM-CBC] No se pudo leer sigeca_usuarios:", e);
  }

  // 3. Por defecto, todo usuario tiene rol 'Usuario' (monitoreo institucional y solo lectura)
  return "Usuario";
}

/**
 * Traduce los códigos de error comunes de Firebase Auth a mensajes amigables en español
 * @param {Error} error
 * @returns {string}
 */
export function formatFirebaseAuthError(error) {
  if (!error) return "Ocurrió un error inesperado al autenticar.";
  const code = error.code || "";

  switch (code) {
    case "auth/invalid-email":
      return "El formato del correo institucional no es válido.";
    case "auth/user-not-found":
      return "No existe ninguna cuenta registrada con este correo en SIM-CBC.";
    case "auth/wrong-password":
    case "auth/invalid-credential":
    case "auth/invalid-login-credentials":
      return "Contraseña incorrecta o credenciales no válidas. Verifique sus datos.";
    case "auth/email-already-in-use":
      return "Este correo institucional ya se encuentra registrado en SIM-CBC.";
    case "auth/weak-password":
      return "La contraseña debe tener al menos 6 caracteres.";
    case "auth/too-many-requests":
      return "Demasiados intentos fallidos. Por seguridad, intente nuevamente en unos minutos.";
    case "auth/network-request-failed":
      return "Error de conexión de red con Firebase. Verifique su acceso a internet.";
    case "auth/operation-not-allowed":
      return "El proveedor de inicio de sesión con correo y contraseña no está habilitado.";
    default:
      return error.message || "Error al procesar la autenticación en el sistema.";
  }
}

/**
 * Inicia sesión con Firebase Auth (Email / Password)
 * @param {string} email
 * @param {string} password
 * @returns {Promise<object>} Objeto de sesión compatible con SIM-CBC
 */
export async function loginWithFirebase(email, password) {
  const auth = initFirebaseAuth();
  if (!auth) {
    throw new Error("El servicio de Firebase Auth no está listo. Recargue la página o revise su conexión.");
  }

  try {
    const userCredential = await auth.signInWithEmailAndPassword(email.trim(), password);
    const fbUser = userCredential.user;
    const role = determineUserRole(fbUser.email);

    const displayName = fbUser.displayName || fbUser.email.split("@")[0];

    // Sincronizar usuario en almacenamiento local para compatibilidad con módulos existentes
    try {
      const localUsers = JSON.parse(localStorage.getItem("sigeca_usuarios") || "[]");
      const idx = localUsers.findIndex(u => (u.email || "").toLowerCase() === fbUser.email.toLowerCase());
      if (idx !== -1) {
        localUsers[idx].nombre = displayName;
        localUsers[idx].rol = role;
        localUsers[idx].status = "active";
        localStorage.setItem("sigeca_usuarios", JSON.stringify(localUsers));
      } else {
        localUsers.push({
          uid: fbUser.uid,
          email: fbUser.email,
          nombre: displayName,
          rol: role,
          subtipo: "N/A",
          areas: ["Licenciamiento"],
          status: "active"
        });
        localStorage.setItem("sigeca_usuarios", JSON.stringify(localUsers));
      }
    } catch (e) {}

    const session = {
      uid: fbUser.uid,
      email: fbUser.email,
      nombre: displayName,
      rol: role,
      subtipo: "N/A",
      areas: ["Licenciamiento"],
      solicitudes: [],
      token: await fbUser.getIdToken()
    };

    return session;
  } catch (err) {
    throw new Error(formatFirebaseAuthError(err));
  }
}

/**
 * Registra un nuevo usuario en Firebase Auth y asigna su nombre
 * @param {string} email
 * @param {string} password
 * @param {string} nombre
 * @returns {Promise<object>}
 */
export async function registerWithFirebase(email, password, nombre) {
  const auth = initFirebaseAuth();
  if (!auth) {
    throw new Error("El servicio de Firebase Auth no está listo. Recargue la página o revise su conexión.");
  }

  try {
    const userCredential = await auth.createUserWithEmailAndPassword(email.trim(), password);
    const fbUser = userCredential.user;

    if (nombre && fbUser.updateProfile) {
      await fbUser.updateProfile({ displayName: nombre.trim() });
    }

    const role = determineUserRole(fbUser.email);
    const resolvedName = nombre ? nombre.trim() : fbUser.email.split("@")[0];

    // Registrar en sigeca_usuarios local
    try {
      const localUsers = JSON.parse(localStorage.getItem("sigeca_usuarios") || "[]");
      const idx = localUsers.findIndex(u => (u.email || "").toLowerCase() === fbUser.email.toLowerCase());
      if (idx === -1) {
        localUsers.push({
          uid: fbUser.uid,
          email: fbUser.email,
          nombre: resolvedName,
          rol: role,
          subtipo: "N/A",
          areas: ["Licenciamiento"],
          status: "active",
          createdAt: new Date().toISOString()
        });
        localStorage.setItem("sigeca_usuarios", JSON.stringify(localUsers));
      }
    } catch (e) {}

    return {
      uid: fbUser.uid,
      email: fbUser.email,
      nombre: resolvedName,
      rol: role
    };
  } catch (err) {
    throw new Error(formatFirebaseAuthError(err));
  }
}

/**
 * Cierra la sesión activa en Firebase Auth
 */
export async function logoutFromFirebase() {
  const auth = initFirebaseAuth();
  if (auth) {
    try {
      await auth.signOut();
    } catch (e) {
      console.warn("[SIM-CBC] Error cerrando sesión en Firebase:", e);
    }
  }
}

/**
 * Envía correo para restablecer la contraseña a través de Firebase
 * @param {string} email
 */
export async function sendPasswordResetEmailFirebase(email) {
  const auth = initFirebaseAuth();
  if (!auth) {
    throw new Error("El servicio de Firebase Auth no está disponible.");
  }
  try {
    await auth.sendPasswordResetEmail(email.trim());
  } catch (err) {
    throw new Error(formatFirebaseAuthError(err));
  }
}
