// Configuración de tu aplicación web de Firebase
const firebaseConfig = {
    apiKey: "AIzaSyC3rvJBFr86-_rPdrPjZElSvSYdLZPgXQ4",
    authDomain: "centros-incidencias.firebaseapp.com",
    projectId: "centros-incidencias",
    storageBucket: "centros-incidencias.firebasestorage.app",
    messagingSenderId: "468757106886",
    appId: "1:468757106886:web:caca683594754c3fb587c7"
};

// Inicializar Firebase (usando la versión de compatibilidad para evitar problemas de módulos)
firebase.initializeApp(firebaseConfig);

// Inicializar Cloud Firestore y obtener una referencia al servicio
const db = firebase.firestore();

// Inicializar Cloud Storage
const storage = typeof firebase.storage === 'function' ? firebase.storage() : null;
