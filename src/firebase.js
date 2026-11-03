import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth, GoogleAuthProvider } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyC8U7MeqRID2H1PmFV4dF-O5Q0drUdwPmo",
  authDomain: "stock-search-cloud.firebaseapp.com",
  projectId: "stock-search-cloud",
  storageBucket: "stock-search-cloud.firebasestorage.app",
  messagingSenderId: "980118380665",
  appId: "1:980118380665:web:6ccc33df7404fb58b834a8",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);
export const googleProvider = new GoogleAuthProvider();
googleProvider.setCustomParameters({ prompt: "select_account" });
