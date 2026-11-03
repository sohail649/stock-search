import { useEffect, useState } from "react";
import {
  onAuthStateChanged,
  signInWithPopup,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  updateProfile,
  signOut,
} from "firebase/auth";
import {
  doc,
  getDoc,
  setDoc,
  serverTimestamp,
  onSnapshot,
} from "firebase/firestore";
import { auth, db, googleProvider } from "./firebase.js";
import "./Auth.css";

// The first Super Admin account is fixed to the UID already verified in Firebase.
const SUPER_ADMIN_UID = "TDXw1GNV7UUtR6cu8JZ1zzgIQIK2";

export default function AuthGate({ children }) {
  const [user, setUser] = useState(null);
  const [accessProfile, setAccessProfile] = useState(null);
  const [isSuperAdmin, setIsSuperAdmin] = useState(false);
  const [checking, setChecking] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [mode, setMode] = useState("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [displayName, setDisplayName] = useState("");

  useEffect(() => {
    let profileUnsubscribe = null;
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      if (profileUnsubscribe) {
        profileUnsubscribe();
        profileUnsubscribe = null;
      }
      setUser(currentUser);
      setError("");
      if (!currentUser) {
        setAccessProfile(null);
        setIsSuperAdmin(false);
        setChecking(false);
        return;
      }

      setChecking(true);
      try {
        if (currentUser.uid === SUPER_ADMIN_UID) {
          setIsSuperAdmin(true);
          setAccessProfile({ status: "approved", role: "superAdmin" });
          setChecking(false);
          return;
        }

        setIsSuperAdmin(false);
        const profileRef = doc(db, "userAccess", currentUser.uid);
        const profileSnap = await getDoc(profileRef);
        if (!profileSnap.exists()) {
          const requestRef = doc(db, "accessRequests", currentUser.uid);
          const requestSnap = await getDoc(requestRef);
          if (!requestSnap.exists()) {
            await setDoc(requestRef, {
              uid: currentUser.uid,
              email: currentUser.email || "",
              displayName: currentUser.displayName || "",
              status: "pending",
              requestedAt: serverTimestamp(),
            });
            setAccessProfile({ status: "pending", role: "user" });
          } else {
            setAccessProfile(requestSnap.data());
          }
        } else {
          setAccessProfile(profileSnap.data());
        }

        // Keep the approval state live: an admin approval opens the app without requiring logout.
        profileUnsubscribe = onSnapshot(profileRef, (snap) => {
          setAccessProfile(snap.exists() ? snap.data() : null);
        }, (err) => {
          console.error("Access listener failed", err);
          setError("Access status update nahi ho saka.");
        });
      } catch (err) {
        console.error("Access check failed", err);
        setError("Access check nahi ho saka. Internet ya Firebase Rules check karein.");
        setAccessProfile(null);
      } finally {
        setChecking(false);
      }
    });
    return () => {
      unsubscribe();
      if (profileUnsubscribe) profileUnsubscribe();
    };
  }, []);

  const signInGoogle = async () => {
    setBusy(true);
    setError("");
    try {
      await signInWithPopup(auth, googleProvider);
    } catch (err) {
      setError(err.message || "Google sign-in failed.");
    } finally {
      setBusy(false);
    }
  };

  const submitEmail = async (event) => {
    event.preventDefault();
    setBusy(true);
    setError("");
    try {
      if (mode === "register") {
        const result = await createUserWithEmailAndPassword(auth, email.trim(), password);
        if (displayName.trim()) await updateProfile(result.user, { displayName: displayName.trim() });
      } else {
        await signInWithEmailAndPassword(auth, email.trim(), password);
      }
    } catch (err) {
      setError(err.code === "auth/email-already-in-use"
        ? "Ye email pehle se registered hai. Login karein."
        : err.code === "auth/invalid-credential"
        ? "Email ya password ghalat hai."
        : err.message || "Authentication failed.");
    } finally {
      setBusy(false);
    }
  };

  const signOutUser = async () => {
    await signOut(auth);
    setAccessProfile(null);
    setIsSuperAdmin(false);
  };

  if (checking) {
    return <div className="auth-screen"><div className="auth-card"><h2>Checking access...</h2><p>Please wait.</p></div></div>;
  }

  if (user && (isSuperAdmin || accessProfile?.status === "approved")) {
    return children({ user, accessProfile, isSuperAdmin, signOutUser });
  }

  if (user) {
    const status = accessProfile?.status || "pending";
    return (
      <div className="auth-screen">
        <div className="auth-card">
          <div className="auth-brand">Stock Manager</div>
          <div className="auth-status-icon">{status === "blocked" ? "×" : "…"}</div>
          <h2>{status === "blocked" ? "Access Blocked" : "Approval Pending"}</h2>
          <p>{status === "blocked"
            ? "Aapka access admin ne block kar diya hai. Admin se rabta karein."
            : "Aapki access request admin ko bhej di gayi hai. Approval milne ke baad app use kar sakenge."}</p>
          <div className="auth-user-email">{user.email}</div>
          {error && <div className="auth-error">{error}</div>}
          <button className="auth-primary" onClick={signOutUser}>Sign out</button>
        </div>
      </div>
    );
  }

  return (
    <div className="auth-screen">
      <div className="auth-card">
        <div className="auth-brand">Stock Manager</div>
        <h2>{mode === "register" ? "Request App Access" : "Sign in"}</h2>
        <p>{mode === "register" ? "Account banayein; admin approval ke baad app open hogi." : "App use karne ke liye approved account se login karein."}</p>
        <button className="auth-google" onClick={signInGoogle} disabled={busy}>Continue with Google</button>
        <div className="auth-divider"><span>OR</span></div>
        <form onSubmit={submitEmail} className="auth-form">
          {mode === "register" && <input value={displayName} onChange={(e) => setDisplayName(e.target.value)} placeholder="Your name" autoComplete="name" />}
          <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="Email address" autoComplete="email" required />
          <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Password (minimum 6 characters)" autoComplete={mode === "register" ? "new-password" : "current-password"} minLength={6} required />
          {error && <div className="auth-error">{error}</div>}
          <button className="auth-primary" type="submit" disabled={busy}>{busy ? "Please wait..." : mode === "register" ? "Create account & request approval" : "Login"}</button>
        </form>
        <button className="auth-link" onClick={() => { setMode(mode === "register" ? "login" : "register"); setError(""); }}>
          {mode === "register" ? "Already registered? Login" : "New user? Create account"}
        </button>
      </div>
    </div>
  );
}
