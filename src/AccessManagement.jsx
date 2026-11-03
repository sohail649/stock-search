import { useEffect, useState } from "react";
import { collection, doc, getDocs, query, serverTimestamp, setDoc, updateDoc } from "firebase/firestore";
import { db } from "./firebase.js";
import "./Auth.css";

export default function AccessManagement({ user, isSuperAdmin, onHome }) {
  const [requests, setRequests] = useState([]);
  const [profiles, setProfiles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [working, setWorking] = useState("");
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");

  const refresh = async () => {
    setLoading(true);
    setError("");
    try {
      const [requestSnap, profileSnap] = await Promise.all([
        getDocs(query(collection(db, "accessRequests"))),
        getDocs(query(collection(db, "userAccess"))),
      ]);
      setRequests(requestSnap.docs.map((d) => ({ id: d.id, ...d.data() })));
      setProfiles(profileSnap.docs.map((d) => ({ id: d.id, ...d.data() })));
    } catch (err) {
      setError("Users load nahi ho sake: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { refresh(); }, []);

  const approveOrBlock = async (request, status) => {
    setWorking(request.id);
    setError("");
    setMessage("");
    try {
      const now = serverTimestamp();
      const role = "user";
      await setDoc(doc(db, "userAccess", request.id), {
        uid: request.id,
        email: request.email || "",
        displayName: request.displayName || "",
        status,
        role,
        approvedBy: user.uid,
        approvedByEmail: user.email || "",
        approvedAt: now,
        updatedAt: now,
      }, { merge: true });
      await setDoc(doc(db, "accessRequests", request.id), {
        uid: request.id,
        email: request.email || "",
        displayName: request.displayName || "",
        status,
        approvedBy: user.uid,
        approvedByEmail: user.email || "",
        approvedAt: now,
        updatedAt: now,
      }, { merge: true });
      setMessage(status === "approved" ? "User approved." : "User blocked.");
      await refresh();
    } catch (err) {
      setError("Action failed: " + err.message);
    } finally {
      setWorking("");
    }
  };

  const changeAdminRole = async (profile, role) => {
    if (!isSuperAdmin) return;
    setWorking(profile.id);
    setError("");
    try {
      await updateDoc(doc(db, "userAccess", profile.id), {
        role,
        roleChangedBy: user.uid,
        roleChangedByEmail: user.email || "",
        roleChangedAt: serverTimestamp(),
      });
      setMessage(role === "admin" ? "Admin authority granted." : "Admin authority removed.");
      await refresh();
    } catch (err) {
      setError("Role update failed: " + err.message);
    } finally {
      setWorking("");
    }
  };

  const dateText = (value) => value?.toDate ? value.toDate().toLocaleString() : "—";

  return (
    <main className="access-page">
      <header className="access-header">
        <div><h1>User Approval & Access</h1><p>{isSuperAdmin ? "Super Admin control panel" : "Admin approval panel"}</p></div>
        <button onClick={onHome}>⌂ Home</button>
      </header>
      {error && <div className="access-alert error">{error}</div>}
      {message && <div className="access-alert success">{message}</div>}
      <div className="access-section-head"><h2>Access Requests</h2><button onClick={refresh} disabled={loading}>Refresh</button></div>
      {loading ? <p>Loading users...</p> : (
        <div className="access-table-wrap"><table className="access-table">
          <thead><tr><th>User</th><th>Status</th><th>Requested</th><th>Approved / Updated By</th><th>Action</th></tr></thead>
          <tbody>{requests.length ? requests.map((item) => (
            <tr key={item.id}>
              <td><strong>{item.displayName || "—"}</strong><small>{item.email || item.id}</small></td>
              <td><span className={`access-pill ${item.status}`}>{item.status || "pending"}</span></td>
              <td>{dateText(item.requestedAt)}</td>
              <td>{item.approvedByEmail || item.approvedBy || "—"}<small>{dateText(item.approvedAt)}</small></td>
              <td className="access-actions">
                {item.status !== "approved" && <button disabled={!!working} onClick={() => approveOrBlock(item, "approved")}>{working === item.id ? "Saving..." : "Approve"}</button>}
                {item.status !== "blocked" && <button className="danger" disabled={!!working} onClick={() => approveOrBlock(item, "blocked")}>Block</button>}
              </td>
            </tr>
          )) : <tr><td colSpan="5">No access requests found.</td></tr>}</tbody>
        </table></div>
      )}

      <div className="access-section-head"><h2>{isSuperAdmin ? "Admins & Approved Users" : "Approved Users"}</h2></div>
      <div className="access-table-wrap"><table className="access-table">
        <thead><tr><th>User</th><th>Role</th><th>Status</th><th>Approved By</th><th>Approved At</th>{isSuperAdmin && <th>Admin Authority</th>}</tr></thead>
        <tbody>{profiles.map((profile) => (
          <tr key={profile.id}>
            <td><strong>{profile.displayName || "—"}</strong><small>{profile.email || profile.id}</small></td>
            <td>{profile.role || "user"}</td>
            <td><span className={`access-pill ${profile.status}`}>{profile.status}</span></td>
            <td>{profile.approvedByEmail || profile.approvedBy || "Super Admin"}</td>
            <td>{dateText(profile.approvedAt)}</td>
            {isSuperAdmin && <td>{profile.role === "admin" ? <button className="danger" disabled={!!working} onClick={() => changeAdminRole(profile, "user")}>Remove Admin</button> : profile.status === "approved" ? <button disabled={!!working} onClick={() => changeAdminRole(profile, "admin")}>Make Admin</button> : "—"}</td>}
          </tr>
        ))}{!profiles.length && <tr><td colSpan={isSuperAdmin ? "6" : "5"}>No approved profiles found.</td></tr>}</tbody>
      </table></div>
      <p className="access-note">Approval history is stored with the user record, including the approving admin and timestamp. Super Admin can grant or remove Admin authority.</p>
    </main>
  );
}
