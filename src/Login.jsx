import { useState } from "react";
import { Snowflake } from "lucide-react";
import { supabase } from "./supabaseClient.js";

export default function Login() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError(""); setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (error) setError(error.message);
  };

  return (
    <div style={{
      minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
      background: "#0B1420", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
    }}>
      <form onSubmit={submit} style={{
        background: "#132234", border: "1px solid #1B2C3E", borderRadius: 13, padding: 28, width: 320,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
          <div style={{
            width: 30, height: 30, borderRadius: 8, background: "linear-gradient(180deg,#6FCBEA,#3E6E85)",
            display: "flex", alignItems: "center", justifyContent: "center",
          }}><Snowflake size={16} color="#08202C" strokeWidth={2.5} /></div>
          <div style={{ color: "#E9F3F7", fontWeight: 700, fontSize: 15, letterSpacing: "0.02em" }}>Alliance Command</div>
        </div>
        <label style={{ fontSize: 11, color: "#8397AA", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em" }}>Email</label>
        <input type="email" required value={email} onChange={(e) => setEmail(e.target.value)}
          style={{ width: "100%", background: "#0F1C2C", border: "1px solid #24384C", color: "#E9F3F7", borderRadius: 7, padding: "8px 10px", fontSize: 13, margin: "5px 0 12px", boxSizing: "border-box" }} />
        <label style={{ fontSize: 11, color: "#8397AA", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em" }}>Password</label>
        <input type="password" required value={password} onChange={(e) => setPassword(e.target.value)}
          style={{ width: "100%", background: "#0F1C2C", border: "1px solid #24384C", color: "#E9F3F7", borderRadius: 7, padding: "8px 10px", fontSize: 13, margin: "5px 0 16px", boxSizing: "border-box" }} />
        {error && <div style={{ color: "#E2604F", fontSize: 12, marginBottom: 12 }}>{error}</div>}
        <button type="submit" disabled={loading} style={{
          width: "100%", background: "#6FCBEA", color: "#08202C", border: "none", borderRadius: 8,
          padding: "10px 0", fontWeight: 700, fontSize: 13, cursor: "pointer", opacity: loading ? 0.6 : 1,
        }}>{loading ? "Signing in…" : "Sign in"}</button>
        <div style={{ fontSize: 11, color: "#5C7086", marginTop: 14, textAlign: "center" }}>
          No self-signup — accounts are created by an admin in the Supabase dashboard.
        </div>
      </form>
    </div>
  );
}
