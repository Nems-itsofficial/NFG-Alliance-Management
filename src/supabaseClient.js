import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !anonKey) {
  // Surfaces a clear error in the browser console instead of a silent failure
  // if the environment variables weren't set at deploy time.
  console.error(
    "Missing Supabase config. Set VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY " +
    "(see .env.example and README.md)."
  );
}

export const supabase = createClient(url || "", anonKey || "");
