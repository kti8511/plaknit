const first = (...values) => values.find((value) => value && String(value).trim()) || "";

export default function handler(_req, res) {
  const projectUrl = first(
    process.env.PLAKNIT_SUPABASE_URL,
    process.env.NEXT_PUBLIC_SUPABASE_URL,
    process.env.VITE_SUPABASE_URL,
    process.env.SUPABASE_URL
  ).replace(/\/$/, "");

  const anonKey = first(
    process.env.PLAKNIT_SUPABASE_ANON_KEY,
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY,
    process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY,
    process.env.VITE_SUPABASE_ANON_KEY,
    process.env.SUPABASE_ANON_KEY,
    process.env.SUPABASE_PUBLISHABLE_KEY
  );

  res.setHeader("Cache-Control", "no-store");
  res.status(200).json({
    url: projectUrl,
    rest: projectUrl ? `${projectUrl}/rest/v1` : "",
    anonKey,
    bucket: first(process.env.PLAKNIT_SUPABASE_BUCKET, process.env.SUPABASE_BUCKET, "todo-files"),
  });
}
