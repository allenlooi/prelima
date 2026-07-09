import { supabase } from "./supabaseClient.js";

async function fetchRows(table, userId) {
  const { data, error } = await supabase.from(table).select("id,data").eq("user_id", userId);
  if (error) throw error;
  return (data || []).map(r => ({ ...r.data, id: r.id }));
}

async function syncRows(table, userId, rows) {
  const upserts = rows.map(r => ({ id: r.id, user_id: userId, data: r, updated_at: new Date().toISOString() }));
  if (upserts.length) {
    const { error } = await supabase.from(table).upsert(upserts);
    if (error) throw error;
  }
  const { data: existing, error: selErr } = await supabase.from(table).select("id").eq("user_id", userId);
  if (selErr) throw selErr;
  const keep = new Set(rows.map(r => r.id));
  const toDelete = (existing || []).filter(r => !keep.has(r.id)).map(r => r.id);
  if (toDelete.length) {
    const { error: delErr } = await supabase.from(table).delete().in("id", toDelete);
    if (delErr) throw delErr;
  }
}

export const fetchProjects = (userId) => fetchRows("projects", userId);
export const syncProjects = (userId, projects) => syncRows("projects", userId, projects);

export const fetchQuotes = (userId) => fetchRows("quotes", userId);
export const syncQuotes = (userId, quotes) => syncRows("quotes", userId, quotes);

export const fetchTaskBriefs = (userId) => fetchRows("task_briefs", userId);
export const syncTaskBriefs = (userId, briefs) => syncRows("task_briefs", userId, briefs);

export async function fetchProfile(userId) {
  const { data, error } = await supabase.from("profiles").select("workspace_name").eq("id", userId).maybeSingle();
  if (error) throw error;
  return data ? data.workspace_name : null;
}

export async function ensureProfile(userId, defaultName) {
  const existing = await fetchProfile(userId);
  if (existing !== null) return existing;
  const { error } = await supabase.from("profiles").insert({ id: userId, workspace_name: defaultName });
  if (error) throw error;
  return defaultName;
}

export async function saveWorkspaceName(userId, name) {
  const { error } = await supabase.from("profiles").upsert({ id: userId, workspace_name: name, updated_at: new Date().toISOString() });
  if (error) throw error;
}
