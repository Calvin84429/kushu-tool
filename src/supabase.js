import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = 'https://zctncvgmwokyvxqovccy.supabase.co'
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpjdG5jdmdtd29reXZ4cW92Y2N5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzM3OTg0MTUsImV4cCI6MjA4OTM3NDQxNX0.fSIfO0iGGzR64miSduyw3D1E1zS_rr_gdyZh_bZkvMk'

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY)

// ── DB helpers ──────────────────────────────────────────────────────────────

// Clients
export async function dbGetClients() {
  const { data, error } = await supabase.from('clients').select('*')
  if (error) { console.error('getClients:', error); return {} }
  const map = {}
  data.forEach(row => {
    map[row.id] = {
      id: row.id, type: row.type, name: row.name,
      account: row.account || '', note: row.note || '',
      weight: row.weight || 'mid', grade: row.grade || 'auto',
      stage: row.stage || '建立信任',
      messages: row.messages || [],
      createdAt: row.created_at || '',
    }
  })
  return map
}

export async function dbUpsertClient(client) {
  const { error } = await supabase.from('clients').upsert({
    id: client.id, type: client.type, name: client.name,
    account: client.account || '', note: client.note || '',
    weight: client.weight || 'mid', grade: client.grade || 'auto',
    stage: client.stage || '建立信任',
    messages: client.messages || [],
    created_at: client.createdAt || new Date().toLocaleDateString('zh-TW'),
  })
  if (error) console.error('upsertClient:', error)
}

export async function dbDeleteClient(id) {
  const { error } = await supabase.from('clients').delete().eq('id', id)
  if (error) console.error('deleteClient:', error)
}

// Client meta (order, pinnedIds)
export async function dbGetMeta(key) {
  const { data, error } = await supabase.from('client_meta').select('value').eq('key', key).single()
  if (error) return null
  return data?.value
}

export async function dbSetMeta(key, value) {
  const { error } = await supabase.from('client_meta').upsert({ key, value })
  if (error) console.error('setMeta:', error)
}

// Shortcuts
export async function dbGetShortcuts() {
  const { data, error } = await supabase.from('shortcuts').select('data').eq('id', 'main').single()
  if (error) return null
  return data?.data ?? null
}

export async function dbSetShortcuts(sc) {
  const { error } = await supabase.from('shortcuts').upsert({ id: 'main', data: sc })
  if (error) console.error('setShortcuts:', error)
}

// Trash
export async function dbGetTrash() {
  const { data, error } = await supabase.from('trash').select('data').eq('id', 'main').single()
  if (error) return []
  return data?.data || []
}

export async function dbSetTrash(trash) {
  const { error } = await supabase.from('trash').upsert({ id: 'main', data: trash })
  if (error) console.error('setTrash:', error)
}
