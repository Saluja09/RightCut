/**
 * RightCut — Auth store using Zustand + Supabase.
 * Falls back to guest mode when Supabase is not configured.
 */
import { create } from 'zustand'
import { supabase, supabaseEnabled } from '../lib/supabase'

const useAuthStore = create((set, get) => ({
  user: null,
  isGuest: false,
  loading: true,   // true while checking existing session

  // Check for existing Supabase session on startup
  initAuth: async () => {
    if (!supabaseEnabled) {
      set({ loading: false, isGuest: true })
      return
    }
    const { data: { session } } = await supabase.auth.getSession()
    set({
      user: session?.user ?? null,
      isGuest: !session?.user,
      loading: false,
    })
    // Listen for auth changes
    supabase.auth.onAuthStateChange((_event, session) => {
      set({ user: session?.user ?? null, isGuest: !session?.user })
    })
  },

  signUp: async (email, password) => {
    const { data, error } = await supabase.auth.signUp({ email, password })
    if (error) throw error
    set({ user: data.user })
    return data.user
  },

  signIn: async (email, password) => {
    const { data, error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) throw error
    set({ user: data.user, isGuest: false })
    return data.user
  },

  signInAsGuest: async () => {
    if (!supabaseEnabled) {
      set({ isGuest: true, user: null })
      return
    }
    const { data, error } = await supabase.auth.signInAnonymously()
    if (error) {
      // fallback to local guest if anon auth not enabled
      set({ isGuest: true, user: null })
      return
    }
    set({ user: data.user, isGuest: true })
  },

  signOut: async () => {
    if (supabaseEnabled) await supabase.auth?.signOut()
    set({ user: null, isGuest: false })
  },

  isAuthenticated: () => {
    const { user, isGuest } = get()
    return !!(user || isGuest)
  },
}))

export default useAuthStore
