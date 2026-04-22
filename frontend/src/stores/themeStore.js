/**
 * RightCut — Theme store. Persists light/dark preference.
 */
import { create } from 'zustand'

const THEME_KEY = 'rightcut_theme'

function getInitialTheme() {
  const saved = localStorage.getItem(THEME_KEY)
  if (saved === 'dark' || saved === 'light') return saved
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme)
  localStorage.setItem(THEME_KEY, theme)
}

const useThemeStore = create((set, get) => ({
  theme: getInitialTheme(),

  initTheme: () => {
    applyTheme(get().theme)
  },

  toggleTheme: () => {
    const next = get().theme === 'light' ? 'dark' : 'light'
    applyTheme(next)
    set({ theme: next })
  },
}))

export default useThemeStore
