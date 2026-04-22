/**
 * RightCut — Login page.
 * Email/password sign-in/sign-up + "Continue as guest".
 */
import { useState } from 'react'
import { AlertCircle, UserCircle2, BarChart3 } from 'lucide-react'
import useAuthStore from '../stores/authStore'

export default function LoginPage() {
  const [mode, setMode]       = useState('signin')  // 'signin' | 'signup'
  const [email, setEmail]     = useState('')
  const [password, setPassword] = useState('')
  const [error, setError]     = useState('')
  const [loading, setLoading] = useState(false)

  const { signIn, signUp, signInAsGuest } = useAuthStore()

  const handleSubmit = async (e) => {
    e.preventDefault()
    if (!email || !password) { setError('Enter email and password.'); return }
    setError('')
    setLoading(true)
    try {
      if (mode === 'signup') {
        await signUp(email, password)
      } else {
        await signIn(email, password)
      }
    } catch (err) {
      setError(err.message || 'Authentication failed.')
    } finally {
      setLoading(false)
    }
  }

  const handleGuest = async () => {
    setLoading(true)
    await signInAsGuest()
    setLoading(false)
  }

  return (
    <div className="login-page">
      <div className="login-card">
        <div className="login-header">
          <div className="login-logo"><BarChart3 size={24} /></div>
          <h1 className="login-title">RightCut</h1>
          <p className="login-subtitle">AI Spreadsheet Agent for Private Markets</p>
        </div>

        <form className="login-form" onSubmit={handleSubmit}>
          {error && <div className="login-error"><AlertCircle size={14} />{error}</div>}

          <div className="form-field">
            <label className="form-label">Email</label>
            <input
              className="form-input"
              type="email"
              placeholder="you@fund.com"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              autoFocus
              disabled={loading}
            />
          </div>

          <div className="form-field">
            <label className="form-label">Password</label>
            <input
              className="form-input"
              type="password"
              placeholder={mode === 'signup' ? 'Min. 8 characters' : '••••••••'}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              disabled={loading}
            />
          </div>

          <button className="login-btn" type="submit" disabled={loading}>
            {loading ? 'Please wait…' : mode === 'signin' ? 'Sign in' : 'Create account'}
          </button>

          <div className="login-divider">or</div>

          <button
            className="guest-btn"
            type="button"
            onClick={handleGuest}
            disabled={loading}
          >
            <UserCircle2 size={15} />
            Continue as guest
          </button>
        </form>

        <div className="login-toggle">
          {mode === 'signin' ? (
            <>No account? <a onClick={() => { setMode('signup'); setError('') }}>Sign up free</a></>
          ) : (
            <>Already have one? <a onClick={() => { setMode('signin'); setError('') }}>Sign in</a></>
          )}
        </div>
      </div>
    </div>
  )
}
