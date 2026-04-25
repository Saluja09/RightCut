/**
 * RoleSelectModal — shown at the start of each new session.
 * Asks the user whether they are a financial professional or a general user,
 * then posts the choice to the backend /configure endpoint.
 */
import { useState } from 'react'
import { Briefcase, GraduationCap } from 'lucide-react'
import { apiUrl } from '../utils/api'

const ROLES = [
  {
    id: 'finance',
    icon: Briefcase,
    title: 'Financial Professional',
    subtitle: 'Deal teams, investors, analysts',
    description:
      'Build institutional-grade models — DCF, LBO, comps tables, deal sheets — with professional formatting and private markets domain knowledge.',
  },
  {
    id: 'general',
    icon: GraduationCap,
    title: 'General User',
    subtitle: 'Students, researchers, anyone with data',
    description:
      'Create clean spreadsheets, data summaries, budgets, reports, or anything else — no finance background required.',
  },
]

export default function RoleSelectModal({ sessionId, onConfirm }) {
  const [selected, setSelected] = useState(null)
  const [loading, setLoading]   = useState(false)

  const handleConfirm = async () => {
    if (!selected || loading) return
    setLoading(true)
    try {
      await fetch(apiUrl(`/configure/${sessionId}`), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ role: selected }),
      })
    } catch (e) {
      console.warn('Role configure failed:', e)
    } finally {
      setLoading(false)
      onConfirm(selected)
    }
  }

  return (
    <div className="role-modal-overlay">
      <div className="role-modal">
        <div className="role-modal-header">
          <div className="role-modal-logo">RC</div>
          <h2 className="role-modal-title">How can we help you today?</h2>
          <p className="role-modal-subtitle">
            Select your context so RightCut can tailor its expertise to your needs.
          </p>
        </div>

        <div className="role-modal-options">
          {ROLES.map((role) => {
            const Icon = role.icon
            const isSelected = selected === role.id
            return (
              <button
                key={role.id}
                className={`role-option${isSelected ? ' role-option--selected' : ''}`}
                onClick={() => setSelected(role.id)}
                type="button"
              >
                <div className="role-option-icon">
                  <Icon size={22} />
                </div>
                <div className="role-option-body">
                  <div className="role-option-title">{role.title}</div>
                  <div className="role-option-subtitle">{role.subtitle}</div>
                  <div className="role-option-desc">{role.description}</div>
                </div>
                <div className={`role-option-check${isSelected ? ' role-option-check--visible' : ''}`}>
                  <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                    <circle cx="8" cy="8" r="8" fill="#217346"/>
                    <path d="M4.5 8l2.5 2.5 4.5-5" stroke="#fff" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </div>
              </button>
            )
          })}
        </div>

        <div className="role-modal-footer">
          <button
            className="role-modal-btn"
            disabled={!selected || loading}
            onClick={handleConfirm}
            type="button"
          >
            {loading ? 'Starting session…' : 'Start Session'}
          </button>
          <p className="role-modal-note">
            You can change this at the start of any new session.
          </p>
        </div>
      </div>
    </div>
  )
}
