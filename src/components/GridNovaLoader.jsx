import { useId, useRef } from 'react'

/**
 * Animated GridNova T-mark loader.
 * fullScreen is used while the lazy editor chunk is loading; the embedded
 * variant fills the workbook workspace while the saved sheet is being read.
 */
export default function GridNovaLoader({ label = 'Loading GridNova', fullScreen = false }) {
  const generatedId = useId().replace(/:/g, '')
  const tiltRef = useRef(null)
  const prefix = 'gridnova-loader-' + generatedId
  const ids = {
    gradient: prefix + '-logo-gradient',
    sheen: prefix + '-sheen-gradient',
    main: prefix + '-main-t',
    accent: prefix + '-accent',
    clip: prefix + '-clip',
  }

  function handlePointerMove(event) {
    if (event.pointerType === 'touch' || !tiltRef.current) return
    const bounds = event.currentTarget.getBoundingClientRect()
    const x = (event.clientX - bounds.left) / bounds.width - 0.5
    const y = (event.clientY - bounds.top) / bounds.height - 0.5
    tiltRef.current.style.setProperty('--rotate-y', (x * 20) + 'deg')
    tiltRef.current.style.setProperty('--rotate-x', (-y * 20) + 'deg')
  }

  function resetTilt() {
    if (!tiltRef.current) return
    tiltRef.current.style.setProperty('--rotate-y', '0deg')
    tiltRef.current.style.setProperty('--rotate-x', '0deg')
  }

  return (
    <div
      className={'gridnova-loader-screen ' + (fullScreen ? 'is-fullscreen' : 'is-embedded')}
      role="status"
      aria-live="polite"
      aria-label={label}
    >
      <div
        className="gridnova-loader-stage"
        onPointerMove={handlePointerMove}
        onPointerLeave={resetTilt}
      >
        <div className="gridnova-loader-tilt" ref={tiltRef}>
          <div className="gridnova-loader-logo-wrap">
            <svg
              className="gridnova-loader-logo"
              viewBox="0 0 220 200"
              xmlns="http://www.w3.org/2000/svg"
              aria-hidden="true"
              focusable="false"
            >
              <defs>
                <linearGradient
                  id={ids.gradient}
                  x1="15%"
                  y1="90%"
                  x2="92%"
                  y2="12%"
                  gradientUnits="userSpaceOnUse"
                >
                  <stop offset="0%" stopColor="var(--loader-grad-1)" />
                  <stop offset="32%" stopColor="var(--loader-grad-2)" />
                  <stop offset="68%" stopColor="var(--loader-grad-3)" />
                  <stop offset="100%" stopColor="var(--loader-grad-4)" />
                </linearGradient>
                <linearGradient id={ids.sheen} x1="0%" y1="0%" x2="100%" y2="0%">
                  <stop offset="0%" stopColor="rgba(255,255,255,0)" />
                  <stop offset="50%" stopColor="rgba(255,255,255,0.45)" />
                  <stop offset="100%" stopColor="rgba(255,255,255,0)" />
                </linearGradient>
                <path
                  id={ids.main}
                  d="M 68 32 L 154 32 L 99 170 L 99 63 L 83 63 L 83 80 L 47 80 Z"
                />
                <path
                  id={ids.accent}
                  d="M 164 32 L 189 32 L 176 64 L 151 64 Z"
                />
                <clipPath id={ids.clip}>
                  <use href={'#' + ids.main} />
                  <use href={'#' + ids.accent} />
                </clipPath>
              </defs>

              <g className="gridnova-loader-main-piece">
                <use href={'#' + ids.main} fill={'url(#' + ids.gradient + ')'} />
              </g>
              <g className="gridnova-loader-accent-piece">
                <use href={'#' + ids.accent} fill={'url(#' + ids.gradient + ')'} />
              </g>
              <g clipPath={'url(#' + ids.clip + ')'} pointerEvents="none">
                <rect
                  className="gridnova-loader-sheen"
                  x="0"
                  y="0"
                  width="80"
                  height="200"
                  fill={'url(#' + ids.sheen + ')'}
                />
              </g>
            </svg>
          </div>
        </div>
      </div>
      <span className="gridnova-loader-caption">{label}</span>
    </div>
  )
}
