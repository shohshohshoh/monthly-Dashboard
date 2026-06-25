import { useState, useEffect } from 'react'
import './App.css'

const API = import.meta.env.VITE_API_URL || 'http://localhost:8000'
const NO_BACKEND = import.meta.env.PROD && !import.meta.env.VITE_API_URL

function validate(value) {
  if (!/^\d{4}\/\d{1,2}$/.test(value)) return '形式は yyyy/m で入力してください（例: 2026/5）'
  const [, m] = value.split('/').map(Number)
  if (m < 1 || m > 12) return '月は 1〜12 で入力してください'
  return ''
}

function parseYM(yearMonth) {
  const [y, m] = yearMonth.split('/').map(Number)
  return { year: y, month: m }
}

function Modal({ title, message, onConfirm, onCancel, confirmLabel = 'はい', cancelLabel = 'キャンセル', danger = false }) {
  return (
    <div className="modal-overlay">
      <div className="modal">
        <h3 className="modal-title">{title}</h3>
        {message && <p className="modal-message">{message}</p>}
        <div className="modal-actions">
          <button className="btn btn--neutral" onClick={onCancel}>{cancelLabel}</button>
          <button className={`btn ${danger ? 'btn--danger' : ''}`} onClick={onConfirm}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  )
}

function Lightbox({ src, label, onClose, onDownload }) {
  useEffect(() => {
    function onKey(e) { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <div className="lightbox" onClick={onClose}>
      <div className="lightbox-inner" onClick={e => e.stopPropagation()}>
        <div className="lightbox-bar">
          <span className="lightbox-label">{label}</span>
          <div className="lightbox-actions">
            <button className="btn btn--small" onClick={onDownload}>PPT ダウンロード</button>
            <button className="lightbox-close" onClick={onClose} title="閉じる（Esc）">✕</button>
          </div>
        </div>
        <img src={src} alt="ダッシュボード" className="lightbox-img" />
      </div>
    </div>
  )
}

export default function App() {
  const [yearMonth, setYearMonth]       = useState('')
  const [error, setError]               = useState('')
  const [phase, setPhase]               = useState('idle') // idle | checking | confirm | generating | done | error
  const [currentYM, setCurrentYM]       = useState(null)
  const [lightboxSrc, setLightboxSrc]   = useState(null)
  const [showDownload, setShowDownload] = useState(false)

  function handleChange(e) {
    setYearMonth(e.target.value)
    setError('')
    if (phase !== 'idle') setPhase('idle')
  }

  async function handleSubmit(e) {
    e.preventDefault()
    const msg = validate(yearMonth)
    if (msg) { setError(msg); return }

    const ym = parseYM(yearMonth)
    setError('')
    setPhase('checking')

    try {
      const res = await fetch(`${API}/api/check`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(ym),
      })
      if (!res.ok) throw new Error('サーバーへの接続に失敗しました')
      const info = await res.json()

      if (!info.source_exists) {
        setPhase('error')
        setError(`★営業日報${ym.year}年${ym.month}月.xlsx が data/ フォルダに見つかりません`)
        return
      }

      setCurrentYM(ym)

      const anyExists = info.daily_exists || info.report_exists ||
                        info.dashboard_exists || info.pptx_exists
      if (anyExists) {
        setPhase('confirm')
      } else {
        await runGenerate(ym)
      }
    } catch (err) {
      setPhase('error')
      setError(err.message + '。サーバーが起動しているか確認してください（uvicorn server:app --port 8000）')
    }
  }

  async function runGenerate(ym) {
    setPhase('generating')
    try {
      const res = await fetch(`${API}/api/generate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(ym),
      })
      const data = await res.json()
      if (!res.ok) throw new Error(data.detail || '生成に失敗しました')

      const { year, month } = ym
      const src = `/data/dashboard_${year}_${month}.png?t=${Date.now()}`

      setPhase('done')
      setCurrentYM(ym)
      setLightboxSrc(src)
      setShowDownload(true)
    } catch (err) {
      setPhase('error')
      setError(err.message || '生成中にエラーが発生しました')
    }
  }

  function handleConfirmOverwrite() {
    setPhase('idle')
    if (currentYM) runGenerate(currentYM)
  }

  function handleCancelOverwrite() {
    setPhase('idle')
    setCurrentYM(null)
  }

  function handleCloseLightbox() {
    setLightboxSrc(null)
    setShowDownload(false)
  }

  async function handleDownload() {
    if (!currentYM) return
    const { year, month } = currentYM
    setShowDownload(false)

    const filename = `dashboard_${year}_${month}.pptx`
    try {
      if (typeof window.showSaveFilePicker === 'function') {
        const handle = await window.showSaveFilePicker({
          suggestedName: filename,
          types: [{
            description: 'PowerPoint',
            accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] },
          }],
        })
        const response = await fetch(`/data/${filename}`)
        if (!response.ok) throw new Error('PPTXファイルが見つかりません')
        const blob = await response.blob()
        const writable = await handle.createWritable()
        await writable.write(blob)
        await writable.close()
      } else {
        const a = document.createElement('a')
        a.href = `/data/${filename}`
        a.download = filename
        a.click()
      }
    } catch (err) {
      if (err.name !== 'AbortError') {
        setError(`ダウンロードに失敗しました: ${err.message}`)
      }
    }
  }

  const isLoading    = phase === 'checking' || phase === 'generating'
  const loadingLabel = phase === 'checking' ? '確認中…' : '生成中…'

  return (
    <div className="app">
      <header className="header">
        <span className="header-icon">◈</span>
        <h1 className="header-title">営業日報 ダッシュボード生成</h1>
      </header>

      <main className="main">
        {NO_BACKEND && (
          <div className="card">
            <p className="msg msg--info">
              ダッシュボードの生成はローカル環境でのみ利用できます。<br />
              ローカルで <code>start.bat</code> を起動してください。
            </p>
          </div>
        )}
        <form className="card" onSubmit={handleSubmit} style={NO_BACKEND ? { display: 'none' } : {}}>
          <p className="card-desc">
            年月を入力すると、営業日報Excelからダッシュボード・レポート・PowerPointを生成します。
          </p>

          <label className="label" htmlFor="ym-input">対象年月</label>
          <div className="input-row">
            <input
              id="ym-input"
              className={`input ${error ? 'input--error' : ''}`}
              type="text"
              placeholder="例: 2026/5"
              value={yearMonth}
              onChange={handleChange}
              maxLength={7}
              autoComplete="off"
              disabled={isLoading}
            />
            <button className="btn" type="submit" disabled={isLoading}>
              {isLoading ? loadingLabel : '生成'}
            </button>
          </div>

          {error && <p className="msg msg--error">⚠ {error}</p>}
          {phase === 'generating' && (
            <p className="msg msg--info">⏳ データ変換・ダッシュボード生成中…</p>
          )}
          {phase === 'done' && (
            <p className="msg msg--success">✔ 生成が完了しました</p>
          )}
        </form>
      </main>

      {phase === 'confirm' && currentYM && (
        <Modal
          title="上書き確認"
          message={`${currentYM.year}/${currentYM.month} のファイルが既に存在します。上書きしますか？`}
          confirmLabel="上書きする"
          cancelLabel="キャンセル"
          danger
          onConfirm={handleConfirmOverwrite}
          onCancel={handleCancelOverwrite}
        />
      )}

      {showDownload && !lightboxSrc && (
        <Modal
          title="ダウンロード"
          message="PowerPointファイルをダウンロードしますか？"
          confirmLabel="ダウンロード"
          cancelLabel="後で"
          onConfirm={handleDownload}
          onCancel={() => setShowDownload(false)}
        />
      )}

      {lightboxSrc && (
        <Lightbox
          src={lightboxSrc}
          label={currentYM ? `ダッシュボード ${currentYM.year}/${currentYM.month}` : 'ダッシュボード'}
          onClose={handleCloseLightbox}
          onDownload={handleDownload}
        />
      )}
    </div>
  )
}
