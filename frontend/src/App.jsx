import { useState, useEffect, useRef } from 'react'
import './App.css'

const API = 'http://localhost:8000'
const STORAGE_KEY = 'dashboard_history'

function loadHistory() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]') }
  catch { return [] }
}

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

export default function App() {
  const [yearMonth, setYearMonth]     = useState('')
  const [error, setError]             = useState('')
  const [phase, setPhase]             = useState('idle') // idle | checking | confirm | generating | done | error
  const [currentYM, setCurrentYM]     = useState(null)  // {year, month}
  const [dashboardSrc, setDashboardSrc] = useState(null)
  const [showDownload, setShowDownload] = useState(false)
  const [history, setHistory]         = useState(loadHistory)
  const imgRef    = useRef(null)
  const scrollRef = useRef(false)

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(history))
  }, [history])

  // dashboardSrc が更新された後にスクロール
  useEffect(() => {
    if (scrollRef.current && imgRef.current) {
      imgRef.current.scrollIntoView({ behavior: 'smooth' })
      scrollRef.current = false
    }
  }, [dashboardSrc])

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

      setCurrentYM(ym)

      if (info.dashboard_exists || info.report_exists) {
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
      setDashboardSrc(`/data/dashboard_${year}_${month}.png?t=${Date.now()}`)
      setCurrentYM(ym)
      setPhase('done')
      setShowDownload(true)

      const key = `${year}/${month}`
      setHistory(prev => {
        const filtered = prev.filter(h => h.yearMonth !== key)
        return [{ yearMonth: key, year, month, generatedAt: new Date().toLocaleString('ja-JP') }, ...filtered]
      })
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

  function handleOpen(entry) {
    const ym = entry.year ? entry : parseYM(entry.yearMonth)
    setDashboardSrc(`/data/dashboard_${ym.year}_${ym.month}.png?t=${Date.now()}`)
    setCurrentYM(ym)
    setShowDownload(false)
    scrollRef.current = true
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
        // フォールバック: ブラウザのダウンロード
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

  function handleDelete(ym) {
    setHistory(prev => prev.filter(h => h.yearMonth !== ym))
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
        {/* 入力フォーム */}
        <form className="card" onSubmit={handleSubmit}>
          <p className="card-desc">
            年月を入力すると、対象の営業日報Excelからダッシュボード・レポート・PowerPointを生成します。
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
            <p className="msg msg--info">⏳ ダッシュボード・レポートを生成しています…</p>
          )}
          {phase === 'done' && (
            <p className="msg msg--success">✔ 生成が完了しました</p>
          )}
        </form>

        {/* ダッシュボード画像表示 */}
        {dashboardSrc && (
          <section className="card dashboard-card" ref={imgRef}>
            <div className="dashboard-header">
              <span className="label">
                {currentYM ? `ダッシュボード ${currentYM.year}/${currentYM.month}` : 'ダッシュボード'}
              </span>
              <button className="btn btn--small" onClick={() => setShowDownload(true)}>
                PPT ダウンロード
              </button>
            </div>
            <img src={dashboardSrc} alt="ダッシュボード" className="dashboard-img" />
          </section>
        )}

        {/* 生成履歴 */}
        {history.length > 0 && (
          <section className="history">
            <h2 className="history-title">生成済みダッシュボード</h2>
            <ul className="history-list">
              {history.map(h => (
                <li key={h.yearMonth} className="history-item">
                  <div className="history-info">
                    <span className="history-ym">{h.yearMonth}</span>
                    <span className="history-date">{h.generatedAt}</span>
                  </div>
                  <div className="history-actions">
                    <button className="btn btn--small" onClick={() => handleOpen(h)}>開く</button>
                    <button
                      className="btn btn--small btn--danger"
                      onClick={() => handleDelete(h.yearMonth)}
                    >
                      削除
                    </button>
                  </div>
                </li>
              ))}
            </ul>
          </section>
        )}
      </main>

      {/* 上書き確認モーダル */}
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

      {/* ダウンロード確認モーダル */}
      {showDownload && (
        <Modal
          title="ダウンロード"
          message="PowerPointファイルをダウンロードしますか？"
          confirmLabel="ダウンロード"
          cancelLabel="後で"
          onConfirm={handleDownload}
          onCancel={() => setShowDownload(false)}
        />
      )}
    </div>
  )
}
