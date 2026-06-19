import { useState, useEffect } from 'react'
import './App.css'

const STORAGE_KEY = 'dashboard_history'

function loadHistory() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]')
  } catch {
    return []
  }
}

function saveHistory(history) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(history))
}

function validate(value) {
  if (!/^\d{4}\/\d{1,2}$/.test(value)) return '形式は yyyy/m で入力してください（例: 2026/5）'
  const [, m] = value.split('/').map(Number)
  if (m < 1 || m > 12) return '月は 1〜12 で入力してください'
  return ''
}

function toFilename(yearMonth) {
  const [y, m] = yearMonth.split('/')
  return `★営業日報${y}年${Number(m)}月.xlsx`
}

export default function App() {
  const [yearMonth, setYearMonth] = useState('')
  const [error, setError]         = useState('')
  const [status, setStatus]       = useState('idle') // idle | loading | done | error
  const [history, setHistory]     = useState(loadHistory)

  // history が変わるたびに localStorage へ書き込む
  useEffect(() => {
    saveHistory(history)
  }, [history])

  function handleChange(e) {
    setYearMonth(e.target.value)
    setError('')
    setStatus('idle')
  }

  async function handleSubmit(e) {
    e.preventDefault()
    const msg = validate(yearMonth)
    if (msg) { setError(msg); return }

    setStatus('loading')
    setError('')

    try {
      // TODO: ダッシュボード生成APIを呼び出す（後で実装）
      await new Promise(r => setTimeout(r, 800))

      const entry = {
        yearMonth,
        filename: toFilename(yearMonth),
        generatedAt: new Date().toLocaleString('ja-JP'),
      }
      setHistory(prev => {
        // 同じ年月が既にあれば上書き、なければ先頭に追加
        const filtered = prev.filter(h => h.yearMonth !== yearMonth)
        return [entry, ...filtered]
      })
      setStatus('done')
    } catch {
      setStatus('error')
      setError('ダッシュボードの生成に失敗しました')
    }
  }

  function handleDelete(yearMonth) {
    setHistory(prev => prev.filter(h => h.yearMonth !== yearMonth))
  }

  const filename = yearMonth && !validate(yearMonth) ? toFilename(yearMonth) : ''

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
            年月を入力すると、対象の営業日報Excelからダッシュボードを生成します。
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
            />
            <button className="btn" type="submit" disabled={status === 'loading'}>
              {status === 'loading' ? '生成中…' : '生成'}
            </button>
          </div>

          {error   && <p className="msg msg--error">⚠ {error}</p>}
          {filename && <p className="msg msg--info">対象ファイル: <code>{filename}</code></p>}
          {status === 'done' && <p className="msg msg--success">✔ ダッシュボードを生成しました</p>}
        </form>

        {/* 生成履歴 */}
        {history.length > 0 && (
          <section className="history">
            <h2 className="history-title">生成済みダッシュボード</h2>
            <ul className="history-list">
              {history.map(h => (
                <li key={h.yearMonth} className="history-item">
                  <div className="history-info">
                    <span className="history-ym">{h.yearMonth}</span>
                    <span className="history-file">{h.filename}</span>
                    <span className="history-date">{h.generatedAt}</span>
                  </div>
                  <div className="history-actions">
                    <button className="btn btn--small">開く</button>
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
    </div>
  )
}
