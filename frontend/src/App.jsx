import { useState, useEffect } from 'react'
import './App.css'

const API = import.meta.env.VITE_API_URL || 'http://localhost:8000'
const NO_BACKEND = import.meta.env.PROD && !import.meta.env.VITE_API_URL
const IS_CLOUD = !!import.meta.env.VITE_API_URL

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

function b64toBlob(b64, mime) {
  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0))
  return new Blob([bytes], { type: mime })
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

function downloadPath(path, filename) {
  const a = document.createElement('a')
  a.href = path
  a.download = filename
  a.click()
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

function Lightbox({ src, label, onClose, onDownloadPptx, onDownloadDaily, showExcel = true }) {
  useEffect(() => {
    function onKey(e) { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <div className="lightbox">
      <div className="lightbox-inner">
        <div className="lightbox-bar">
          <span className="lightbox-label">{label}</span>
          <div className="lightbox-actions">
            {showExcel && <button className="btn btn--small" onClick={onDownloadDaily}>日次 Excel</button>}
            <button className="btn btn--small" onClick={onDownloadPptx}>PowerPoint</button>
            <button className="lightbox-close" onClick={onClose} title="閉じる（Esc）">✕</button>
          </div>
        </div>
        <img src={src} alt="ダッシュボード" className="lightbox-img" />
      </div>
    </div>
  )
}

export default function App() {
  const [yearMonth, setYearMonth]     = useState('')
  const [error, setError]             = useState('')
  const [phase, setPhase]             = useState('idle')
  const [currentYM, setCurrentYM]     = useState(null)
  const [lightboxSrc, setLightboxSrc] = useState(null)

  // ダウンロード用 blob（クラウドモード・生成直後）
  const [pptxBlob, setPptxBlob]         = useState(null)
  const [pptxFilename, setPptxFilename] = useState(null)
  const [dailyBlob, setDailyBlob]       = useState(null)
  const [dailyFilename, setDailyFilename] = useState(null)
  // 生成済みレポート一覧（Drive）
  const [reports, setReports]           = useState([])
  const [reportsLoading, setReportsLoading] = useState(false)
  const [loadingReportKey, setLoadingReportKey] = useState(null)  // 表示中の年月キー

  // 既存レポートを表示中（Drive ファイル ID を保持）
  const [viewReport, setViewReport]     = useState(null)

  // アプリ起動時にレポート一覧を取得
  useEffect(() => {
    if (!IS_CLOUD) return
    loadReports()
  }, [])

  async function loadReports() {
    setReportsLoading(true)
    try {
      const res = await fetch(`${API}/api/list-reports`)
      const data = await res.json()
      const sorted = (data.reports || []).sort((a, b) =>
        b.year !== a.year ? b.year - a.year : b.month - a.month
      )
      setReports(sorted)
    } catch {
      // 取得失敗は無視
    } finally {
      setReportsLoading(false)
    }
  }

  function handleChange(e) {
    setYearMonth(e.target.value)
    setError('')
    if (phase !== 'idle') setPhase('idle')
  }

  async function doCloudGenerate(ym) {
    const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    const PPTX = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

    const tryGenerate = async () => {
      const res = await fetch(`${API}/api/drive-generate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(ym),
      })
      const data = await res.json()
      if (!res.ok) throw new Error(data.detail || '生成に失敗しました')
      return data
    }

    setViewReport(null)
    setPhase('generating')
    try {
      let data
      try {
        data = await tryGenerate()
      } catch (err) {
        if (err.message === 'Failed to fetch') {
          setPhase('waking')
          await new Promise(r => setTimeout(r, 12000))
          setPhase('generating')
          data = await tryGenerate()
        } else {
          throw err
        }
      }

      setPptxBlob(b64toBlob(data.pptx_base64, PPTX))
      setPptxFilename(data.pptx_filename)
      setDailyBlob(b64toBlob(data.daily_base64, XLSX))
      setDailyFilename(data.daily_filename)
      const pngUrl = URL.createObjectURL(b64toBlob(data.png_base64, 'image/png'))
      setCurrentYM(ym)
      setPhase('done')
      setLightboxSrc(pngUrl)

      // レポート一覧を更新
      loadReports()
    } catch (err) {
      setPhase('error')
      setError(
        err.message === 'Failed to fetch'
          ? 'サーバーに接続できませんでした。しばらくしてから再試行してください。'
          : err.message || '生成中にエラーが発生しました'
      )
    }
  }

  async function handleSubmit(e) {
    e.preventDefault()
    const msg = validate(yearMonth)
    if (msg) { setError(msg); return }

    const ym = parseYM(yearMonth)
    setError('')

    if (IS_CLOUD) {
      // 同月ファイルが Drive に既存なら上書き確認
      const exists = reports.some(r => r.year === ym.year && r.month === ym.month)
      if (exists) {
        setCurrentYM(ym)
        setPhase('confirm')
        return
      }
      await doCloudGenerate(ym)
      return
    }

    // ローカルモード
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
      setPhase('done')
      setCurrentYM(ym)
      setLightboxSrc(`/data/dashboard_${year}_${month}.png?t=${Date.now()}`)
    } catch (err) {
      setPhase('error')
      setError(err.message || '生成中にエラーが発生しました')
    }
  }

  function handleConfirmOverwrite() {
    setPhase('idle')
    if (IS_CLOUD && currentYM) {
      doCloudGenerate(currentYM)
    } else if (currentYM) {
      runGenerate(currentYM)
    }
  }

  function handleCancelOverwrite() {
    setPhase('idle')
    setCurrentYM(null)
  }

  // Drive の PPTX から画像を抽出して表示
  async function handleViewReport(report) {
    if (!report.pptx_id) return
    const key = `${report.year}-${report.month}`
    setLoadingReportKey(key)
    try {
      const res = await fetch(`${API}/api/get-pptx-image/${report.pptx_id}`)
      const data = await res.json()
      if (!res.ok) throw new Error(data.detail || '取得に失敗しました')
      const pngUrl = URL.createObjectURL(b64toBlob(data.png_base64, 'image/png'))
      setViewReport(report)
      setCurrentYM({ year: report.year, month: report.month })
      setLightboxSrc(pngUrl)
    } catch (err) {
      setError(err.message || 'レポートの取得に失敗しました')
    } finally {
      setLoadingReportKey(null)
    }
  }

  // Drive ファイルをダウンロード
  async function downloadDriveFile(fileId, filename) {
    try {
      const res = await fetch(`${API}/api/get-file/${fileId}`)
      const data = await res.json()
      downloadBlob(b64toBlob(data.base64, data.mime), filename)
    } catch {
      alert('ダウンロードに失敗しました')
    }
  }

  function handleCloseLightbox() {
    if (lightboxSrc && lightboxSrc.startsWith('blob:')) URL.revokeObjectURL(lightboxSrc)
    setLightboxSrc(null)
    setViewReport(null)
  }

  function handleDownloadPptx() {
    if (viewReport?.pptx_id) {
      downloadDriveFile(viewReport.pptx_id,
        `dashboard_${viewReport.year}_${viewReport.month}.pptx`)
      return
    }
    if (IS_CLOUD && pptxBlob) { downloadBlob(pptxBlob, pptxFilename); return }
    if (currentYM) downloadPath(`/data/dashboard_${currentYM.year}_${currentYM.month}.pptx`,
                                `dashboard_${currentYM.year}_${currentYM.month}.pptx`)
  }

  function handleDownloadDaily() {
    if (viewReport?.daily_id) {
      downloadDriveFile(viewReport.daily_id, `daily_${viewReport.year}_${viewReport.month}.xlsx`)
      return
    }
    if (IS_CLOUD && dailyBlob) { downloadBlob(dailyBlob, dailyFilename); return }
    if (currentYM) downloadPath(`/data/daily_${currentYM.year}_${currentYM.month}.xlsx`,
                                `daily_${currentYM.year}_${currentYM.month}.xlsx`)
  }

  const isLoading    = phase === 'checking' || phase === 'generating' || phase === 'waking'
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
            {IS_CLOUD
              ? 'Google Drive の共有フォルダに保存した★営業日報Excelから、ダッシュボード・レポート・PowerPointを生成します。'
              : '年月を入力すると、営業日報Excelからダッシュボード・レポート・PowerPointを生成します。'}
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
          {phase === 'waking' && (
            <p className="msg msg--info">⏳ サーバーを起動しています（初回は約50秒かかります）…</p>
          )}
          {phase === 'done' && (
            <p className="msg msg--success">✔ 生成が完了しました</p>
          )}
        </form>

        {/* 生成済みレポート一覧 */}
        {IS_CLOUD && (
          <div className="card reports-card">
            <h2 className="reports-title">生成済みレポート</h2>
            {reportsLoading && <p className="msg msg--info">読み込み中…</p>}
            {!reportsLoading && reports.length === 0 && (
              <p className="msg">まだ生成済みのレポートはありません</p>
            )}
            {!reportsLoading && reports.length > 0 && (() => {
              const COLS = 3
              const sorted = [...reports].sort((a, b) =>
                b.year !== a.year ? b.year - a.year : b.month - a.month)
              const numRows = Math.ceil(sorted.length / COLS)
              const cells = Array.from({ length: numRows * COLS }, (_, i) =>
                sorted[(i % COLS) * numRows + Math.floor(i / COLS)] ?? null
              )
              return (
                <ul className="reports-list">
                  {cells.map((r, i) => {
                    if (!r) return <li key={`pad-${i}`} className="report-item-pad" />
                    const key = `${r.year}-${r.month}`
                    const isViewing = loadingReportKey === key
                    return (
                      <li key={key} className="report-item">
                        <span className="report-label">{r.year}年{r.month}月</span>
                        <button
                          className="btn btn--small"
                          onClick={() => handleViewReport(r)}
                          disabled={isViewing || !!loadingReportKey}
                        >
                          {isViewing ? '読込中…' : '表示'}
                        </button>
                      </li>
                    )
                  })}
                </ul>
              )
            })()}
          </div>
        )}
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

      {lightboxSrc && (
        <Lightbox
          src={lightboxSrc}
          label={currentYM ? `ダッシュボード ${currentYM.year}/${currentYM.month}` : 'ダッシュボード'}
          onClose={handleCloseLightbox}
          onDownloadPptx={handleDownloadPptx}
          onDownloadDaily={handleDownloadDaily}
          showExcel={!viewReport}
        />
      )}
    </div>
  )
}
