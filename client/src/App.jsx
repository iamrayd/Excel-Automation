import { useState, useEffect, useRef } from 'react'
import UploadZone from './components/UploadZone.jsx'
import logo from './logo.png'

export default function App() {
  const downloadPanelRef = useRef(null)
  
  const [fromFile, setFromFile] = useState(null)
  const [toFile, setToFile] = useState(null)
  const [fromFileId, setFromFileId] = useState(null)
  const [toFileId, setToFileId] = useState(null)
  const [fromHeaders, setFromHeaders] = useState([])
  const [toHeaders, setToHeaders] = useState([])
  const [mappings, setMappings] = useState([])
  
  const [parsing, setParsing] = useState(false)
  const [transferring, setTransferring] = useState(false)
  const [banner, setBanner] = useState(null)
  const [downloadUrl, setDownloadUrl] = useState(null)
  const [audit, setAudit] = useState(null)

  // Automatically trigger parsing when both FROM and TO spreadsheets are set
  // Fixed race condition by using an active flag.
  useEffect(() => {
    let active = true
    if (fromFile && toFile) {
      const runParse = async () => {
        setParsing(true)
        setBanner(null)
        setDownloadUrl(null)
        setAudit(null)
        try {
          const formData = new FormData()
          formData.append('from', fromFile)
          formData.append('to', toFile)

          const res = await fetch('/api/parse', { method: 'POST', body: formData })
          let data = {}
          try {
            data = await res.json()
          } catch (e) {
            if (!active) return
            throw new Error(`Server returned status ${res.status}. Could not parse response.`)
          }

          if (!active) return
          if (!res.ok) throw new Error(data.error || 'Failed to parse sheets')

          setFromFileId(data.fromFileId)
          setToFileId(data.toFileId)
          setFromHeaders(data.fromHeaders)
          setToHeaders(data.toHeaders)
          // Assign a stable unique ID to avoid reconciliation key-index anti-patterns
          setMappings((data.mappings || []).map((m, idx) => ({
            ...m,
            id: `mapping-${idx}-${Date.now()}-${Math.random()}`
          })))
        } catch (err) {
          if (active) {
            setBanner({ type: 'error', text: err.message })
          }
        } finally {
          if (active) {
            setParsing(false)
          }
        }
      }
      runParse()
    } else {
      setFromFileId(null)
      setToFileId(null)
      setFromHeaders([])
      setToHeaders([])
      setMappings([])
      setDownloadUrl(null)
      setAudit(null)
      setBanner(null)
    }
    return () => {
      active = false
    }
  }, [fromFile, toFile])

  // Auto-scroll to download and audit reports when they load
  useEffect(() => {
    if (downloadUrl && downloadPanelRef.current) {
      downloadPanelRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' })
    }
  }, [downloadUrl])

  const resetAll = () => {
    setFromFile(null)
    setToFile(null)
  }

  const handleTransfer = async () => {
    setTransferring(true)
    setBanner(null)
    setDownloadUrl(null)
    setAudit(null)
    try {
      // Filter out empty selections and strip frontend ID property to keep request payload clean
      const activeMappings = mappings
        .filter(m => m.from && m.to)
        .map(({ from, to }) => ({ from, to }))

      if (activeMappings.length === 0) {
        throw new Error('Please configure at least one column mapping before transferring.')
      }

      const res = await fetch('/api/transfer', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fromFileId,
          toFileId,
          mappings: activeMappings
        })
      })
      let data = {}
      try {
        data = await res.json()
      } catch (e) {
        throw new Error(`Server returned status ${res.status}. Could not parse response.`)
      }

      if (!res.ok) throw new Error(data.error || 'Transformation failed')

      setBanner({ type: 'info', text: data.message })
      setDownloadUrl(data.downloadUrl)
      setAudit(data.audit)
    } catch (err) {
      setBanner({ type: 'error', text: err.message })
    } finally {
      setTransferring(false)
    }
  }

  const updateMappingRow = (id, key, val) => {
    setMappings(mappings.map(m => m.id === id ? { ...m, [key]: val } : m))
  }

  const deleteMappingRow = (id) => {
    setMappings(mappings.filter(m => m.id !== id))
  }

  const addMappingRow = () => {
    setMappings([...mappings, { id: `mapping-new-${Date.now()}-${Math.random()}`, from: '', to: '' }])
  }

  return (
    <div className="app-shell">
      <div className="main">
        <div className="topbar">
          <div className="sidebar-brand">
            <img src={logo} alt="Sheetshift Logo" className="logo-img" />
            Sheetshift
          </div>
          <button className="new-session-btn" onClick={resetAll}>
            Reset Workspace
          </button>
        </div>

        <div className="workspace">
          <p className="workspace-subtitle">
            Upload your source and target spreadsheets, customize the auto-detected column mappings, and execute the transfer.
          </p>

          {banner && <div className={`status-banner ${banner.type}`}>{banner.text}</div>}

          <div className="workspace-card">
            <div className="upload-row">
              <UploadZone label="FROM" sublabel="Source file" variant="from" file={fromFile} onFileSelected={setFromFile} />
              <div className="upload-arrow">→</div>
              <UploadZone label="TO" sublabel="Target file" variant="to" file={toFile} onFileSelected={setToFile} />
            </div>

            {parsing && (
              <div className="mapping-loading">
                <span className="loading-spinner"></span>
                <span>Parsing headers and auto-matching columns...</span>
              </div>
            )}

            {!parsing && mappings.length > 0 && (
              <div className="mapping-section">
                <h3 className="mapping-heading">Customize Column Mappings</h3>
                
                <div className="mapping-grid-header">
                  <span>Source Column (FROM)</span>
                  <span></span>
                  <span>Target Column (TO)</span>
                  <span></span>
                </div>

                <div className="mapping-grid">
                  {mappings.map((mapping) => (
                    <div className="mapping-row" key={mapping.id}>
                      <div className="mapping-col select-wrapper">
                        <select
                          value={mapping.from}
                          onChange={(e) => updateMappingRow(mapping.id, 'from', e.target.value)}
                          className="mapping-select source"
                        >
                          <option value="">(Skip column)</option>
                          {fromHeaders.map((h, i) => (
                            <option key={i} value={h}>{h}</option>
                          ))}
                        </select>
                      </div>

                      <div className="mapping-arrow-small">➔</div>

                      <div className="mapping-col select-wrapper">
                        <select
                          value={mapping.to}
                          onChange={(e) => updateMappingRow(mapping.id, 'to', e.target.value)}
                          className="mapping-select target"
                        >
                          <option value="">(Select target)</option>
                          {toHeaders.map((h, i) => (
                            <option key={i} value={h}>{h}</option>
                          ))}
                        </select>
                      </div>

                      <button
                        onClick={() => deleteMappingRow(mapping.id)}
                        className="delete-mapping-btn"
                        title="Delete mapping"
                      >
                        ×
                      </button>
                    </div>
                  ))}
                </div>

                <div className="mapping-actions">
                  <button onClick={addMappingRow} className="add-mapping-row-btn">
                    Add Column Mapping
                  </button>
                </div>
              </div>
            )}

            {mappings.length > 0 && !parsing && (
              <div className="transfer-actions">
                <button
                  className="transfer-btn"
                  disabled={transferring}
                  onClick={handleTransfer}
                >
                  {transferring ? 'Transferring Data...' : 'Transfer Data'}
                </button>
                <button className="clear-btn-text" onClick={resetAll}>
                  Clear files
                </button>
              </div>
            )}
          </div>

          {downloadUrl && (
            <div className="download-panel" ref={downloadPanelRef}>
              <div className="download-panel-header">
                <span>Transformation Complete!</span>
              </div>
              <div className="download-panel-body">
                <p>The copy-paste operation was executed successfully on the server.</p>
                <a
                  href={downloadUrl}
                  className="download-btn"
                  download
                >
                  Download Transformed Excel File
                </a>

                {audit && (
                  <div className={`audit-card ${audit.status.toLowerCase()}`}>
                    <div className="audit-card-header">
                      <span className="audit-title">AI Quality Audit</span>
                      <span className={`audit-status-badge ${audit.status.toLowerCase()}`}>
                        {audit.status}
                      </span>
                    </div>
                    <div className="audit-card-body">
                      <p className="audit-summary">{audit.summary}</p>
                      {audit.details && audit.details.length > 0 && (
                        <ul className="audit-details-list">
                          {audit.details.map((detail, idx) => (
                            <li key={idx}>{detail}</li>
                          ))}
                        </ul>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
