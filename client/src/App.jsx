import { useState, useEffect, useRef } from 'react'
import UploadZone from './components/UploadZone.jsx'
import logo from './logo.png'

export default function App() {
  const downloadPanelRef = useRef(null)
  const logTerminalRef = useRef(null)

  // Navigation & Page State
  const [currentStep, setCurrentStep] = useState('connect') // 'connect' | 'dashboard' | 'map' | 'processing' | 'done'
  const [banner, setBanner] = useState(null)
  
  // Connection Configuration State
  const [dbConfig, setDbConfig] = useState({
    dbType: 'mssql',
    server: '',
    port: '1433',
    companyDb: '',
    username: '',
    password: ''
  })
  const [connecting, setConnecting] = useState(false)
  const [connectedDb, setConnectedDb] = useState(null)

  // Dashboard & Table selection State
  const [tables, setTables] = useState([])
  const [searchQuery, setSearchQuery] = useState('')
  const [selectedTable, setSelectedTable] = useState(null)

  // Excel template State
  const [templateFile, setTemplateFile] = useState(null)
  const [parsing, setParsing] = useState(false)
  const [fromHeaders, setFromHeaders] = useState([]) // Columns from SAP B1
  const [toHeaders, setToHeaders] = useState([])     // Columns from Excel template
  const [mappings, setMappings] = useState([])
  const [deduplicate, setDeduplicate] = useState(true)

  // Progress and Export State
  const [transferring, setTransferring] = useState(false)
  const [progressPercent, setProgressPercent] = useState(0)
  const [progressText, setProgressText] = useState('')
  const [logs, setLogs] = useState([])
  const [downloadUrl, setDownloadUrl] = useState(null)
  const [transferStats, setTransferStats] = useState(null)

  // 1. Fetch tables list when entering the dashboard
  useEffect(() => {
    if (currentStep === 'dashboard') {
      fetchTables()
    }
  }, [currentStep])

  // 2. Fetch schema when target Excel template is loaded
  useEffect(() => {
    if (templateFile && selectedTable) {
      handleParseTemplate()
    }
  }, [templateFile, selectedTable])

  // 3. Auto-scroll terminal logs during processing
  useEffect(() => {
    if (logTerminalRef.current) {
      logTerminalRef.current.scrollTop = logTerminalRef.current.scrollHeight
    }
  }, [logs])

  // 4. Auto-scroll to download stats panel on completion
  useEffect(() => {
    if (currentStep === 'done' && downloadPanelRef.current) {
      downloadPanelRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' })
    }
  }, [currentStep])

  const fetchTables = async () => {
    try {
      const url = searchQuery ? `/api/sap/tables?search=${encodeURIComponent(searchQuery)}` : '/api/sap/tables'
      const res = await fetch(url)
      const data = await res.json()
      if (res.ok) {
        setTables(data)
      }
    } catch (err) {
      console.error('Failed to load tables list:', err)
    }
  }

  // Refetch tables when user searches
  const handleSearchChange = (e) => {
    setSearchQuery(e.target.value)
  }

  useEffect(() => {
    if (currentStep === 'dashboard') {
      const delayDebounce = setTimeout(() => {
        fetchTables()
      }, 300)
      return () => clearTimeout(delayDebounce)
    }
  }, [searchQuery])

  // Connect to Database Handler
  const handleConnect = async (e) => {
    e.preventDefault()
    setConnecting(true)
    setBanner(null)
    try {
      const res = await fetch('/api/sap/connect', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(dbConfig)
      })
      let data = {}
      try {
        data = await res.json()
      } catch (err) {
        throw new Error('Server returned an unparsable response.')
      }

      if (!res.ok) throw new Error(data.error || 'Failed to authenticate connection.')

      setConnectedDb(data.db)
      setBanner({ type: 'success', text: `Success: Connection established with ${data.db}` })
      
      setTimeout(() => {
        setBanner(null)
        setCurrentStep('dashboard')
      }, 1200)

    } catch (err) {
      setBanner({ type: 'error', text: err.message })
    } finally {
      setConnecting(false)
    }
  }

  // Parse Excel template columns and query SAP schema
  const handleParseTemplate = async () => {
    setParsing(true)
    setBanner(null)
    try {
      // 1. Fetch SAP DB columns for the selected table
      const sapParams = new URLSearchParams({
        table: selectedTable.code,
        server: dbConfig.server,
        companyDb: dbConfig.companyDb
      })
      const sapRes = await fetch(`/api/sap/schema?${sapParams.toString()}`)
      const sapCols = await sapRes.json()
      if (!sapRes.ok) throw new Error(sapCols.error || 'Failed to load SAP table schema.')

      // 2. Parse Excel headers via temporary upload
      const formData = new FormData()
      formData.append('from', templateFile) // source dummy
      formData.append('to', templateFile)   // target template
      const excelRes = await fetch('/api/parse', { method: 'POST', body: formData })
      const excelData = await excelRes.json()
      if (!excelRes.ok) throw new Error(excelData.error || 'Failed to parse Excel template columns.')

      const excelCols = excelData.toHeaders || []

      setFromHeaders(sapCols)
      setToHeaders(excelCols)

      // Auto-match headers case-insensitively
      const initialMappings = []
      excelCols.forEach((excelCol) => {
        if (!excelCol) return
        const excelClean = excelCol.trim().toLowerCase()
        // Try to match SAP B1 fields
        const match = sapCols.find(sapCol => sapCol && sapCol.trim().toLowerCase() === excelClean)
        if (match) {
          initialMappings.push({
            id: `mapping-${Date.now()}-${Math.random()}`,
            from: match,
            to: excelCol
          })
        }
      })

      setMappings(initialMappings)

    } catch (err) {
      setBanner({ type: 'error', text: err.message })
    } finally {
      setParsing(false)
    }
  }

  // Run direct streaming Excel copy-paste transfer
  const handleTransfer = async () => {
    setCurrentStep('processing')
    setTransferring(true)
    setLogs([])
    setProgressPercent(10)
    setProgressText('Preparing session configuration...')

    // Log Stream simulation helper
    const logList = []
    const addLog = (text) => {
      const time = new Date().toLocaleTimeString();
      logList.push(`[${time}] ${text}`)
      setLogs([...logList])
    }

    try {
      addLog('Session handshake initialized.')
      addLog(`Selected target table: ${selectedTable.code} (${selectedTable.name})`)
      
      const activeMappings = mappings.filter(m => m.from && m.to).map(({ from, to }) => ({ from, to }))
      if (activeMappings.length === 0) {
        throw new Error('Please configure at least one column mapping.')
      }

      await new Promise(r => setTimeout(r, 600))
      setProgressPercent(25)
      setProgressText('Reading target template header styles...')
      addLog(`Uploading template file: ${templateFile.name}`)
      addLog(`Auto-Deduplication check: ${deduplicate ? 'ENABLED' : 'DISABLED'}`)

      const formData = new FormData()
      formData.append('template', templateFile)
      formData.append('table', selectedTable.code)
      formData.append('server', dbConfig.server)
      formData.append('companyDb', dbConfig.companyDb)
      formData.append('username', dbConfig.username)
      formData.append('password', dbConfig.password)
      formData.append('deduplicate', String(deduplicate))
      formData.append('mappings', JSON.stringify(activeMappings))

      await new Promise(r => setTimeout(r, 800))
      setProgressPercent(50)
      setProgressText('Streaming database records in batches of 10k...')
      addLog('Querying MS SQL Server database...')
      addLog('Copying columns: ' + activeMappings.map(m => `"${m.from}" ➔ "${m.to}"`).join(', '))

      const startTime = Date.now()
      const res = await fetch('/api/sap/transfer', {
        method: 'POST',
        body: formData
      })
      let data = {}
      try {
        data = await res.json()
      } catch (err) {
        throw new Error('Failed to retrieve streaming transfer outputs.')
      }

      if (!res.ok) throw new Error(data.error || 'Transformation execution failed.')

      setProgressPercent(80)
      setProgressText('Finalizing Excel file buffers...')
      addLog('Database streaming completed successfully.')
      addLog(`Rows imported: ${data.summary.totalImported}`)
      addLog(`Duplicates filtered out: ${data.summary.totalDuplicates}`)
      addLog('Compressing output stream...')

      await new Promise(r => setTimeout(r, 800))
      setProgressPercent(100)
      setProgressText('Transformation Complete!')

      setDownloadUrl(data.downloadUrl)
      setTransferStats(data.summary)

      setTimeout(() => {
        setCurrentStep('done')
        setTransferring(false)
      }, 600)

    } catch (err) {
      setCurrentStep('map')
      setTransferring(false)
      setBanner({ type: 'error', text: err.message })
    }
  }

  // Mapping grid functions
  const updateMappingRow = (id, key, val) => {
    setMappings(mappings.map(m => m.id === id ? { ...m, [key]: val } : m))
  }

  const deleteMappingRow = (id) => {
    setMappings(mappings.filter(m => m.id !== id))
  }

  const addMappingRow = () => {
    setMappings([...mappings, { id: `mapping-new-${Date.now()}-${Math.random()}`, from: '', to: '' }])
  }

  const handleDisconnect = () => {
    setConnectedDb(null)
    setDbConfig({ dbType: 'mssql', server: '', port: '1433', companyDb: '', username: '', password: '' })
    setTemplateFile(null)
    setSelectedTable(null)
    setCurrentStep('connect')
  }

  const handleResetWorkspace = () => {
    setTemplateFile(null)
    setSelectedTable(null)
    setDownloadUrl(null)
    setTransferStats(null)
    setBanner(null)
    if (connectedDb) {
      setCurrentStep('dashboard')
    } else {
      setCurrentStep('connect')
    }
  }

  return (
    <div className="app-shell">
      <div className="main">
        {/* TOPBAR */}
        <div className="topbar">
          <div className="sidebar-brand">
            <img src={logo} alt="Sheetshift Logo" className="logo-img" />
            Sheetshift
          </div>
          <div className="topbar-actions">
            {connectedDb && (
              <span className="status-badge connected">
                <span className="status-dot pulsing"></span>
                Connected: {connectedDb}
              </span>
            )}
            <button className="new-session-btn" onClick={handleResetWorkspace}>
              Reset Workspace
            </button>
          </div>
        </div>

        {/* WORKSPACE CONTENT CONTAINER */}
        <div className="workspace">
          
          {/* BANNER NOTIFICATION */}
          {banner && <div className={`status-banner ${banner.type}`}>{banner.text}</div>}

          {/* PAGE 1: CONNECTION LOGIN */}
          {currentStep === 'connect' && (
            <div className="login-container">
              <div className="login-card">
                <h2 className="login-title">Connect to SAP Business One</h2>
                <p className="login-subtitle">Enter your SQL Server details below to establish a database session.</p>
                
                <form onSubmit={handleConnect} className="login-form">
                  <div className="form-group">
                    <label>Database Engine Type</label>
                    <select
                      value={dbConfig.dbType}
                      onChange={(e) => setDbConfig({ ...dbConfig, dbType: e.target.value })}
                      className="form-select"
                      disabled={connecting}
                    >
                      <option value="mssql">Microsoft SQL Server (MSSQL)</option>
                      <option value="hana">SAP HANA (Service Layer)</option>
                    </select>
                  </div>

                  <div className="form-row">
                    <div className="form-group flex-2">
                      <label>Server Address</label>
                      <input
                        type="text"
                        placeholder="e.g. 192.168.1.100"
                        value={dbConfig.server}
                        onChange={(e) => setDbConfig({ ...dbConfig, server: e.target.value })}
                        className="form-input"
                        disabled={connecting}
                      />
                    </div>
                    <div className="form-group flex-1">
                      <label>Port</label>
                      <input
                        type="text"
                        placeholder="1433"
                        value={dbConfig.port}
                        onChange={(e) => setDbConfig({ ...dbConfig, port: e.target.value })}
                        className="form-input"
                        disabled={connecting}
                      />
                    </div>
                  </div>

                  <div className="form-group">
                    <label>Company Database Name</label>
                    <input
                      type="text"
                      placeholder="e.g. SBO_Live_Company"
                      value={dbConfig.companyDb}
                      onChange={(e) => setDbConfig({ ...dbConfig, companyDb: e.target.value })}
                      className="form-input"
                      disabled={connecting}
                    />
                  </div>

                  <div className="form-row">
                    <div className="form-group">
                      <label>Username</label>
                      <input
                        type="text"
                        placeholder="e.g. manager"
                        value={dbConfig.username}
                        onChange={(e) => setDbConfig({ ...dbConfig, username: e.target.value })}
                        className="form-input"
                        disabled={connecting}
                      />
                    </div>
                    <div className="form-group">
                      <label>Password</label>
                      <input
                        type="password"
                        placeholder="••••••••"
                        value={dbConfig.password}
                        onChange={(e) => setDbConfig({ ...dbConfig, password: e.target.value })}
                        className="form-input"
                        disabled={connecting}
                      />
                    </div>
                  </div>

                  <button type="submit" className="login-btn" disabled={connecting}>
                    {connecting ? 'Establishing Connection...' : 'Establish Connection'}
                  </button>

                  <div className="mock-tip">
                    <span>💡 Tip: Leave host details empty to start with offline <strong>Mock Database mode</strong>.</span>
                  </div>
                </form>
              </div>
            </div>
          )}

          {/* PAGE 2: TABLE EXPLORER DASHBOARD */}
          {currentStep === 'dashboard' && (
            <div className="dashboard-container">
              <h2 className="dashboard-title">Select Database Table</h2>
              <p className="dashboard-subtitle">Choose a SAP B1 table or entity to start mapping your export template.</p>

              {/* SEARCH BAR */}
              <div className="search-bar-wrapper">
                <input
                  type="text"
                  placeholder="Search table by code or description (e.g. OCRD, Items, Invoices...)"
                  value={searchQuery}
                  onChange={handleSearchChange}
                  className="search-input"
                />
                <span className="search-hint">⌘K</span>
              </div>

              {/* TABLES LIST GRID */}
              <div className="tables-grid">
                {tables.map((table) => (
                  <div className="table-card" key={table.code}>
                    <div className="table-card-header">
                      <span className="table-icon">📄</span>
                      <span className="table-code-badge">{table.code}</span>
                    </div>
                    <h3 className="table-card-title">{table.name}</h3>
                    <p className="table-card-desc">{table.desc}</p>
                    <div className="table-card-meta">
                      <span>{table.fields} Fields</span>
                      <span>·</span>
                      <span>{table.count.toLocaleString()} Records</span>
                    </div>
                    <button
                      onClick={() => {
                        setSelectedTable(table)
                        setCurrentStep('map')
                      }}
                      className="map-data-btn"
                    >
                      Map Data ➔
                    </button>
                  </div>
                ))}
                {tables.length === 0 && (
                  <div className="empty-search">
                    <p>No SAP B1 tables matched your query.</p>
                  </div>
                )}
              </div>

              <div className="disconnect-row">
                <button className="disconnect-btn" onClick={handleDisconnect}>
                  Disconnect Database
                </button>
              </div>
            </div>
          )}

          {/* PAGE 3: TEMPLATE UPLOAD & COLUMN MAPPING */}
          {currentStep === 'map' && selectedTable && (
            <div className="map-view-container">
              <div className="back-navigation">
                <button className="back-link" onClick={() => setCurrentStep('dashboard')}>
                  ← Back to Tables
                </button>
              </div>

              <h2 className="dashboard-title">Map SAP B1 Table to Excel Template</h2>
              <p className="dashboard-subtitle">
                Selected Entity: <strong>{selectedTable.name} ({selectedTable.code})</strong>. Upload your template file to begin.
              </p>

              {/* UPLOAD ZONE */}
              <div className="centered-upload">
                <UploadZone
                  label="Target Excel Template"
                  sublabel="The spreadsheet where records will be injected"
                  variant="to"
                  file={templateFile}
                  onFileSelected={setTemplateFile}
                />
              </div>

              {/* PARSING INDICATOR */}
              {parsing && (
                <div className="mapping-loading">
                  <span className="loading-spinner"></span>
                  <span>Fetching fields and auto-detecting column mappings...</span>
                </div>
              )}

              {/* INTERACTIVE MAPPING GRID */}
              {!parsing && templateFile && mappings.length > 0 && (
                <div className="workspace-card mapping-card-grow">
                  <div className="mapping-section">
                    <h3 className="mapping-heading">Customize Column Mappings</h3>
                    
                    <div className="mapping-grid-header">
                      <span>Source Field (SAP B1 {selectedTable.code})</span>
                      <span></span>
                      <span>Target Column (Excel Template)</span>
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
                              <option value="">(Skip field)</option>
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

                    {/* DEDUPLICATION TOGGLE */}
                    <div className="dedup-option-row">
                      <label className="switch-wrapper">
                        <input
                          type="checkbox"
                          checked={deduplicate}
                          onChange={(e) => setDeduplicate(e.target.checked)}
                          className="switch-checkbox"
                        />
                        <span className="switch-slider"></span>
                      </label>
                      <div className="switch-details">
                        <span className="switch-label">Enable ID Deduplication</span>
                        <span className="switch-sublabel">Automatically skip rows with duplicate Customer/Product IDs to prevent redundant entries.</span>
                      </div>
                    </div>
                  </div>

                  <div className="transfer-actions">
                    <button
                      className="transfer-btn"
                      disabled={transferring}
                      onClick={handleTransfer}
                    >
                      {transferring ? 'Transferring Data...' : 'Transfer Data'}
                    </button>
                    <button className="clear-btn-text" onClick={handleResetWorkspace}>
                      Clear files
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* PAGE 4: PROCESSING DATA TRANSFER LOGS */}
          {currentStep === 'processing' && selectedTable && (
            <div className="processing-container">
              <h2 className="dashboard-title">Processing Data Transfer</h2>
              <p className="dashboard-subtitle">
                Streaming records from table <strong>{selectedTable.code}</strong> into your target Excel template...
              </p>

              <div className="workspace-card progress-card">
                <div className="progress-details-row">
                  <span className="progress-label-text">{progressText}</span>
                  <span className="progress-value-text">{progressPercent}%</span>
                </div>
                
                <div className="progress-bar-track">
                  <div className="progress-bar-fill" style={{ width: `${progressPercent}%` }}></div>
                </div>

                {/* LOGS TERMINAL */}
                <div className="log-terminal-header">
                  <span>Import Activity Logs</span>
                </div>
                <div className="log-terminal" ref={logTerminalRef}>
                  {logs.map((log, i) => (
                    <div className="log-line" key={i}>{log}</div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* PAGE 5: TRANSFORMATION DONE & DOWNLOAD */}
          {currentStep === 'done' && transferStats && (
            <div className="done-container" ref={downloadPanelRef}>
              <div className="download-panel border-success-green">
                <div className="download-panel-header bg-success-green">
                  <span>Transformation Complete!</span>
                </div>
                <div className="download-panel-body">
                  <p className="success-paragraph">
                    The copy-paste operation was executed successfully on the server.
                  </p>
                  
                  <a
                    href={downloadUrl}
                    className="download-btn btn-success-green"
                    download
                  >
                    Download Transformed Excel File
                  </a>

                  {/* EXECUTION SUMMARY GRID */}
                  <div className="summary-stats-header">Execution Summary Statistics</div>
                  <div className="stats-grid">
                    <div className="stat-card">
                      <div className="stat-label">Total Rows Imported</div>
                      <div className="stat-value">{transferStats.totalImported.toLocaleString()}</div>
                    </div>
                    <div className="stat-card">
                      <div className="stat-label">Duplicates Filtered</div>
                      <div className="stat-value">{transferStats.totalDuplicates.toLocaleString()}</div>
                    </div>
                    <div className="stat-card">
                      <div className="stat-label">Execution Time</div>
                      <div className="stat-value">{transferStats.timeElapsed}s</div>
                    </div>
                    <div className="stat-card">
                      <div className="stat-label">Output File Size</div>
                      <div className="stat-value">{transferStats.fileSize} MB</div>
                    </div>
                  </div>

                  <div className="new-export-actions">
                    <button className="new-export-btn" onClick={handleResetWorkspace}>
                      Start New Export
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
