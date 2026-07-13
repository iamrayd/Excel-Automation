import { useRef, useState } from 'react'

export default function UploadZone({ label, sublabel, variant, file, onFileSelected }) {
  const inputRef = useRef(null)
  const [dragging, setDragging] = useState(false)

  const handleFiles = (fileList) => {
    const picked = fileList?.[0]
    if (!picked) return
    const isExcel = /\.(xlsx|xls|xlsm)$/i.test(picked.name)
    if (!isExcel) {
      alert('Please upload a .xlsx, .xls, or .xlsm file.')
      return
    }
    onFileSelected(picked)
  }

  return (
    <div className="upload-col">
      <div className="upload-col-label">
        <span>{label}</span>
        <span className="plain">{sublabel}</span>
      </div>

      <div
        className={`dropzone ${dragging ? 'dragging' : ''} ${file ? 'filled' : ''}`}
        onClick={() => inputRef.current?.click()}
        onDragOver={(e) => {
          e.preventDefault()
          setDragging(true)
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => {
          e.preventDefault()
          setDragging(false)
          handleFiles(e.dataTransfer.files)
        }}
      >
        {file ? (
          <>
            <div className="dropzone-title">{file.name}</div>
            <div className="dropzone-hint">Click to replace</div>
          </>
        ) : (
          <>
            <div className="dropzone-title">Drop your {label.toLowerCase() === 'from' ? 'source' : 'target'} Excel file</div>
            <div className="dropzone-hint">⇧ or click to browse · .xlsx, .xls, .xlsm</div>
          </>
        )}
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls,.xlsm"
          style={{ display: 'none' }}
          onChange={(e) => handleFiles(e.target.files)}
        />
      </div>
    </div>
  )
}
