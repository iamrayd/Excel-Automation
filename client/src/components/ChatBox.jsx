import { useEffect, useRef } from 'react'

const SUGGESTIONS = [
  'Match products using SKU',
  'Combine duplicate rows',
  'Convert dates to YYYY-MM-DD',
  'Generate VBA that transforms source into target',
  'Explain the differences between both files',
]

const MAX_TEXTAREA_HEIGHT = 160

export default function ChatBox({ prompt, onPromptChange, onGenerate, onClear, onUploadNew, canGenerate, generating }) {
  const textareaRef = useRef(null)

  // Auto-grow the textarea as the user types, capped at MAX_TEXTAREA_HEIGHT.
  useEffect(() => {
    const el = textareaRef.current
    if (!el) return
    el.style.height = 'auto'
    el.style.height = `${Math.min(el.scrollHeight, MAX_TEXTAREA_HEIGHT)}px`
  }, [prompt])

  const handleKeyDown = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      if (canGenerate && !generating) onGenerate()
    }
  }

  return (
    <>
      <div className="card-footer">
        <button className="footer-link" onClick={onClear}>
          🗑 Clear
        </button>
      </div>

      <div className="chat-input-bar">
        <button className="chat-icon-btn" title="Upload new files" onClick={onUploadNew}>
          +
        </button>

        <textarea
          ref={textareaRef}
          className="chat-textarea"
          rows={1}
          placeholder='Describe how you want to transform the "From" Excel file into the "To" Excel file...'
          value={prompt}
          onChange={(e) => onPromptChange(e.target.value)}
          onKeyDown={handleKeyDown}
        />

        <button className="send-btn" disabled={!canGenerate || generating} onClick={onGenerate} title="Generate Transformation">
          {generating ? '…' : '↑'}
        </button>
      </div>
      <div className="chat-input-hint">Enter to send · Shift + Enter for a new line</div>
    </>
  )
}
