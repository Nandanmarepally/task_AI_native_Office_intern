import { useState, useRef, useCallback, useMemo, useEffect } from 'react'
import { ArrowUp, ArrowDown, ArrowUpDown, Filter } from 'lucide-react'
import './App.css'
import { createEngine } from './engine/core.js'
import { saveState, loadState, deserializeEngineData } from './storage.js'

const TOTAL_ROWS = 50
const TOTAL_COLS = 50

export default function App() {
  // Engine instance is created once and reused across renders
  // Note: The engine maintains its own internal state, so React state is only used for UI updates
  const [engine] = useState(() => createEngine(TOTAL_ROWS, TOTAL_COLS))
  const [version, setVersion] = useState(0)
  const [selectedCell, setSelectedCell] = useState(null)
  const [editingCell, setEditingCell] = useState(null)
  const [editValue, setEditValue] = useState('')
  // Cell styles are stored separately from engine data
  // Format: { "row,col": { bold: bool, italic: bool, ... } }
  const [cellStyles, setCellStyles] = useState({})
  const cellInputRef = useRef(null)

  // ────── Sorting & Filtering State ──────
  // columnSortState: { colIndex: 'asc' | 'desc' | null }
  const [columnSortState, setColumnSortState] = useState({})
  // columnFilters: { colIndex: Set of selected values (when empty, all values shown) }
  const [columnFilters, setColumnFilters] = useState({})
  // filterMenuOpen: tracks which column (if any) has its filter menu open
  const [filterMenuOpen, setFilterMenuOpen] = useState(null)
  // filterMenuPosition: stores calculated position for filter menu
  const [filterMenuPosition, setFilterMenuPosition] = useState({ top: 0, left: 0 })
  const filterButtonRefs = useRef({})

  // ────── Local Storage Persistence ──────
  const saveTimeoutRef = useRef(null)
  
  // Load persisted state on mount (only runs once)
  useEffect(() => {
    const persistedState = loadState()
    if (persistedState) {
      console.log('Restored spreadsheet from local storage')
      try {
        // Restore engine data (cells)
        deserializeEngineData(engine, persistedState.engineData)
        
        // Restore cell styles and UI state
        setCellStyles(persistedState.cellStyles)
        
        // Convert columnSortState keys from strings to numbers (JSON serialization converts keys to strings)
        const sortStateWithNumKeys = Object.fromEntries(
          Object.entries(persistedState.columnSortState).map(([key, value]) => [parseInt(key), value])
        )
        setColumnSortState(sortStateWithNumKeys)
        
        // Convert columnFilters keys from strings to numbers (same reason)
        const filtersWithNumKeys = Object.fromEntries(
          Object.entries(persistedState.columnFilters).map(([key, value]) => [parseInt(key), value])
        )
        setColumnFilters(filtersWithNumKeys)
      } catch (err) {
        console.error('Error restoring persisted state:', err)
      }
    }
  }, [engine])

  // Debounced autosave - saves whenever engine data, cellStyles, columnSortState, or columnFilters change
  // The version trigger tracks engine changes (incremented whenever forceRerender is called)
  // Does NOT save undo/redo history (as per requirements)
  useEffect(() => {
    // Clear existing timeout
    if (saveTimeoutRef.current) {
      clearTimeout(saveTimeoutRef.current)
    }

    // Set a new timeout for debounced save (500ms)
    saveTimeoutRef.current = setTimeout(() => {
      const success = saveState(engine, cellStyles, columnSortState, columnFilters)
      if (!success) {
        console.warn('Failed to save state to local storage')
      }
    }, 500)

    // Cleanup: cancel pending save if dependencies change again before timeout fires
    return () => {
      if (saveTimeoutRef.current) {
        clearTimeout(saveTimeoutRef.current)
      }
    }
  }, [version, engine, cellStyles, columnSortState, columnFilters])

  const forceRerender = useCallback(() => setVersion(v => v + 1), [])

  // ────── Cell style helpers ──────

  const getCellStyle = useCallback((row, col) => {
    const key = `${row},${col}`
    return cellStyles[key] || {
      bold: false, italic: false, underline: false,
      bg: 'white', color: '#202124', align: 'left', fontSize: 13
    }
  }, [cellStyles])

  const updateCellStyle = useCallback((row, col, updates) => {
    const key = `${row},${col}`
    setCellStyles(prev => ({
      ...prev,
      [key]: { ...getCellStyle(row, col), ...updates }
    }))
  }, [getCellStyle])

  // ────── Sorting & Filtering Logic ──────

  // Compute visibleRows: filtered + sorted array of row indices
  const visibleRows = useMemo(() => {
    // Step 1: Apply filters
    let filtered = Array.from({ length: TOTAL_ROWS }, (_, i) => i)

    for (const [colStr, filterSet] of Object.entries(columnFilters)) {
      if (filterSet.size === 0) continue
      const col = parseInt(colStr)

      filtered = filtered.filter(rowIndex => {
        const cellData = engine.getCell(rowIndex, col)
        // Display value: computed (formula result) or raw value
        const displayValue = cellData.computed !== null && cellData.computed !== ''
          ? String(cellData.computed)
          : cellData.raw
        return filterSet.has(displayValue)
      })
    }

    // Step 2: Apply sorting
    const sortCol = Object.keys(columnSortState)[0]
    if (sortCol !== undefined) {
      const col = parseInt(sortCol)
      const sortDir = columnSortState[col]

      filtered.sort((aIdx, bIdx) => {
        const aCell = engine.getCell(aIdx, col)
        const bCell = engine.getCell(bIdx, col)

        const aVal = aCell.computed !== null && aCell.computed !== ''
          ? aCell.computed
          : aCell.raw
        const bVal = bCell.computed !== null && bCell.computed !== ''
          ? bCell.computed
          : bCell.raw

        // Numeric comparison first
        const aNum = parseFloat(aVal)
        const bNum = parseFloat(bVal)

        let comparison
        if (!isNaN(aNum) && !isNaN(bNum)) {
          comparison = aNum - bNum
        } else {
          comparison = String(aVal).localeCompare(String(bVal))
        }

        return sortDir === 'asc' ? comparison : -comparison
      })
    }

    return filtered
  }, [engine, columnFilters, columnSortState])

  const toggleColumnSort = useCallback((col) => {
    setColumnSortState(prev => {
      const current = prev[col] || null
      let next
      if (current === null) next = 'asc'
      else if (current === 'asc') next = 'desc'
      else next = null

      const updated = { ...prev }
      if (next === null) {
        delete updated[col]
      } else {
        updated[col] = next
      }
      return updated
    })
  }, [])

  const calculateFilterMenuPosition = useCallback((colIndex) => {
    const button = filterButtonRefs.current[colIndex]
    if (!button) return
    
    const rect = button.getBoundingClientRect()
    const scrollContainer = document.querySelector('.grid-scroll')
    const scrollLeft = scrollContainer ? scrollContainer.scrollLeft : 0
    
    // Position below the button, adjust horizontally to prevent cutoff
    const top = rect.bottom + 4
    let left = rect.left
    
    // Prevent menu from going off-screen to the right
    const menuWidth = 280 // max-width from CSS
    if (left + menuWidth > window.innerWidth) {
      left = window.innerWidth - menuWidth - 8
    }
    
    // Prevent menu from going off-screen to the left
    if (left < 4) {
      left = 4
    }
    
    setFilterMenuPosition({ top, left })
  }, [])

  const toggleFilterValue = useCallback((col, value) => {
    setColumnFilters(prev => {
      const filterSet = prev[col] ? new Set(prev[col]) : new Set()
      if (filterSet.has(value)) {
        filterSet.delete(value)
      } else {
        filterSet.add(value)
      }

      const updated = { ...prev }
      if (filterSet.size === 0) {
        delete updated[col]
      } else {
        updated[col] = filterSet
      }
      return updated
    })
  }, [])

  const clearColumnFilter = useCallback((col) => {
    setColumnFilters(prev => {
      const updated = { ...prev }
      delete updated[col]
      return updated
    })
  }, [])

  const clearAllFilters = useCallback(() => {
    setColumnFilters({})
    setColumnSortState({})
  }, [])

  // Get unique values for a column (for filter dropdown)
  const getColumnUniqueValues = useCallback((col) => {
    const values = new Set()
    for (let row = 0; row < TOTAL_ROWS; row++) {
      const cellData = engine.getCell(row, col)
      const displayValue = cellData.computed !== null && cellData.computed !== ''
        ? String(cellData.computed)
        : cellData.raw
      if (displayValue) {
        values.add(displayValue)
      }
    }
    return Array.from(values).sort((a, b) => {
      const numA = parseFloat(a)
      const numB = parseFloat(b)
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB
      return String(a).localeCompare(String(b))
    })
  }, [engine])

  // Close filter menu when clicking outside or on cells
  useEffect(() => {
    const handleClick = (e) => {
      // Close filter menu when clicking outside filter buttons/menus
      if (!e.target.closest('.filter-dropdown-wrapper') && !e.target.closest('.filter-btn')) {
        setFilterMenuOpen(null)
      }
    }

    if (filterMenuOpen !== null) {
      document.addEventListener('click', handleClick)
      return () => document.removeEventListener('click', handleClick)
    }
  }, [filterMenuOpen])

  // ────── Cell editing ──────

  const startEditing = useCallback((row, col) => {
    setSelectedCell({ r: row, c: col })
    setEditingCell({ r: row, c: col })
    const cellData = engine.getCell(row, col)
    setEditValue(cellData.raw)
    setTimeout(() => cellInputRef.current?.focus(), 0)
  }, [engine])

  const commitEdit = useCallback((row, col) => {
    // Only commit if the value actually changed to avoid unnecessary recalculations
    const currentCell = engine.getCell(row, col)
    if (currentCell.raw !== editValue) {
      engine.setCell(row, col, editValue)
      forceRerender()
    }
    setEditingCell(null)
  }, [engine, editValue, forceRerender])

  const handleCellClick = useCallback((row, col) => {
    if (editingCell && (editingCell.r !== row || editingCell.c !== col)) {
      commitEdit(editingCell.r, editingCell.c)
    }
    if (!editingCell || editingCell.r !== row || editingCell.c !== col) {
      startEditing(row, col)
    }
  }, [editingCell, commitEdit, startEditing])

  // ────── Keyboard navigation ──────

  const handleKeyDown = useCallback((event, row, col) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col)
    } else if (event.key === 'Tab') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    } else if (event.key === 'Escape') {
      setEditValue(engine.getCell(row, col).raw)
      setEditingCell(null)
    } else if (event.key === 'ArrowDown') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col)
    } else if (event.key === 'ArrowUp') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.max(row - 1, 0), col)
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault()
      commitEdit(row, col)
      if (col > 0) {
        startEditing(row, col - 1)
      } else if (row > 0) {
        startEditing(row - 1, engine.cols - 1)
      }
    } else if (event.key === 'ArrowRight') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    }
  }, [engine, commitEdit, startEditing])

  // ────── Formula bar handlers ──────

  const handleFormulaBarKeyDown = useCallback((event) => {
    if (!editingCell) return
    handleKeyDown(event, editingCell.r, editingCell.c)
  }, [editingCell, handleKeyDown])

  const handleFormulaBarFocus = useCallback(() => {
    if (selectedCell && !editingCell) {
      setEditingCell(selectedCell)
      setEditValue(engine.getCell(selectedCell.r, selectedCell.c).raw)
    }
  }, [selectedCell, editingCell, engine])

  const handleFormulaBarChange = useCallback((value) => {
    if (!editingCell && selectedCell) setEditingCell(selectedCell)
    setEditValue(value)
  }, [editingCell, selectedCell])

  // ────── Undo / Redo ──────

  const handleUndo = useCallback(() => { if (engine.undo()) forceRerender() }, [engine, forceRerender])
  const handleRedo = useCallback(() => { if (engine.redo()) forceRerender() }, [engine, forceRerender])

  // ────── Clipboard (Copy/Paste) ──────

  // Copy selected cell (Ctrl+C) - copies computed value
  const handleCopy = useCallback((e) => {
    if (!selectedCell || editingCell) return
    e.preventDefault()
    
    const cellData = engine.getCell(selectedCell.r, selectedCell.c)
    // Copy the computed value (formula result), not the raw formula
    const value = cellData.computed !== null && cellData.computed !== '' 
      ? String(cellData.computed) 
      : cellData.raw
    
    navigator.clipboard.writeText(value).catch(() => {
      // Fallback if clipboard API fails
      console.warn('Copy to clipboard failed')
    })
  }, [selectedCell, editingCell, engine])

  // Paste from clipboard (Ctrl+V) - handles tab-separated multi-row/multi-col data
  const handlePaste = useCallback(async (e) => {
    if (!selectedCell || editingCell) return
    e.preventDefault()
    
    try {
      const text = await navigator.clipboard.readText()
      
      // Split by newlines first (rows), then by tabs (columns)
      const rows = text.split('\n').filter(line => line.trim())
      const maxRows = rows.length
      const maxCols = Math.max(...rows.map(row => row.split('\t').length))
      
      // Check bounds
      const endRow = selectedCell.r + maxRows - 1
      const endCol = selectedCell.c + maxCols - 1
      
      if (endRow >= engine.rows || endCol >= engine.cols) {
        console.warn(`Paste area exceeds grid bounds (${endRow + 1} rows, ${endCol + 1} cols)`)
        // Still paste what fits
      }
      
      // Paste data into grid starting from selectedCell
      // Each paste action is tracked as a single undo entry via pushToUndoStack
      for (let i = 0; i < rows.length; i++) {
        const rowData = rows[i].split('\t')
        for (let j = 0; j < rowData.length; j++) {
          const pasteRow = selectedCell.r + i
          const pasteCol = selectedCell.c + j
          
          // Boundary check - skip cells outside grid
          if (pasteRow >= engine.rows || pasteCol >= engine.cols) continue
          
          const value = rowData[j].trim()
          engine.setCell(pasteRow, pasteCol, value)
        }
      }
      
      forceRerender()
      // Move selection to end of pasted range
      if (endRow < engine.rows && endCol < engine.cols) {
        setSelectedCell({ r: Math.min(endRow, engine.rows - 1), c: Math.min(endCol, engine.cols - 1) })
      }
    } catch (err) {
      console.warn('Paste from clipboard failed:', err)
    }
  }, [selectedCell, editingCell, engine, forceRerender])

  // Global keyboard listener for Ctrl+C and Ctrl+V
  useEffect(() => {
    const handleGlobalKeyDown = (e) => {
      // Ctrl+C (or Cmd+C on Mac)
      if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
        handleCopy(e)
      }
      // Ctrl+V (or Cmd+V on Mac)
      if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
        handlePaste(e)
      }
    }
    
    document.addEventListener('keydown', handleGlobalKeyDown)
    return () => document.removeEventListener('keydown', handleGlobalKeyDown)
  }, [handleCopy, handlePaste])

  // ────── Formatting toggles ──────

  const toggleBold = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { bold: !style.bold })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleItalic = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { italic: !style.italic })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleUnderline = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { underline: !style.underline })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const changeFontSize = useCallback((size) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { fontSize: size })
  }, [selectedCell, updateCellStyle])

  const changeAlignment = useCallback((align) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { align })
  }, [selectedCell, updateCellStyle])

  const changeFontColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { color })
  }, [selectedCell, updateCellStyle])

  const changeBackgroundColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { bg: color })
  }, [selectedCell, updateCellStyle])

  // ────── Clear operations ──────

  const clearSelectedCell = useCallback(() => {
    if (!selectedCell) return
    engine.setCell(selectedCell.r, selectedCell.c, '')
    forceRerender()
    // Remove style entry for cleared cell
    // Note: This deletes the style object entirely - if you need to preserve default styles,
    // you may want to set them explicitly rather than deleting
    const key = `${selectedCell.r},${selectedCell.c}`
    setCellStyles(prev => { const next = { ...prev }; delete next[key]; return next })
    setEditValue('')
  }, [selectedCell, engine, forceRerender])

  const clearAllCells = useCallback(() => {
    for (let r = 0; r < engine.rows; r++) {
      for (let c = 0; c < engine.cols; c++) {
        engine.setCell(r, c, '')
      }
    }
    forceRerender()
    setCellStyles({})
    setSelectedCell(null)
    setEditingCell(null)
    setEditValue('')
  }, [engine, forceRerender])

  // ────── Row / Column operations ──────

  const insertRow = useCallback(() => {
    if (!selectedCell) return
    engine.insertRow(selectedCell.r)
    forceRerender()
    setSelectedCell({ r: selectedCell.r + 1, c: selectedCell.c })
  }, [selectedCell, engine, forceRerender])

  const deleteRow = useCallback(() => {
    if (!selectedCell) return
    engine.deleteRow(selectedCell.r)
    forceRerender()
    if (selectedCell.r >= engine.rows) {
      setSelectedCell({ r: engine.rows - 1, c: selectedCell.c })
    }
  }, [selectedCell, engine, forceRerender])

  const insertColumn = useCallback(() => {
    if (!selectedCell) return
    engine.insertColumn(selectedCell.c)
    forceRerender()
    setSelectedCell({ r: selectedCell.r, c: selectedCell.c + 1 })
  }, [selectedCell, engine, forceRerender])

  const deleteColumn = useCallback(() => {
    if (!selectedCell) return
    engine.deleteColumn(selectedCell.c)
    forceRerender()
    if (selectedCell.c >= engine.cols) {
      setSelectedCell({ r: selectedCell.r, c: engine.cols - 1 })
    }
  }, [selectedCell, engine, forceRerender])

  // ────── Derived state ──────

  const selectedCellStyle = useMemo(() => {
    return selectedCell ? getCellStyle(selectedCell.r, selectedCell.c) : null
  }, [selectedCell, getCellStyle])

  const getColumnLabel = useCallback((col) => {
    let label = ''
    let num = col + 1
    while (num > 0) {
      num--
      label = String.fromCharCode(65 + (num % 26)) + label
      num = Math.floor(num / 26)
    }
    return label
  }, [])

  const selectedCellLabel = selectedCell
    ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}`
    : 'No cell'

  // Formula bar shows the raw formula text, not the computed value
  // When editing, show the current editValue; otherwise show the cell's raw content
  // Note: This is different from the cell display, which shows computed values
  const formulaBarValue = editingCell
    ? editValue
    : (selectedCell ? engine.getCell(selectedCell.r, selectedCell.c).raw : '')

  // ────── Render ──────

  return (
    <div className="app-wrapper">
      <div className="app-header">
        <h2 className="app-title">📊 Spreadsheet App</h2>
      </div>

      <div className="main-content">

        {/* ── Toolbar ── */}
        <div className="toolbar">
          <div className="toolbar-group">
            <button className={`toolbar-btn bold-btn ${selectedCellStyle?.bold ? 'active' : ''}`} onClick={toggleBold} title="Bold">B</button>
            <button className={`toolbar-btn italic-btn ${selectedCellStyle?.italic ? 'active' : ''}`} onClick={toggleItalic} title="Italic">I</button>
            <button className={`toolbar-btn underline-btn ${selectedCellStyle?.underline ? 'active' : ''}`} onClick={toggleUnderline} title="Underline">U</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Size:</span>
            <select className="toolbar-select" value={selectedCellStyle?.fontSize || 13} onChange={(e) => changeFontSize(parseInt(e.target.value))}>
              {[8, 10, 11, 12, 13, 14, 16, 18, 20, 24].map(s => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
          </div>

          <div className="toolbar-group">
            <button className={`align-btn ${selectedCellStyle?.align === 'left' ? 'active' : ''}`} onClick={() => changeAlignment('left')} title="Align Left">⬤←</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'center' ? 'active' : ''}`} onClick={() => changeAlignment('center')} title="Align Center">⬤</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'right' ? 'active' : ''}`} onClick={() => changeAlignment('right')} title="Align Right">⬤→</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Text:</span>
            <input
              type="color"
              value={selectedCellStyle?.color || '#000000'}
              onChange={(e) => changeFontColor(e.target.value)}
              title="Font color"
              style={{ width: '32px', height: '32px', border: '1px solid #dadce0', cursor: 'pointer', borderRadius: '4px' }}
            />
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Fill:</span>
            <select className="toolbar-select" value={selectedCellStyle?.bg || 'white'} onChange={(e) => changeBackgroundColor(e.target.value)}>
              <option value="white">White</option>
              <option value="#ffff99">Yellow</option>
              <option value="#99ffcc">Green</option>
              <option value="#ffcccc">Red</option>
              <option value="#cce5ff">Blue</option>
              <option value="#e0ccff">Purple</option>
              <option value="#ffd9b3">Orange</option>
              <option value="#f0f0f0">Gray</option>
            </select>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={handleUndo} disabled={!engine.canUndo()} title="Undo">↶ Undo</button>
            <button className="toolbar-btn" onClick={handleRedo} disabled={!engine.canRedo()} title="Redo">↷ Redo</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={insertRow} title="Insert Row">+ Row</button>
            <button className="toolbar-btn" onClick={deleteRow} title="Delete Row">- Row</button>
            <button className="toolbar-btn" onClick={insertColumn} title="Insert Column">+ Col</button>
            <button className="toolbar-btn" onClick={deleteColumn} title="Delete Column">- Col</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn danger" onClick={clearSelectedCell}>✕ Cell</button>
            <button className="toolbar-btn danger" onClick={clearAllCells}>✕ All</button>
            <button className="toolbar-btn" onClick={clearAllFilters} title="Clear all filters and sorting">⊗ Filters</button>
          </div>
        </div>

        {/* ── Formula Bar ── */}
        <div className="formula-bar">
          <span className="formula-bar-label">{selectedCellLabel}</span>
          <input
            className="formula-bar-input"
            value={formulaBarValue}
            onChange={(e) => handleFormulaBarChange(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            onFocus={handleFormulaBarFocus}
            placeholder="Select a cell then type, or enter a formula like =SUM(A1:A5)"
          />
        </div>

        {/* ── Grid ── */}
        <div className="grid-scroll">
          <table className="grid-table">
            <thead>
              <tr>
                <th className="col-header-blank"></th>
                {Array.from({ length: engine.cols }, (_, colIndex) => {
                  const sortState = columnSortState[colIndex]
                  const filterSet = columnFilters[colIndex]
                  const isFiltered = filterSet && filterSet.size > 0 && filterSet.size < getColumnUniqueValues(colIndex).length
                  const uniqueValues = getColumnUniqueValues(colIndex)

                  return (
                    <th key={colIndex} className="col-header">
                      <div className="col-header-content">
                        <span className="col-label">{getColumnLabel(colIndex)}</span>
                        <div className="col-header-actions">
                          <button
                            className={`sort-btn ${sortState ? 'active' : ''}`}
                            onClick={() => toggleColumnSort(colIndex)}
                            title={sortState ? `Sorted ${sortState}` : 'Click to sort'}
                            aria-label={`Sort ${getColumnLabel(colIndex)}`}
                          >
                            {sortState === 'asc' ? (
                              <ArrowUp size={16} strokeWidth={2} />
                            ) : sortState === 'desc' ? (
                              <ArrowDown size={16} strokeWidth={2} />
                            ) : (
                              <ArrowUpDown size={16} strokeWidth={2} />
                            )}
                          </button>
                          <div className="filter-dropdown-wrapper">
                            <button
                              ref={(el) => {
                                if (el) filterButtonRefs.current[colIndex] = el
                              }}
                              className={`filter-btn ${isFiltered ? 'active' : ''}`}
                              onClick={() => {
                                if (filterMenuOpen === colIndex) {
                                  setFilterMenuOpen(null)
                                } else {
                                  setFilterMenuOpen(colIndex)
                                  // Calculate position after menu opens
                                  setTimeout(() => calculateFilterMenuPosition(colIndex), 0)
                                }
                              }}
                              title={isFiltered ? 'Column is filtered' : 'Filter column'}
                              aria-label={`Filter ${getColumnLabel(colIndex)}`}
                            >
                              <Filter size={16} strokeWidth={2} />
                            </button>
                            {filterMenuOpen === colIndex && (
                              <div 
                                className="filter-menu"
                                style={{
                                  top: `${filterMenuPosition.top}px`,
                                  left: `${filterMenuPosition.left}px`
                                }}
                              >
                                <div className="filter-controls">
                                  <label className="filter-checkbox">
                                    <input
                                      type="checkbox"
                                      checked={!filterSet || filterSet.size === 0 || filterSet.size === uniqueValues.length}
                                      onChange={(e) => {
                                        if (e.target.checked) {
                                          clearColumnFilter(colIndex)
                                        }
                                      }}
                                      aria-label="Select all values"
                                    />
                                    <span>All</span>
                                  </label>
                                  {filterSet && filterSet.size > 0 && (
                                    <button className="filter-clear-btn" onClick={() => clearColumnFilter(colIndex)}>
                                      Clear
                                    </button>
                                  )}
                                </div>
                                <div className="filter-values">
                                  {uniqueValues.length === 0 ? (
                                    <div className="filter-empty">No values</div>
                                  ) : (
                                    uniqueValues.map(value => (
                                      <label key={value} className="filter-checkbox">
                                        <input
                                          type="checkbox"
                                          checked={!filterSet || filterSet.size === 0 || filterSet.has(value)}
                                          onChange={() => toggleFilterValue(colIndex, value)}
                                        />
                                        <span>{value}</span>
                                      </label>
                                    ))
                                  )}
                                </div>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    </th>
                  )
                })}
              </tr>
            </thead>
            <tbody>
              {visibleRows.map((rowIndex) => (
                <tr key={rowIndex}>
                  <td className="row-header">{rowIndex + 1}</td>
                  {Array.from({ length: engine.cols }, (_, colIndex) => {
                    const isSelected = selectedCell?.r === rowIndex && selectedCell?.c === colIndex
                    const isEditing = editingCell?.r === rowIndex && editingCell?.c === colIndex
                    const cellData = engine.getCell(rowIndex, colIndex)
                    const style = cellStyles[`${rowIndex},${colIndex}`] || {}
                    const displayValue = cellData.error
                      ? cellData.error
                      : (cellData.computed !== null && cellData.computed !== '' ? String(cellData.computed) : cellData.raw)

                    return (
                      <td
                        key={colIndex}
                        className={`cell ${isSelected ? 'selected' : ''}`}
                        style={{ background: style.bg || 'white' }}
                        onMouseDown={(e) => { e.preventDefault(); handleCellClick(rowIndex, colIndex) }}
                      >
                        {isEditing ? (
                          <input
                            autoFocus
                            className="cell-input"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onBlur={() => commitEdit(rowIndex, colIndex)}
                            onKeyDown={(e) => handleKeyDown(e, rowIndex, colIndex)}
                            ref={isSelected ? cellInputRef : undefined}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: style.color || '#202124',
                              fontSize: (style.fontSize || 13) + 'px',
                              textAlign: style.align || 'left',
                              background: style.bg || 'white',
                            }}
                          />
                        ) : (
                          <div
                            className={`cell-display align-${style.align || 'left'} ${cellData.error ? 'error' : ''}`}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: cellData.error ? '#d93025' : (style.color || '#202124'),
                              fontSize: (style.fontSize || 13) + 'px',
                            }}
                          >
                            {displayValue}
                          </div>
                        )}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <p className="footer-hint">
          Click a cell to edit · Enter/Tab/Arrow keys to navigate · Formulas: =A1+B1 · =SUM(A1:A5) · =AVG(A1:A5) · =MAX(A1:A5) · =MIN(A1:A5)
        </p>
      </div>
    </div>
  )
}
