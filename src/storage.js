/**
 * Local Storage Persistence Module
 * Handles saving and loading spreadsheet state with error handling
 */

const STORAGE_KEY = 'spreadsheet_state'
const STORAGE_VERSION = 1
const MAX_STORAGE_SIZE = 5 * 1024 * 1024 // 5MB limit

/**
 * Convert engine cells to serializable JSON format
 * The engine stores cells in a Map with Cell Key (e.g. "A1")
 * We need to convert this to a JSON-compatible format
 */
function serializeEngineData(engine) {
  const cellsArray = []
  
  // Iterate through all cells and collect non-empty ones
  for (let row = 0; row < engine.rows; row++) {
    for (let col = 0; col < engine.cols; col++) {
      const cellData = engine.getCell(row, col)
      if (cellData.raw || cellData.computed !== null) {
        cellsArray.push({
          r: row,
          c: col,
          raw: cellData.raw,
          computed: cellData.computed
        })
      }
    }
  }
  
  return {
    rows: engine.rows,
    cols: engine.cols,
    cells: cellsArray
  }
}

/**
 * Restore engine cells from serialized data
 * Does NOT restore undo/redo history (as per requirements)
 */
export function deserializeEngineData(engine, data) {
  if (!data || !Array.isArray(data.cells)) return false
  
  try {
    // Restore grid dimensions if they differ
    if (data.rows !== engine.rows || data.cols !== engine.cols) {
      // Note: Current engine doesn't have resize methods,
      // so we work with the default dimensions
      console.warn(`Saved dimensions (${data.rows}x${data.cols}) differ from current (${engine.rows}x${engine.cols})`)
    }
    
    // Restore all cells
    for (const cellData of data.cells) {
      if (cellData.r >= 0 && cellData.r < engine.rows && 
          cellData.c >= 0 && cellData.c < engine.cols) {
        engine.setCell(cellData.r, cellData.c, cellData.raw)
      }
    }
    
    return true
  } catch (err) {
    console.error('Error deserializing engine data:', err)
    return false
  }
}

/**
 * Main save function - persists all relevant state to localStorage
 * Does NOT save undo/redo history
 */
export function saveState(engine, cellStyles, columnSortState, columnFilters) {
  try {
    const state = {
      version: STORAGE_VERSION,
      timestamp: Date.now(),
      engineData: serializeEngineData(engine),
      cellStyles,
      columnSortState,
      // Convert Set to Array for JSON serialization
      columnFilters: Object.fromEntries(
        Object.entries(columnFilters).map(([col, set]) => [col, Array.from(set)])
      )
    }
    
    const jsonString = JSON.stringify(state)
    
    // Check storage size limit
    if (jsonString.length > MAX_STORAGE_SIZE) {
      console.warn('State too large to save to localStorage', {
        size: jsonString.length,
        limit: MAX_STORAGE_SIZE
      })
      return false
    }
    
    localStorage.setItem(STORAGE_KEY, jsonString)
    return true
  } catch (err) {
    // Handle QuotaExceededError and other storage errors gracefully
    if (err.name === 'QuotaExceededError') {
      console.warn('localStorage quota exceeded - state not saved', err)
    } else {
      console.error('Error saving state to localStorage:', err)
    }
    return false
  }
}

/**
 * Main load function - restores state from localStorage
 * Returns null if no saved state exists or if data is corrupted
 */
export function loadState() {
  try {
    const jsonString = localStorage.getItem(STORAGE_KEY)
    if (!jsonString) return null
    
    const state = JSON.parse(jsonString)
    
    // Version check for future compatibility
    if (state.version !== STORAGE_VERSION) {
      console.warn('Saved state version mismatch, data may be incompatible')
      return null
    }
    
    // Validate structure
    if (!state.engineData || !state.cellStyles || !state.columnSortState || !state.columnFilters) {
      console.warn('Saved state missing required fields')
      return null
    }
    
    // Convert columnFilters back from Arrays to Sets
    const columnFilters = Object.fromEntries(
      Object.entries(state.columnFilters).map(([col, arr]) => [col, new Set(arr)])
    )
    
    return {
      engineData: state.engineData,
      cellStyles: state.cellStyles,
      columnSortState: state.columnSortState,
      columnFilters: columnFilters,
      timestamp: state.timestamp
    }
  } catch (err) {
    if (err instanceof SyntaxError) {
      console.error('Corrupted localStorage data - clearing', err)
      // Clear corrupted data
      try {
        localStorage.removeItem(STORAGE_KEY)
      } catch (e) {
        console.error('Failed to clear corrupted data:', e)
      }
    } else {
      console.error('Error loading state from localStorage:', err)
    }
    return null
  }
}

/**
 * Clear all saved state (useful for testing or user reset)
 */
export function clearSavedState() {
  try {
    localStorage.removeItem(STORAGE_KEY)
    return true
  } catch (err) {
    console.error('Error clearing saved state:', err)
    return false
  }
}
