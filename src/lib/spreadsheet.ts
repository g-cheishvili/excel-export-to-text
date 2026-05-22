// Spreadsheet reading is delegated to SheetJS (xlsx), which auto-detects the
// format (.xlsx / .xls / .csv / ...) and handles the long tail of edge cases.
// We only adapt its output into an A1-aligned grid of plain CellValues.

import * as XLSX from 'xlsx'
import type { CellValue } from './cells'

export interface ParsedSheet {
  name: string
  grid: CellValue[][]
  // Hyperlink targets, aligned 1:1 with `grid`. null where a cell has no link.
  links: (string | null)[][]
}

export function parseSpreadsheet(buffer: ArrayBuffer): ParsedSheet[] {
  const wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true })
  return wb.SheetNames.map((name) => ({
    name,
    ...sheetToData(wb.Sheets[name]),
  }))
}

/**
 * Convert a worksheet to dense, A1-anchored 2D arrays of values and hyperlink
 * targets. We force the read range to start at A1 so `grid[row][col]` always
 * matches the cell reference the user types (e.g. C2), regardless of where the
 * sheet's used range actually begins.
 */
function sheetToData(ws: XLSX.WorkSheet): {
  grid: CellValue[][]
  links: (string | null)[][]
} {
  const ref = ws['!ref']
  if (!ref) return { grid: [], links: [] }
  const range = XLSX.utils.decode_range(ref)
  range.s.r = 0
  range.s.c = 0

  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, {
    header: 1,
    raw: true,
    defval: null,
    blankrows: true,
    range,
  })
  const grid = rows.map((row) => (row ?? []).map(toCellValue))

  // Hyperlinks live on the cell object (`cell.l.Target`), not in sheet_to_json,
  // so walk the same range by address to pull them out.
  const links: (string | null)[][] = []
  for (let r = 0; r <= range.e.r; r++) {
    const linkRow: (string | null)[] = []
    for (let c = 0; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined
      linkRow.push(cell?.l?.Target ?? null)
    }
    links.push(linkRow)
  }

  return { grid, links }
}

function toCellValue(value: unknown): CellValue {
  if (value == null) return null
  if (value instanceof Date) return dateToISO(value)
  if (typeof value === 'number' || typeof value === 'boolean' || typeof value === 'string') {
    return value
  }
  return String(value)
}

/** ISO date (or date-time when a time component is present). */
function dateToISO(d: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0')
  const date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
  if (d.getHours() === 0 && d.getMinutes() === 0 && d.getSeconds() === 0) return date
  return `${date}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`
}
