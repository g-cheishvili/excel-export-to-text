// Cell-reference helpers and the logic that turns a sheet grid into `items`.

export type CellValue = string | number | boolean | null

export interface CellRef {
  row: number // 0-indexed (A1 -> row 0)
  col: number // 0-indexed (A1 -> col 0)
}

/** "A" -> 0, "Z" -> 25, "AA" -> 26 */
export function colToIndex(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64)
  }
  return n - 1
}

/** 0 -> "A", 25 -> "Z", 26 -> "AA" */
export function indexToCol(index: number): string {
  let n = index + 1
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

/** Parse "A1" / "B12" into a zero-based {row, col}. Returns null if malformed. */
export function parseCellRef(ref: string): CellRef | null {
  const m = /^\s*([A-Za-z]+)\s*(\d+)\s*$/.exec(ref)
  if (!m) return null
  return { col: colToIndex(m[1]), row: parseInt(m[2], 10) - 1 }
}

export function formatCellRef(ref: CellRef): string {
  return `${indexToCol(ref.col)}${ref.row + 1}`
}

/** Normalize a header label to English snake_case. */
export function toSnakeCase(input: unknown): string {
  const s = String(input ?? '')
    .normalize('NFKD')
    .replace(/\p{M}/gu, '') // strip combining diacritical marks
  return s
    .replace(/([a-z0-9])([A-Z])/g, '$1_$2') // split camelCase before lowercasing
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toLowerCase()
}

/** Build snake_case keys from a header row, with fallbacks and de-duplication. */
function makeKeys(headerCells: CellValue[]): string[] {
  const used = new Map<string, number>()
  return headerCells.map((cell, i) => {
    let key = toSnakeCase(cell)
    if (!key) key = `col_${i + 1}`
    const seen = used.get(key) ?? 0
    used.set(key, seen + 1)
    return seen === 0 ? key : `${key}_${seen + 1}`
  })
}

export interface ColumnInfo {
  key: string // snake_case key used on each item
  header: CellValue // original heading-row label
  col: number // zero-based grid column index
  isLink?: boolean // true for the synthesized "<key>_url" hyperlink column
}

export interface BuiltItems {
  keys: string[]
  columns: ColumnInfo[]
  headerRow: number // 1-based sheet row treated as the heading row
  items: Record<string, CellValue>[]
}

/**
 * Slice `grid` to the inclusive [start, end] range, treat the first row as
 * headers and the rest as records. Fully empty data rows are dropped.
 */
export function buildItems(
  grid: CellValue[][],
  start: CellRef,
  end: CellRef,
  links?: (string | null)[][],
): BuiltItems {
  const r0 = Math.min(start.row, end.row)
  const r1 = Math.max(start.row, end.row)
  const c0 = Math.min(start.col, end.col)
  const c1 = Math.max(start.col, end.col)

  const headerCells: CellValue[] = []
  for (let c = c0; c <= c1; c++) headerCells.push(grid[r0]?.[c] ?? null)
  const keys = makeKeys(headerCells)

  // A column gets a companion "<key>_url" key if any of its data cells is a
  // hyperlink. The url key is de-duplicated against the real column keys.
  const taken = new Set(keys)
  const linkKeys = keys.map((key, i) => {
    const hasLink =
      !!links &&
      (() => {
        for (let r = r0 + 1; r <= r1; r++) if (links[r]?.[c0 + i]) return true
        return false
      })()
    if (!hasLink) return null
    let name = `${key}_url`
    while (taken.has(name)) name = `${name}_`
    taken.add(name)
    return name
  })

  const columns: ColumnInfo[] = []
  keys.forEach((key, i) => {
    columns.push({ key, header: headerCells[i], col: c0 + i })
    if (linkKeys[i]) {
      columns.push({ key: linkKeys[i]!, header: headerCells[i], col: c0 + i, isLink: true })
    }
  })

  const items: Record<string, CellValue>[] = []
  for (let r = r0 + 1; r <= r1; r++) {
    const row = grid[r] ?? []
    const obj: Record<string, CellValue> = {}
    let hasValue = false
    keys.forEach((key, i) => {
      const v = row[c0 + i] ?? null
      if (v !== null && v !== '') hasValue = true
      obj[key] = v
      if (linkKeys[i]) obj[linkKeys[i]!] = links?.[r]?.[c0 + i] ?? null
    })
    if (hasValue) items.push(obj)
  }

  return { keys, columns, headerRow: r0 + 1, items }
}

/** Used range of a grid as start/end refs. start is always A1. */
export function usedRange(grid: CellValue[][]): { start: CellRef; end: CellRef } {
  const rows = grid.length
  let cols = 0
  for (const row of grid) if (row && row.length > cols) cols = row.length
  return {
    start: { row: 0, col: 0 },
    end: { row: Math.max(0, rows - 1), col: Math.max(0, cols - 1) },
  }
}
