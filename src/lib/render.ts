// Thin wrapper around EJS so it can run entirely in the browser.

import ejs from 'ejs'
import type { CellValue } from './cells'

export function renderTemplate(
  template: string,
  items: Record<string, CellValue>[],
): string {
  return ejs.render(template, { items }, { rmWhitespace: false })
}
