import { useMemo, useState } from 'react'
import {
  buildItems,
  formatCellRef,
  indexToCol,
  parseCellRef,
  usedRange,
} from './lib/cells'
import { parseSpreadsheet, type ParsedSheet } from './lib/spreadsheet'
import { renderTemplate } from './lib/render'
import './App.css'

const DEFAULT_TEMPLATE = `<ul>
<% items.forEach(function (item) { %>
  <li><%= JSON.stringify(item) %></li>
<% }); %>
</ul>`

const ACCEPT = '.xlsx,.xls,.csv'

export default function App() {
  const [sheets, setSheets] = useState<ParsedSheet[]>([])
  const [sheetIndex, setSheetIndex] = useState(0)
  const [startCell, setStartCell] = useState('A1')
  const [endCell, setEndCell] = useState('A1')
  const [template, setTemplate] = useState(DEFAULT_TEMPLATE)
  const [output, setOutput] = useState('')
  const [view, setView] = useState<'preview' | 'html'>('preview')
  const [error, setError] = useState('')

  const grid = sheets[sheetIndex]?.grid ?? null

  // გასაღებებისა და სტრიქონების ცოცხალი გადახედვა მიმდინარე დიაპაზონისთვის.
  const preview = useMemo(() => {
    if (!grid) return null
    const start = parseCellRef(startCell)
    const end = parseCellRef(endCell)
    if (!start || !end) return null
    return buildItems(grid, start, end)
  }, [grid, startCell, endCell])

  async function onFile(file: File) {
    setError('')
    setOutput('')
    try {
      const parsed = parseSpreadsheet(await file.arrayBuffer())
      if (parsed.length === 0) throw new Error('ფაილში ფურცლები ვერ მოიძებნა.')
      setSheets(parsed)
      selectSheet(parsed, 0)
    } catch (e) {
      setSheets([])
      setError(`ფაილის წაკითხვა ვერ მოხერხდა: ${(e as Error).message}`)
    }
  }

  function selectSheet(parsed: ParsedSheet[], index: number) {
    setSheetIndex(index)
    const range = usedRange(parsed[index].grid)
    setStartCell(formatCellRef(range.start))
    setEndCell(formatCellRef(range.end))
  }

  function generate() {
    setError('')
    if (!preview) {
      setError('ატვირთეთ ფაილი და მიუთითეთ სწორი დიაპაზონი (მაგ. A1 და X32).')
      return
    }
    try {
      setOutput(renderTemplate(template, preview.items))
    } catch (e) {
      setOutput('')
      setError(`შაბლონის შეცდომა: ${(e as Error).message}`)
    }
  }

  return (
    <main className="app">
      <header>
        <h1>Excel → EJS</h1>
        <p className="sub">
          ატვირთეთ ცხრილი, აირჩიეთ ცხრილის დიაპაზონი და დააგენერირეთ შედეგი EJS
          შაბლონით. ყველაფერი თქვენს ბრაუზერში მუშაობს — მონაცემები არსად
          იგზავნება.
        </p>
      </header>

      <section className="card">
        <label className="field">
          <span className="label">ცხრილის ფაილი</span>
          <input
            type="file"
            accept={ACCEPT}
            onChange={(e) => {
              const f = e.target.files?.[0]
              if (f) void onFile(f)
            }}
          />
          <span className="hint">მიიღება .xlsx, .xls და .csv</span>
        </label>

        {sheets.length > 0 && (
          <>
            {sheets.length > 1 && (
              <label className="field">
                <span className="label">ფურცელი</span>
                <select
                  value={sheetIndex}
                  onChange={(e) => selectSheet(sheets, Number(e.target.value))}
                >
                  {sheets.map((s, i) => (
                    <option key={i} value={i}>
                      {s.name}
                    </option>
                  ))}
                </select>
              </label>
            )}

            <div className="range">
              <label className="field">
                <span className="label">საწყისი უჯრა</span>
                <input
                  value={startCell}
                  onChange={(e) => setStartCell(e.target.value.toUpperCase())}
                  placeholder="A1"
                />
              </label>
              <span className="range-sep">→</span>
              <label className="field">
                <span className="label">ბოლო უჯრა</span>
                <input
                  value={endCell}
                  onChange={(e) => setEndCell(e.target.value.toUpperCase())}
                  placeholder="X32"
                />
              </label>
            </div>

            {preview && (
              <p className="meta">
                <strong>{preview.items.length}</strong> სტრიქონი ·{' '}
                <strong>{preview.keys.length}</strong> სვეტი ·{' '}
                <span className="keys">{preview.keys.join(', ')}</span>
              </p>
            )}
          </>
        )}

        <label className="field">
          <span className="label">EJS შაბლონი</span>
          <textarea
            className="template"
            spellCheck={false}
            value={template}
            onChange={(e) => setTemplate(e.target.value)}
            rows={10}
          />
          <span className="hint">
            ხელმისაწვდომი ცვლადი: <code>items</code> — ობიექტების მასივი ფორმით{' '}
            <code>{'{ სვეტი: მნიშვნელობა }'}</code>.
          </span>
        </label>

        <details className="help" open>
          <summary>როგორ დავწერო EJS შაბლონი?</summary>

          <h3>ხელმისაწვდომი ცვლადები</h3>
          <ul>
            <li>
              <code>items</code> — ყველა მონაცემთა სტრიქონის მასივი. თითოეული
              ელემენტი არის ობიექტი, რომლის გასაღებებიც სათაურის სტრიქონიდან
              მოდის (snake_case).
            </li>
            <li>
              ციკლის შიგნით <code>item</code> — მიმდინარე სტრიქონი. მნიშვნელობას
              იღებთ <code>item.გასაღები</code>-ით (გასაღებების სია იხილეთ ქვემოთ).
            </li>
            <li>
              <code>items.length</code> — სტრიქონების რაოდენობა.
            </li>
          </ul>

          <h3>ტეგების სახეები</h3>
          <ul>
            <li>
              <code>{'<%= მნიშვნელობა %>'}</code> — ბეჭდავს მნიშვნელობას; HTML
              სიმბოლოები უსაფრთხოდ ჩაანაცვლდება (escaped).
            </li>
            <li>
              <code>{'<%- მნიშვნელობა %>'}</code> — ბეჭდავს მნიშვნელობას როგორც
              raw HTML (არ ეკრანირდება).
            </li>
            <li>
              <code>{'<% კოდი %>'}</code> — JavaScript კოდი ბეჭდვის გარეშე
              (ციკლი, პირობა).
            </li>
            <li>
              <code>{'<%# კომენტარი %>'}</code> — კომენტარი, შედეგში არ ჩანს.
            </li>
          </ul>

          <h3>ყველა სტრიქონის გავლა</h3>
          <pre className="help-code">{`<% items.forEach(function (item) { %>
  <%= item.name %> — <%= item.price %>
<% }); %>`}</pre>

          <h3>პირობა</h3>
          <pre className="help-code">{`<% if (item.price > 100) { %>
  ძვირადღირებული
<% } else { %>
  იაფი
<% } %>`}</pre>

          <h3>ახალი ხაზი</h3>
          <p>
            ახალ ხაზზე გადასასვლელად ჩაწერეთ <code>{'<br>'}</code> იქ, სადაც
            გსურთ ხაზის გატეხვა. შაბლონში უბრალო <kbd>Enter</kbd> შედეგში ახალ
            ხაზს <strong>არ</strong> ქმნის.
          </p>
          <p className="help-io">შაბლონი:</p>
          <pre className="help-code">{`<%= item.name %><br>
<%= item.city %>`}</pre>
          <p className="help-io">შედეგი:</p>
          <pre className="help-code">{`გიორგი
თბილისი`}</pre>

          <h3>ინდექსი (ნომერი)</h3>
          <pre className="help-code">{`<% items.forEach(function (item, i) { %>
  <%= i + 1 %>. <%= item.name %><br>
<% }); %>`}</pre>
        </details>

        {preview && preview.columns.length > 0 && (
          <div className="legend">
            <p className="legend-title">
              სტრიქონი <strong>{preview.headerRow}</strong> არის სათაურების
              სტრიქონი. ქვემოთ თითოეული სტრიქონი იქცევა ერთ <code>item</code>-ად,
              შემდეგი გასაღებებით:
            </p>
            <ul className="legend-list">
              {preview.columns.map((c) => (
                <li key={c.key}>
                  <code>item.{c.key}</code>
                  <span className="legend-eq">=</span>
                  <span className="legend-src">
                    მნიშვნელობა სვეტიდან {indexToCol(c.col)} —{' '}
                    {c.header === null || c.header === ''
                      ? <em>ცარიელი სათაური</em>
                      : `„${c.header}“`}
                  </span>
                </li>
              ))}
            </ul>
          </div>
        )}

        <button className="generate" type="button" onClick={generate}>
          გენერაცია
        </button>

        {error && <p className="error">{error}</p>}
      </section>

      {output && (
        <section className="card">
          <div className="output-head">
            <h2>შედეგი</h2>
            <div className="output-actions">
              <div className="toggle" role="tablist" aria-label="შედეგის ხედი">
                <button
                  type="button"
                  role="tab"
                  aria-selected={view === 'preview'}
                  className={view === 'preview' ? 'active' : ''}
                  onClick={() => setView('preview')}
                >
                  გადახედვა
                </button>
                <button
                  type="button"
                  role="tab"
                  aria-selected={view === 'html'}
                  className={view === 'html' ? 'active' : ''}
                  onClick={() => setView('html')}
                >
                  HTML
                </button>
              </div>
              <button
                type="button"
                className="copy"
                onClick={() => void navigator.clipboard?.writeText(output)}
              >
                კოპირება
              </button>
            </div>
          </div>
          {view === 'preview' ? (
            <iframe
              className="preview"
              title="HTML გადახედვა"
              sandbox=""
              srcDoc={output}
            />
          ) : (
            <pre className="output">{output}</pre>
          )}
        </section>
      )}
    </main>
  )
}
