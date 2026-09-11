// Dev helper: dump the first N non-empty rows of every tab in a workbook.
//   node tools/peek.js <file.xlsx> [rows] [tabFilter]
import ExcelJS from 'exceljs'

const clean = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim()
function txt(c) {
  const v = c == null ? null : c.value
  if (v == null) return ''
  if (v instanceof Date) return v.toISOString().slice(0, 16).replace('T', ' ')
  if (typeof v !== 'object') return clean(v)
  if (Array.isArray(v.richText)) return clean(v.richText.map((r) => r.text).join(''))
  if ('result' in v) { const r = v.result; return r == null || (typeof r === 'object' && !(r instanceof Date)) ? '' : txt({ value: r }) }
  if ('text' in v) return txt({ value: v.text })
  return ''
}

const [file, n = '8', filter = ''] = process.argv.slice(2)
const wb = new ExcelJS.Workbook()
await wb.xlsx.readFile(file)
for (const ws of wb.worksheets) {
  if (filter && !ws.name.toLowerCase().includes(filter.toLowerCase())) continue
  let lastRow = 0, lastCol = 0
  ws.eachRow({ includeEmpty: false }, (row, i) => {
    let any = false
    row.eachCell({ includeEmpty: false }, (c, j) => { if (txt(c)) { any = true; if (j > lastCol) lastCol = j } })
    if (any) lastRow = i
  })
  console.log(`\n================ ${ws.name}  (${lastRow} rows x ${lastCol} cols) ================`)
  let shown = 0
  for (let i = 1; i <= lastRow && shown < Number(n); i++) {
    const row = ws.getRow(i)
    const cells = []
    for (let j = 1; j <= lastCol; j++) cells.push(txt(row.getCell(j)))
    while (cells.length && !cells[cells.length - 1]) cells.pop()
    if (!cells.length) continue
    // collapse merged banners
    const uniq = new Set(cells.filter(Boolean))
    const line = uniq.size === 1 && cells.filter(Boolean).length > 1 ? '[banner] ' + [...uniq][0] : cells.join(' | ')
    console.log(String(i).padStart(3) + ' | ' + line.slice(0, 300))
    shown++
  }
}
