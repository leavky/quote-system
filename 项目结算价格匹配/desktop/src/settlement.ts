import { readFile, writeFile } from '@tauri-apps/plugin-fs'
import ExcelJS from 'exceljs'
import * as XLSX from 'xlsx'

export type CellValue = string | number | boolean | Date | null | undefined
export type PriorityEntry = { name: string; enabled: boolean }
export type PriceSystem = { name: string; column_index: number; non_empty: number }
export type ManualMatch = { row_id: string; price_sheet: string; price_row_number: number }
export type ManualSettlementEdit = {
  row_id: string
  unit_price: number | null
  quantity: number | null
  discount: number | null
  subtotal: number | null
  total: number | null
  unit_billing: number | null
}
export type Summary = {
  total: number
  matched: number
  unresolved: number
  total_amount: number
  status_counts: Record<string, number>
}
export type RecordRow = {
  id: string
  sheet: string
  row_number: number
  original: Record<string, CellValue>
  date: string
  report_number: string
  project_name: string
  unit_project: string
  client_name: string
  report_category: string
  billing_item: string
  quantity: number | null
  source_unit_price: number | null
  status: string
  matched_code: string
  price_sheet: string
  price_row_number: number | null
  selected_system: string
  selected_price: number | null
  settlement_amount: number | null
  available_prices: Record<string, number | null>
  candidate_count: number
  manual_match: boolean
}
export type MatchResult = {
  ledger_headers: string[]
  ledger_records: LedgerRecord[]
  price_items: PriceItem[]
  systems: PriceSystem[]
  priority: PriorityEntry[]
  manual_matches: ManualMatch[]
  manual_edits: ManualSettlementEdit[]
  records: RecordRow[]
  summary: Summary
}
export type MatchPayload = {
  ledger_name: string
  price_name: string
  systems: PriceSystem[]
  priority: PriorityEntry[]
  summary: Summary
  records: RecordRow[]
  result: MatchResult
}

type SheetData = { name: string; rows: CellValue[][] }
type LedgerRecord = Omit<
  RecordRow,
  | 'status'
  | 'matched_code'
  | 'price_sheet'
  | 'price_row_number'
  | 'selected_system'
  | 'selected_price'
  | 'settlement_amount'
  | 'available_prices'
  | 'candidate_count'
  | 'manual_match'
>
export type PriceItem = {
  sheet: string
  row_number: number
  sequence: string
  report_category: string
  billing_item: string
  code: string
  category: string
  prices: Record<string, number | null>
  raw_prices: Record<string, CellValue>
}

const LEDGER_HINTS = new Set(['委托日期', '报告编号', '报告类别', '工程名称', '计费项目', '数量'])
const PRICE_HINTS = new Set(['序号', '分类', '报告类别', '计费项目', '计费项目编号'])
const LEDGER_EXPORT_HEADERS = [
  '委托日期',
  '批准日期',
  '报告编号',
  '受理编号',
  '检测方案号',
  '报告类别',
  '账号编号',
  '工程名称',
  '单体工程',
  '委托单位',
  '进度',
  '委托书号',
  '送样人',
  '发放时间',
  '计费项目',
  '项目折扣(元)',
  '单价(元)',
  '数量',
  '小计',
  '折扣(元)',
  '总价',
  '单位计费',
  '已付',
  '是否付款',
  '备注',
  '备注',
  '所属质监站',
]

export async function buildMatchPayload(ledgerPath: string, pricePath: string): Promise<MatchPayload> {
  const result = await buildMatchResult(ledgerPath, pricePath)
  return {
    ledger_name: basename(ledgerPath),
    price_name: basename(pricePath),
    systems: result.systems,
    priority: result.priority,
    summary: result.summary,
    records: result.records,
    result,
  }
}

export async function loadSettlementPriceBook(path: string): Promise<{ items: PriceItem[]; systems: PriceSystem[] }> {
  return readPriceBook(path)
}

export async function updateSettlementPriceBook(
  path: string,
  updates: Array<{ sheet: string; row_number: number; prices: Record<string, number | null> }>,
): Promise<number> {
  const bytes = await readFile(path)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
  let changed = 0
  for (const update of updates) {
    const sheet = workbook.getWorksheet(update.sheet)
    if (!sheet) continue
    const headerRow = findExcelJsHeaderRow(sheet, PRICE_HINTS)
    if (!headerRow) continue
    const headers = new Map<string, number>()
    sheet.getRow(headerRow).eachCell((cell, index) => {
      const name = cellText(cell.value as CellValue)
      if (name) headers.set(name, index)
    })
    const row = sheet.getRow(update.row_number)
    let rowChanged = false
    Object.entries(update.prices).forEach(([system, value]) => {
      const column = headers.get(system)
      if (!column) return
      const cell = row.getCell(column)
      const next = value === null ? null : value
      if (cell.value !== next) {
        cell.value = next
        rowChanged = true
      }
    })
    if (rowChanged) changed += 1
  }
  if (changed) {
    const buffer = await workbook.xlsx.writeBuffer()
    await writeFile(path, new Uint8Array(buffer))
  }
  return changed
}

export function repricePayload(payload: MatchPayload, requestedPriority: PriorityEntry[]): MatchPayload {
  const priority = normalizePriority(payload.result.systems, requestedPriority)
  const records = matchRecords(payload.result.ledger_records, payload.result.price_items, priority, payload.result.manual_matches, payload.result.manual_edits)
  const result = { ...payload.result, priority, records, summary: summarize(records) }
  return { ...payload, priority, records, summary: result.summary, result }
}

export function applyManualMatch(payload: MatchPayload, rowId: string, priceItem: PriceItem, requestedPriority: PriorityEntry[]): MatchPayload {
  const manualMatches = [
    ...payload.result.manual_matches.filter((match) => match.row_id !== rowId),
    { row_id: rowId, price_sheet: priceItem.sheet, price_row_number: priceItem.row_number },
  ]
  const priority = normalizePriority(payload.result.systems, requestedPriority)
  const records = matchRecords(payload.result.ledger_records, payload.result.price_items, priority, manualMatches, payload.result.manual_edits)
  const result = { ...payload.result, priority, manual_matches: manualMatches, records, summary: summarize(records) }
  return { ...payload, priority, records, summary: result.summary, result }
}

export function applyManualSettlementEdit(payload: MatchPayload, edit: ManualSettlementEdit): MatchPayload {
  const manualEdits = [
    ...payload.result.manual_edits.filter((item) => item.row_id !== edit.row_id),
    edit,
  ]
  const records = matchRecords(payload.result.ledger_records, payload.result.price_items, payload.result.priority, payload.result.manual_matches, manualEdits)
  const result = { ...payload.result, manual_edits: manualEdits, records, summary: summarize(records) }
  return { ...payload, records, summary: result.summary, result }
}

export async function exportResult(payload: MatchPayload, outputPath: string, rowIds: string[]): Promise<number> {
  const selectedIds = new Set(rowIds)
  const records = payload.result.records.filter((record) => selectedIds.has(record.id))
  const columns = buildLedgerExportColumns(payload.result.ledger_headers)
  const manualEditIndex = new Map(payload.result.manual_edits.map((edit) => [edit.row_id, edit]))
  const workbook = new ExcelJS.Workbook()
  workbook.creator = '项目结算价格匹配'
  workbook.created = new Date()

  const detailSheet = workbook.addWorksheet('台账明细', {
    views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
  })

  columns.forEach((column, index) => {
    detailSheet.getRow(1).getCell(index + 1).value = column.label
  })
  detailSheet.getRow(1).height = 28
  detailSheet.getRow(1).eachCell((cell) => {
    cell.font = { name: 'Microsoft YaHei', size: 10, bold: true, color: { argb: 'FF17324D' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } }
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C4CE' } } }
  })

  const columnIndex = Object.fromEntries(columns.map((column, index) => [column.key, index + 1]))
  records.forEach((record, index) => {
    const rowIndex = index + 2
    const row = detailSheet.getRow(rowIndex)
    columns.forEach((column, exportColumnIndex) => {
      const cell = row.getCell(exportColumnIndex + 1)
      cell.value = safeExcelValue(record.original[column.key])
      if (cell.value instanceof Date) cell.numFmt = 'yyyy-mm-dd'
    })
    if (record.status === '已匹配' && record.selected_price !== null) {
      const manualEdit = manualEditIndex.get(record.id)
      const quantity = record.quantity ?? numberValue(record.original['数量'])
      const subtotal = manualEdit?.subtotal ?? (quantity !== null ? round2(record.selected_price * quantity) : null)
      const discount = manualEdit?.discount ?? numberValue(record.original['折扣(元)']) ?? 0
      const total = manualEdit?.total ?? (subtotal !== null ? round2(subtotal - discount) : null)
      setLedgerCell(row, columnIndex, '单价(元)', record.selected_price)
      setLedgerCell(row, columnIndex, '数量', quantity)
      setLedgerCell(row, columnIndex, '小计', subtotal)
      setLedgerCell(row, columnIndex, '折扣(元)', discount)
      setLedgerCell(row, columnIndex, '总价', total)
      setLedgerCell(row, columnIndex, '单位计费', manualEdit?.unit_billing ?? total)
    }
    row.eachCell((cell) => {
      cell.alignment = { vertical: 'middle' }
      cell.border = { bottom: { style: 'hair', color: { argb: 'FFE5E7EB' } } }
      if (record.status !== '已匹配') cell.fill = statusFill(record.status)
    })
  })

  for (const header of ['项目折扣(元)', '单价(元)', '数量', '小计', '折扣(元)', '总价', '单位计费', '已付']) {
    const index = columnIndex[header]
    if (!index) continue
    const column = detailSheet.getColumn(index)
    column.numFmt = '#,##0.00'
    column.alignment = { horizontal: 'right', vertical: 'middle' }
  }
  detailSheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: Math.max(1, records.length + 1), column: columns.length },
  }
  columns.forEach((column, index) => {
    detailSheet.getColumn(index + 1).width = columnWidth(column.label, records)
  })

  const buffer = await workbook.xlsx.writeBuffer()
  await writeFile(outputPath, new Uint8Array(buffer))
  return records.length
}

async function buildMatchResult(ledgerPath: string, pricePath: string): Promise<MatchResult> {
  const ledger = await readLedger(ledgerPath)
  const priceBook = await readPriceBook(pricePath)
  const priority = normalizePriority(priceBook.systems)
  const manualMatches: ManualMatch[] = []
  const manualEdits: ManualSettlementEdit[] = []
  const records = matchRecords(ledger.records, priceBook.items, priority, manualMatches, manualEdits)
  return {
    ledger_headers: ledger.headers,
    ledger_records: ledger.records,
    price_items: priceBook.items,
    systems: priceBook.systems,
    priority,
    manual_matches: manualMatches,
    manual_edits: manualEdits,
    records,
    summary: summarize(records),
  }
}

async function readSheets(path: string): Promise<SheetData[]> {
  const bytes = await readFile(path)
  const workbook = XLSX.read(bytes, { type: 'array', cellDates: true })
  return workbook.SheetNames.map((name) => ({
    name,
    rows: XLSX.utils.sheet_to_json<CellValue[]>(workbook.Sheets[name], { header: 1, raw: true, defval: null }),
  }))
}

async function readLedger(path: string): Promise<{ headers: string[]; records: LedgerRecord[] }> {
  const records: LedgerRecord[] = []
  const headers: string[] = []
  let recognizedSheets = 0

  for (const sheet of await readSheets(path)) {
    const headerIndex = findHeaderRow(sheet.rows, LEDGER_HINTS)
    if (headerIndex === null) continue
    recognizedSheets += 1
    const sheetHeaders = uniqueHeaders(sheet.rows[headerIndex])
    sheetHeaders.forEach((header) => {
      if (!headers.includes(header)) headers.push(header)
    })

    sheet.rows.slice(headerIndex + 1).forEach((values, offset) => {
      const rowNumber = headerIndex + offset + 2
      const original = rowObject(sheetHeaders, values)
      if (!Object.values(original).some((value) => cellText(value))) return
      const reportCategory = original['报告类别']
      const billingItem = original['计费项目']
      const reportNumber = original['报告编号'] || original['受理编号']
      if (![reportCategory, billingItem, reportNumber].some((value) => cellText(value))) return
      const date = asExcelDate(original['委托日期'])
      records.push({
        id: `${sheet.name}-${rowNumber}`,
        sheet: sheet.name,
        row_number: rowNumber,
        original,
        date: date ? toIsoDate(date) : '',
        report_number: cellText(reportNumber),
        project_name: cellText(original['工程名称']),
        unit_project: cellText(original['单体工程']),
        client_name: cellText(original['委托单位']),
        report_category: cellText(reportCategory),
        billing_item: cellText(billingItem),
        quantity: numberValue(original['数量']),
        source_unit_price: numberValue(original['单价(元)'] || original['单价（元）']),
      })
    })
  }

  if (!recognizedSheets) throw new Error('台账中未找到包含“报告类别”和“计费项目”的表头')
  return { headers, records }
}

async function readPriceBook(path: string): Promise<{ items: PriceItem[]; systems: PriceSystem[] }> {
  const items: PriceItem[] = []
  const systems: PriceSystem[] = []
  let recognizedSheets = 0

  for (const sheet of await readSheets(path)) {
    const headerIndex = findHeaderRow(sheet.rows, PRICE_HINTS)
    if (headerIndex === null) continue
    recognizedSheets += 1
    const rawHeaders = sheet.rows[headerIndex]
    const headers = uniqueHeaders(rawHeaders)
    if (headers.length < 6) throw new Error(`价格表 ${sheet.name} 的价格列不足，预期 F-J 列`)

    const sheetSystems: Array<[string, number]> = []
    for (let columnIndex = 5; columnIndex < Math.min(10, headers.length); columnIndex += 1) {
      const name = cellText(rawHeaders[columnIndex]) || `价格体系${columnIndex + 1}`
      if (!systems.some((system) => system.name === name)) systems.push({ name, column_index: columnIndex, non_empty: 0 })
      sheetSystems.push([name, columnIndex])
    }

    sheet.rows.slice(headerIndex + 1).forEach((values, offset) => {
      const rowNumber = headerIndex + offset + 2
      const row = rowObject(headers, values)
      const reportCategory = cellText(row['报告类别'])
      const billingItem = cellText(row['计费项目'])
      if (!reportCategory && !billingItem) return
      const prices: Record<string, number | null> = {}
      const rawPrices: Record<string, CellValue> = {}
      sheetSystems.forEach(([name, columnIndex]) => {
        const raw = values[columnIndex] ?? null
        prices[name] = numberValue(raw)
        rawPrices[name] = raw
        if (prices[name] !== null) systems.find((system) => system.name === name)!.non_empty += 1
      })
      items.push({
        sheet: sheet.name,
        row_number: rowNumber,
        sequence: cellText(row['序号']) || String(rowNumber),
        report_category: reportCategory,
        billing_item: billingItem,
        code: cellText(row['计费项目编号']),
        category: cellText(row['分类']),
        prices,
        raw_prices: rawPrices,
      })
    })
  }

  if (!recognizedSheets) throw new Error('价格汇总表中未找到包含“报告类别”和“计费项目”的表头')
  if (!systems.length) throw new Error('价格汇总表未识别到 F-J 价格体系')
  return { items, systems }
}

function matchRecords(
  ledgerRecords: LedgerRecord[],
  priceItems: PriceItem[],
  priority: PriorityEntry[],
  manualMatches: ManualMatch[],
  manualEdits: ManualSettlementEdit[] = [],
): RecordRow[] {
  const priceIndex = new Map<string, PriceItem[]>()
  priceItems.forEach((item) => {
    const key = pairKey(item.report_category, item.billing_item)
    priceIndex.set(key, [...(priceIndex.get(key) || []), item])
  })
  const enabledPriority = priority.filter((entry) => entry.enabled).map((entry) => entry.name)
  const manualIndex = new Map(manualMatches.map((match) => [match.row_id, match]))
  const editIndex = new Map(manualEdits.map((edit) => [edit.row_id, edit]))

  return ledgerRecords.map((source) => {
    const manualEdit = editIndex.get(source.id)
    const manualMatch = manualIndex.get(source.id)
    const manualItem = manualMatch
      ? priceItems.find((item) => item.sheet === manualMatch.price_sheet && item.row_number === manualMatch.price_row_number) || null
      : null
    const candidates = source.report_category && source.billing_item ? priceIndex.get(pairKey(source.report_category, source.billing_item)) || [] : []
    let selectedItem: PriceItem | null = null
    let selectedSystem = ''
    let selectedPrice: number | null = null
    let status = '未匹配'

    if (manualItem) {
      selectedItem = manualItem
      for (const systemName of enabledPriority) {
        const price = selectedItem.prices[systemName]
        if (price !== null && price !== undefined) {
          selectedSystem = systemName
          selectedPrice = price
          break
        }
      }
      status = selectedPrice !== null ? '已匹配' : '无可用价格'
    } else if (!exactKey(source.report_category) || !exactKey(source.billing_item)) {
      status = '字段缺失'
    } else if (!candidates.length) {
      status = '未匹配'
    } else if (candidates.length > 1) {
      status = '价格表重复'
    } else {
      selectedItem = candidates[0]
      for (const systemName of enabledPriority) {
        const price = selectedItem.prices[systemName]
        if (price !== null && price !== undefined) {
          selectedSystem = systemName
          selectedPrice = price
          break
        }
      }
      status = selectedPrice !== null ? '已匹配' : '无可用价格'
    }

    if (manualEdit) {
      selectedSystem = '人工录入'
      selectedPrice = manualEdit.unit_price
      status = '已匹配'
    }

    const quantity = manualEdit ? manualEdit.quantity : source.quantity
    const amount = manualEdit
      ? manualEdit.total
      : selectedPrice !== null && quantity !== null ? round2(selectedPrice * quantity) : null
    return {
      ...source,
      status,
      matched_code: manualEdit ? '人工录入' : selectedItem?.code || '',
      price_sheet: selectedItem?.sheet || '',
      price_row_number: selectedItem?.row_number || null,
      selected_system: selectedSystem,
      selected_price: selectedPrice,
      settlement_amount: amount,
      available_prices: selectedItem?.prices || {},
      candidate_count: candidates.length,
      manual_match: Boolean(manualItem || manualEdit),
      quantity,
    }
  })
}

function normalizePriority(systems: PriceSystem[], requested: PriorityEntry[] = []): PriorityEntry[] {
  const known = new Set(systems.map((system) => system.name))
  const result: PriorityEntry[] = []
  requested.forEach((entry) => {
    if (known.has(entry.name) && !result.some((item) => item.name === entry.name)) result.push({ name: entry.name, enabled: Boolean(entry.enabled) })
  })
  systems.forEach((system) => {
    if (!result.some((entry) => entry.name === system.name)) result.push({ name: system.name, enabled: true })
  })
  if (!result.some((entry) => entry.enabled)) throw new Error('至少启用一个价格体系')
  return result
}

function summarize(records: RecordRow[]): Summary {
  return records.reduce(
    (summary, record) => {
      summary.total += 1
      summary.status_counts[record.status] = (summary.status_counts[record.status] || 0) + 1
      if (record.status === '已匹配') summary.matched += 1
      else summary.unresolved += 1
      if (record.settlement_amount !== null) summary.total_amount = round2(summary.total_amount + record.settlement_amount)
      return summary
    },
    { total: 0, matched: 0, unresolved: 0, total_amount: 0, status_counts: {} } as Summary,
  )
}

function cellText(value: CellValue): string {
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return value.toISOString().slice(0, 19).replace('T', ' ')
  if (typeof value === 'number' && Number.isInteger(value)) return String(value)
  return String(value).trim()
}

function exactKey(value: CellValue): string {
  return cellText(value).normalize('NFKC').replace(/\s+/g, '').trim()
}

function pairKey(reportCategory: CellValue, billingItem: CellValue): string {
  return `${exactKey(reportCategory)}\u0000${exactKey(billingItem)}`
}

function numberValue(value: CellValue): number | null {
  if (value === null || value === undefined || typeof value === 'boolean') return null
  if (typeof value === 'number') return Number.isNaN(value) ? null : value
  const text = cellText(value).normalize('NFKC').replace(/,/g, '')
  return /^-?\d+(?:\.\d+)?$/.test(text) ? Number(text) : null
}

function uniqueHeaders(values: CellValue[]): string[] {
  const counts = new Map<string, number>()
  return values.map((value, index) => {
    const base = cellText(value) || `未命名${index + 1}`
    const count = (counts.get(base) || 0) + 1
    counts.set(base, count)
    return count === 1 ? base : `${base}_${count}`
  })
}

function rowObject(headers: string[], values: CellValue[]): Record<string, CellValue> {
  return Object.fromEntries(headers.map((header, index) => [header, values[index] ?? null]))
}

function findHeaderRow(rows: CellValue[][], hints: Set<string>): number | null {
  let bestIndex: number | null = null
  let bestScore = 0
  rows.slice(0, 20).forEach((row, index) => {
    const values = new Set(row.map(cellText).filter(Boolean))
    const score = Array.from(values).filter((value) => hints.has(value)).length
    if (score > bestScore) {
      bestScore = score
      bestIndex = index
    }
  })
  return bestScore >= 2 ? bestIndex : null
}

function findExcelJsHeaderRow(sheet: ExcelJS.Worksheet, hints: Set<string>): number | null {
  let bestRow: number | null = null
  let bestScore = 0
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 20); rowNumber += 1) {
    const values = sheet.getRow(rowNumber).values as CellValue[]
    const score = values.map(cellText).filter((value) => hints.has(value)).length
    if (score > bestScore) {
      bestScore = score
      bestRow = rowNumber
    }
  }
  return bestScore >= 2 ? bestRow : null
}

function asExcelDate(value: CellValue): Date | null {
  if (value instanceof Date) return value
  if (typeof value === 'number' && value >= 1 && value <= 100000) return new Date(Math.round((value - 25569) * 86400 * 1000))
  const text = cellText(value)
  const match = text.match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})/)
  if (!match) return null
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]))
}

function toIsoDate(value: Date): string {
  const year = value.getFullYear()
  const month = String(value.getMonth() + 1).padStart(2, '0')
  const day = String(value.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

function safeExcelValue(value: CellValue): ExcelJS.CellValue {
  if (typeof value === 'string' && /^[=+\-@]/.test(value)) return `'${value}`
  if (value === undefined) return null
  return value as ExcelJS.CellValue
}

function buildLedgerExportColumns(headers: string[]): Array<{ key: string; label: string }> {
  const available = new Set(headers)
  return LEDGER_EXPORT_HEADERS.map((label, index) => {
    if (label !== '备注') return { key: available.has(label) ? label : label, label }
    const remarkKey = index === LEDGER_EXPORT_HEADERS.indexOf('备注') ? '备注' : '备注_2'
    return { key: available.has(remarkKey) ? remarkKey : remarkKey, label }
  })
}

function setLedgerCell(
  row: ExcelJS.Row,
  columnIndex: Record<string, number>,
  header: string,
  value: CellValue,
): void {
  const index = columnIndex[header]
  if (!index) return
  row.getCell(index).value = safeExcelValue(value)
}

function statusFill(status: string): ExcelJS.Fill {
  const colors: Record<string, string> = {
    已匹配: 'FFDCFCE7',
    未匹配: 'FFFEE2E2',
    字段缺失: 'FFFEF3C7',
    价格表重复: 'FFFFD7AA',
    无可用价格: 'FFEDE9FE',
  }
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: colors[status] || 'FFFFFFFF' } }
}

function columnWidth(header: string, records: RecordRow[]): number {
  const defaults: Record<string, number> = {
    委托日期: 12,
    报告编号: 20,
    受理编号: 20,
    报告类别: 18,
    工程名称: 38,
    单体工程: 18,
    委托单位: 28,
    计费项目: 32,
    匹配状态: 14,
    计费项目编号: 16,
    匹配价格体系: 16,
  }
  if (defaults[header]) return defaults[header]
  const values = [header, ...records.slice(0, 200).map((record) => cellText(record.original[header]))]
  return Math.min(Math.max(...values.map((value) => value.length)) + 2, 36)
}

function round2(value: number): number {
  return Math.round((value + Number.EPSILON) * 100) / 100
}

function basename(path: string) {
  return path.split(/[\\/]/).pop() || path
}
