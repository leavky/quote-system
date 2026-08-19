import { readFile, writeFile } from '@tauri-apps/plugin-fs'
import ExcelJS from 'exceljs'
import * as XLSX from 'xlsx'

type CellValue = string | number | boolean | Date | null | undefined
type SheetData = { name: string; rows: CellValue[][] }

export type QuoteStatus = '自动匹配' | '待确认' | '未匹配'

export type QuotePriceItem = {
  id: string
  sheet: string
  row_number: number
  seq: string
  category: string
  material: string
  parameter: string
  unit: string
  price: number | null
  raw_price: string
  remark: string
  code: string
  project_aliases: string[]
  parameter_aliases: string[]
  aliases: string[]
  search_text: string
  price_rule: null
}

export type QuoteLine = {
  id: string
  sheet: string
  row_number: number
  seq: string
  sample_name: string
  project_name: string
  parameter: string
  unit: string
  quantity: number | null
  source_price: number | null
  source_total: number | null
  remark: string
  search_text: string
}

export type QuoteCandidate = {
  score: number
  method: string
  item: QuotePriceItem
}

export type QuoteMatchRow = QuoteLine & {
  match_status: QuoteStatus
  match_score: number
  match_method: string
  matched: QuotePriceItem | null
  top_candidates: QuoteCandidate[]
  matched_price: number | null
  matched_price_text: string | null
  matched_code: string | null
  matched_label: string | null
  calculated_total: number | null
  price_rule_result: null
  price_explanation: string
  matched_remark?: string
  alias_learned?: boolean
  alias_message?: string
  manual_confirmed?: boolean
}

export type QuoteMatchPayload = {
  quote_name: string
  price_name: string
  price_items: QuotePriceItem[]
  tabs: string[]
  matches: QuoteMatchRow[]
  summary: {
    price_items: number
    quote_lines: number
    auto: number
    review: number
    unmatched: number
    manual: number
  }
}

const PRICE_COLUMNS = new Set([
  '序号',
  '检测项目',
  '检测材料',
  '检测参数',
  '单位',
  '单价（元）',
  '单价',
  '备注',
  '报价编号',
  '检测项目别名',
  '检测参数别名',
])
const QUOTE_HINT_COLUMNS = new Set(['序号', '样品名称', '检测材料', '检测项目', '检测参数', '具体检测项目', '组/点数', '备注'])
const STOP_WORDS = ['检测', '试验', '项目', '参数', '性能', '含量', '测定']

export async function buildQuoteMatchPayload(quotePath: string, pricePath: string): Promise<QuoteMatchPayload> {
  const priceItems = await readQuotePriceBook(pricePath)
  const quoteLines = await readQuoteBook(quotePath)
  const matches = quoteLines.map((line) => matchQuoteLine(line, priceItems))
  return makePayload(quotePath, pricePath, priceItems, matches)
}

export async function loadQuotePriceBook(path: string): Promise<QuotePriceItem[]> {
  return readQuotePriceBook(path)
}

export async function updateQuotePriceBook(
  path: string,
  updates: Array<{ sheet: string; row_number: number; price: number | null; remark?: string }>,
): Promise<number> {
  const bytes = await readFile(path)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
  let changed = 0
  for (const update of updates) {
    const sheet = workbook.getWorksheet(update.sheet)
    if (!sheet) continue
    const located = findHeaderCells(sheet, PRICE_COLUMNS)
    if (!located) continue
    const row = sheet.getRow(update.row_number)
    const priceColumn = located.headers.get('单价（元）') || located.headers.get('单价')
    const remarkColumn = located.headers.get('备注')
    let rowChanged = false
    if (priceColumn) {
      const cell = row.getCell(priceColumn)
      const next = update.price === null ? null : update.price
      if (cell.value !== next) {
        cell.value = next
        rowChanged = true
      }
    }
    if (remarkColumn && update.remark !== undefined) {
      const cell = row.getCell(remarkColumn)
      if (cell.value !== update.remark) {
        cell.value = update.remark
        rowChanged = true
      }
    }
    if (rowChanged) changed += 1
  }
  if (changed) {
    const buffer = await workbook.xlsx.writeBuffer()
    await writeFile(path, new Uint8Array(buffer))
  }
  return changed
}

export function applyQuoteManualMatch(payload: QuoteMatchPayload, matchId: string, item: QuotePriceItem): QuoteMatchPayload {
  const matches = payload.matches.map((match) => {
    if (match.id !== matchId) return match
    return applyItemToMatch(match, item, '人工确认', true)
  })
  return makePayload(payload.quote_name, payload.price_name, payload.price_items, matches, true)
}

export function markQuoteAliasLearned(payload: QuoteMatchPayload, matchId: string, result: { updated: boolean; message: string }): QuoteMatchPayload {
  const matches = payload.matches.map((match) => (
    match.id === matchId ? { ...match, alias_learned: result.updated, alias_message: result.message } : match
  ))
  return makePayload(payload.quote_name, payload.price_name, payload.price_items, matches, true)
}

export async function appendQuotePriceAliases(
  path: string,
  item: QuotePriceItem,
  projectAlias: string,
  parameterAlias: string,
): Promise<{ updated: boolean; message: string; project_aliases: string[]; parameter_aliases: string[] }> {
  const bytes = await readFile(path)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
  const sheet = workbook.getWorksheet(item.sheet)
  if (!sheet) return { updated: false, message: `报价库中未找到分类：${item.sheet}`, project_aliases: [], parameter_aliases: [] }
  const located = findHeaderCells(sheet, PRICE_COLUMNS)
  if (!located) return { updated: false, message: '未找到报价库表头', project_aliases: [], parameter_aliases: [] }
  const rowNumber = resolvePriceItemRow(sheet, located.headers, located.headerRow, item)
  if (!rowNumber) return { updated: false, message: '未能定位报价库对应行', project_aliases: [], parameter_aliases: [] }
  const projectResult = appendAliasToColumn(sheet, located.headers, located.headerRow, rowNumber, '检测项目别名', projectAlias, [item.category, item.material])
  const parameterResult = appendAliasToColumn(sheet, located.headers, located.headerRow, rowNumber, '检测参数别名', parameterAlias, [item.parameter])
  const updated = projectResult.updated || parameterResult.updated
  if (updated) {
    const buffer = await workbook.xlsx.writeBuffer()
    await writeFile(path, new Uint8Array(buffer))
  }
  const message = [projectResult.message, parameterResult.message].filter(Boolean).join('；') || '没有可写入的别名'
  return { updated, message, project_aliases: projectResult.aliases, parameter_aliases: parameterResult.aliases }
}

export function rankQuotePriceItemsForLine(
  priceItems: QuotePriceItem[],
  line: QuoteLine,
  options: { query?: string; tab?: string; mode?: 'fuzzy' | 'exact'; limit?: number | null } = {},
): Array<QuotePriceItem & { score: number; method: string }> {
  const query = cellText(options.query)
  const tab = options.tab || '推荐'
  const mode = options.mode || 'fuzzy'
  const ranked: Array<{ score: number; method: string; item: QuotePriceItem }> = []
  for (const item of priceItems) {
    if (tab !== '推荐' && item.sheet !== tab) continue
    const scored = scorePriceItem(line, item, query, mode)
    if (!scored) continue
    const [score, method] = scored
    if (query && score < 20) continue
    ranked.push({ score, method, item })
  }
  if (tab !== '推荐') ranked.sort((left, right) => left.item.row_number - right.item.row_number || left.item.seq.localeCompare(right.item.seq))
  else ranked.sort((left, right) => right.score - left.score)
  const sliced = options.limit === null ? ranked : ranked.slice(0, options.limit ?? 80)
  return sliced.map(({ item, score, method }) => ({ ...item, score: round(score), method }))
}

export async function exportQuoteMatches(payload: QuoteMatchPayload, outputPath: string): Promise<number> {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = '报价匹配'
  workbook.created = new Date()
  const sheet = workbook.addWorksheet('匹配结果', { views: [{ state: 'frozen', ySplit: 1, showGridLines: false }] })
  const headers = ['序号', '样品名称', '检测项目', '组/点数', '单价', '合价', '备注', '报价编号', '匹配状态']
  sheet.addRow(headers)
  sheet.getRow(1).height = 26
  sheet.getRow(1).eachCell((cell) => {
    cell.font = { name: 'Microsoft YaHei', size: 10, bold: true, color: { argb: 'FF171717' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } }
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFE5E5E5' } } }
  })
  payload.matches.forEach((match, index) => {
    const item = match.matched
    const status = match.manual_confirmed ? '手动确认' : match.match_status
    const row = sheet.addRow([
      match.seq || index + 1,
      match.sample_name,
      match.parameter || match.project_name,
      match.quantity,
      match.matched_price_text || item?.raw_price || '',
      match.calculated_total,
      item?.remark || '',
      match.matched_code || item?.code || '',
      status,
    ])
    row.eachCell((cell) => {
      cell.alignment = { vertical: 'middle' }
      cell.border = { bottom: { style: 'hair', color: { argb: 'FFE5E7EB' } } }
      cell.fill = quoteStatusFill(status)
    })
    row.getCell(headers.length).font = { name: 'Microsoft YaHei', size: 10, bold: true, color: { argb: quoteStatusFont(status) } }
  })
  for (const column of sheet.columns) {
    let max = 10
    column.eachCell?.((cell) => {
      max = Math.max(max, cellText(cell.value).length + 2)
    })
    column.width = Math.min(max, 34)
  }
  const buffer = await workbook.xlsx.writeBuffer()
  await writeFile(outputPath, new Uint8Array(buffer))
  return payload.matches.length
}

export async function createQuoteImportTemplate(outputPath: string): Promise<void> {
  const workbook = new ExcelJS.Workbook()
  const info = workbook.addWorksheet('填写说明', { views: [{ showGridLines: false }] })
  info.addRow(['列名', '是否必填', '填写说明'])
  const columns = [
    ['序号', '选填', '同一样品的多行共用同一序号'],
    ['样品名称', '必填', '材料/构件名称，如：钢筋、预拌砂浆DM'],
    ['检测项目', '必填', '具体检测参数名称，每行填一项'],
    ['组/点数', '选填', '检测数量（纯数字）'],
    ['单价', '选填', '对方报出的单价（纯数字，元）'],
    ['合价', '选填', '对方报出的合价（纯数字，元）'],
    ['备注', '选填', '规格、等级或其他说明'],
  ]
  columns.forEach((row) => info.addRow(row))
  const sheet = workbook.addWorksheet('报价清单', { views: [{ state: 'frozen', ySplit: 1, showGridLines: false }] })
  const headers = ['序号', '样品名称', '检测项目', '组/点数', '单价', '合价', '备注']
  sheet.addRow(headers)
  const examples = [
    [1, '钢筋', '抗拉、弯曲', 11, 138, null, '按三级钢考虑'],
    [1, '', '重量偏差', null, 59, null, ''],
    [2, '预拌砂浆DM', '抗压强度', 2, 275, null, ''],
  ]
  examples.forEach((row) => sheet.addRow(row))
  ;[info, sheet].forEach((ws) => {
    ws.getRow(1).eachCell((cell) => {
      cell.font = { name: 'Microsoft YaHei', bold: true, color: { argb: 'FFFFFFFF' } }
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF171717' } }
      cell.alignment = { horizontal: 'center', vertical: 'middle' }
    })
    ws.columns.forEach((column) => {
      column.width = 18
    })
  })
  const buffer = await workbook.xlsx.writeBuffer()
  await writeFile(outputPath, new Uint8Array(buffer))
}

async function readSheets(path: string): Promise<SheetData[]> {
  const bytes = await readFile(path)
  const workbook = XLSX.read(bytes, { type: 'array', cellDates: true })
  return workbook.SheetNames.map((name) => ({
    name,
    rows: XLSX.utils.sheet_to_json<CellValue[]>(workbook.Sheets[name], { header: 1, raw: true, defval: null }),
  }))
}

async function readQuotePriceBook(path: string): Promise<QuotePriceItem[]> {
  const items: QuotePriceItem[] = []
  for (const sheet of await readSheets(path)) {
    const headerIndex = findHeaderRow(sheet.rows, PRICE_COLUMNS)
    if (headerIndex === null) continue
    const headers = uniqueHeaders(sheet.rows[headerIndex])
    let category = ''
    let material = ''
    for (let index = headerIndex + 1; index < sheet.rows.length; index += 1) {
      const raw = rowObject(headers, sheet.rows[index])
      category = cellText(raw['检测项目']) || category
      material = cellText(raw['检测材料']) || material
      const parameter = cellText(raw['检测参数'])
      const rawPrice = cellText(raw['单价（元）'] ?? raw['单价'])
      if (!parameter || !rawPrice) continue
      const projectAliases = splitAliases(raw['检测项目别名'])
      const parameterAliases = splitAliases(raw['检测参数别名'])
      const code = cellText(raw['报价编号'])
      const rowNumber = index + 1
      items.push({
        id: code || `${sheet.name}-${rowNumber}`,
        sheet: sheet.name,
        row_number: rowNumber,
        seq: cellText(raw['序号']),
        category,
        material,
        parameter,
        unit: cellText(raw['单位']),
        price: toNumber(raw['单价（元）'] ?? raw['单价']),
        raw_price: rawPrice,
        remark: cellText(raw['备注']),
        code,
        project_aliases: projectAliases,
        parameter_aliases: parameterAliases,
        aliases: [...projectAliases, ...parameterAliases],
        search_text: [sheet.name, category, material, parameter, ...projectAliases, ...parameterAliases].join(' '),
        price_rule: null,
      })
    }
  }
  if (!items.length) throw new Error('报价库未识别到有效数据，请确认包含“检测项目/检测材料/检测参数/单价”等表头')
  return items
}

async function readQuoteBook(path: string): Promise<QuoteLine[]> {
  if (path.toLowerCase().endsWith('.xls')) {
    throw new Error('暂不支持 .xls，请先另存为 .xlsx 后再上传')
  }
  const lines: QuoteLine[] = []
  for (const sheet of await readSheets(path)) {
    const headerIndex = findHeaderRow(sheet.rows, QUOTE_HINT_COLUMNS)
    if (headerIndex === null) continue
    const headers = uniqueHeaders(sheet.rows[headerIndex])
    let sample = ''
    let quantity: number | null = null
    for (let index = headerIndex + 1; index < sheet.rows.length; index += 1) {
      const raw = rowObject(headers, sheet.rows[index])
      sample = cellText(raw['检测材料']) || cellText(raw['样品名称']) || cellText(raw['检测项名称']) || sample
      quantity = toNumber(raw['数量'] ?? raw['组/点数']) ?? quantity
      const project = cellText(raw['检测项目'])
      const parameter = cellText(raw['检测参数']) || cellText(raw['具体检测项目']) || project
      if (!sample && !parameter) continue
      const rowNumber = index + 1
      const remark = cellText(raw['备注'])
      lines.push({
        id: `${sheet.name}-${rowNumber}`,
        sheet: sheet.name,
        row_number: rowNumber,
        seq: String(lines.length + 1),
        sample_name: sample,
        project_name: project,
        parameter,
        unit: cellText(raw['单位']),
        quantity,
        source_price: toNumber(raw['单价（元）'] ?? raw['单价']),
        source_total: toNumber(raw['合价（元）'] ?? raw['合价']),
        remark,
        search_text: [sample, project, parameter, remark].join(' '),
      })
    }
  }
  if (!lines.length) throw new Error('报价清单未识别到有效数据，请确认包含“样品名称/检测项目/组/点数”等表头')
  return lines
}

function matchQuoteLine(line: QuoteLine, priceItems: QuotePriceItem[]): QuoteMatchRow {
  let best: { score: number; method: string; item: QuotePriceItem | null } = { score: 0, method: '', item: null }
  const candidates: QuoteCandidate[] = []
  for (const item of priceItems) {
    const [score, method] = scorePriceItem(line, item) || [0, '']
    candidates.push({ score: round(score), method, item })
    if (score > best.score) best = { score, method, item }
  }
  candidates.sort((left, right) => right.score - left.score)
  const status: QuoteStatus = best.score >= 90 ? '自动匹配' : best.score >= 65 ? '待确认' : '未匹配'
  const base: QuoteMatchRow = {
    ...line,
    match_status: status,
    match_score: round(best.score),
    match_method: best.item ? best.method : '',
    matched: null,
    top_candidates: candidates.slice(0, 5),
    matched_price: null,
    matched_price_text: null,
    matched_code: null,
    matched_label: null,
    calculated_total: null,
    price_rule_result: null,
    price_explanation: '',
  }
  return best.item ? applyItemToMatch(base, best.item, best.method, false) : base
}

function applyItemToMatch(match: QuoteMatchRow, item: QuotePriceItem, method: string, manual: boolean): QuoteMatchRow {
  const qty = match.quantity || 1
  return {
    ...match,
    match_status: manual ? '自动匹配' : match.match_status,
    match_method: method,
    matched: item,
    match_score: manual ? (item as QuotePriceItem & { score?: number }).score || match.match_score || 100 : match.match_score,
    matched_price: item.price,
    matched_price_text: item.raw_price,
    matched_code: item.code || item.parameter || item.id,
    matched_label: [item.category, item.material || item.parameter].filter(Boolean).join(' / '),
    matched_remark: item.remark || '',
    price_rule_result: null,
    calculated_total: item.price === null ? null : round(qty * item.price),
    manual_confirmed: manual || match.manual_confirmed,
  }
}

function scorePriceItem(line: QuoteLine, item: QuotePriceItem, query = '', mode: 'fuzzy' | 'exact' = 'fuzzy'): [number, string] | null {
  let [score, method] = exactOrAliasScore(line, item)
  if (!score) [score, method] = scoreLineAgainstItem(line, item)
  if (query) {
    const blob = searchBlob(item)
    const queryCompact = compactText(query)
    if (mode === 'exact' && !compactText(blob).includes(queryCompact)) return null
    const queryScore = compactText(blob).includes(queryCompact) ? 100 : similarity(query, blob)
    score = Math.max(score, queryScore * 0.9)
  }
  return [score, method]
}

function exactOrAliasScore(line: QuoteLine, item: QuotePriceItem): [number, string] {
  const lineParam = compactText(line.parameter || line.project_name)
  const lineAll = compactText(line.search_text)
  const linePair = compactText(makeLineAlias(line.sample_name, line.parameter || line.project_name))
  const itemParam = compactText(item.parameter)
  const lineSample = compactText(line.sample_name)
  const materialScore = Math.max(similarity(line.sample_name, item.material), similarity(line.sample_name, item.category))
  const projectAliasMatch = item.project_aliases.some((alias) => {
    const aliasCompact = compactText(alias)
    return aliasCompact && (lineSample === aliasCompact || lineAll.includes(aliasCompact))
  })
  if (lineParam && lineParam === itemParam) {
    if (materialScore >= 70 || projectAliasMatch) return [Math.min(88 + materialScore * 0.12, 100), '材料+参数匹配']
    return [Math.min(62 + materialScore * 0.28, 88), '参数匹配']
  }
  for (const alias of item.parameter_aliases) {
    const aliasCompact = compactText(alias)
    if (aliasCompact && (linePair === aliasCompact || lineParam === aliasCompact || lineAll.includes(aliasCompact))) {
      return [Math.min((projectAliasMatch || materialScore >= 60 ? 90 : 62) + materialScore * 0.1, 100), '参数别名匹配']
    }
  }
  for (const alias of item.aliases) {
    const aliasCompact = compactText(alias)
    if (aliasCompact && (linePair === aliasCompact || lineParam === aliasCompact || lineAll.includes(aliasCompact))) {
      return [Math.min((projectAliasMatch || materialScore >= 60 ? 86 : 62) + materialScore * 0.1, 100), '别名匹配']
    }
  }
  return [0, '']
}

function scoreLineAgainstItem(line: QuoteLine, item: QuotePriceItem): [number, string] {
  const lineMaterial = line.sample_name
  const lineDetection = line.parameter || line.project_name
  const materialScore = Math.max(similarity(lineMaterial, item.material), similarity(lineMaterial, item.category))
  const projectScore = line.project_name ? similarity(line.project_name, projectBlob(item)) : 0
  const parameterScore = similarity(lineDetection, parameterBlob(item))
  const contextScore = Math.max(materialScore, projectScore)
  const materialExact = Boolean(compactText(lineMaterial) && compactText(projectBlob(item)).includes(compactText(lineMaterial)))
  const parameterExact = Boolean(compactText(lineDetection) && compactText(parameterBlob(item)).includes(compactText(lineDetection)))
  if (parameterExact && (materialExact || contextScore >= 70)) return [98, '材料+参数匹配']
  if (parameterExact) return [Math.min(62 + contextScore * 0.28, 88), '参数匹配']
  if (materialExact) return [Math.min(66 + parameterScore * 0.28 + projectScore * 0.06, 92), '材料匹配']
  return [parameterScore * 0.5 + materialScore * 0.38 + projectScore * 0.12, '三列模糊匹配']
}

function makePayload(
  quotePathOrName: string,
  pricePathOrName: string,
  priceItems: QuotePriceItem[],
  matches: QuoteMatchRow[],
  alreadyNamed = false,
): QuoteMatchPayload {
  const tabs = Array.from(new Set(priceItems.map((item) => item.sheet)))
  return {
    quote_name: alreadyNamed ? quotePathOrName : basename(quotePathOrName),
    price_name: alreadyNamed ? pricePathOrName : basename(pricePathOrName),
    price_items: priceItems,
    tabs,
    matches,
    summary: {
      price_items: priceItems.length,
      quote_lines: matches.length,
      auto: matches.filter((m) => m.match_status === '自动匹配').length,
      review: matches.filter((m) => m.match_status === '待确认').length,
      unmatched: matches.filter((m) => m.match_status === '未匹配').length,
      manual: matches.filter((m) => m.match_method === '人工确认' || m.manual_confirmed).length,
    },
  }
}

function findHeaderRow(rows: CellValue[][], hints: Set<string>): number | null {
  let bestIndex = -1
  let bestScore = 0
  for (let index = 0; index < Math.min(rows.length, 15); index += 1) {
    const values = new Set(rows[index].map((value) => cellText(value)))
    const score = [...values].filter((value) => hints.has(value)).length
    if (score > bestScore) {
      bestScore = score
      bestIndex = index
    }
  }
  return bestScore >= 2 ? bestIndex : null
}

function uniqueHeaders(row: CellValue[]): string[] {
  const counts = new Map<string, number>()
  return row.map((value, index) => {
    const base = cellText(value) || `未命名${index + 1}`
    const count = counts.get(base) || 0
    counts.set(base, count + 1)
    return count ? `${base}_${count + 1}` : base
  })
}

function rowObject(headers: string[], row: CellValue[]): Record<string, CellValue> {
  return Object.fromEntries(headers.map((header, index) => [header, row[index] ?? null]))
}

function findHeaderCells(sheet: ExcelJS.Worksheet, hints: Set<string>): { headerRow: number; headers: Map<string, number> } | null {
  let bestRow = 0
  let bestScore = 0
  let bestHeaders = new Map<string, number>()
  for (let rowIndex = 1; rowIndex <= Math.min(sheet.rowCount, 15); rowIndex += 1) {
    const row = sheet.getRow(rowIndex)
    const headers = new Map<string, number>()
    const values = new Set<string>()
    row.eachCell((cell, colNumber) => {
      const text = cellText(cell.value)
      if (text) {
        headers.set(text, colNumber)
        values.add(text)
      }
    })
    const score = [...values].filter((value) => hints.has(value)).length
    if (score > bestScore) {
      bestScore = score
      bestRow = rowIndex
      bestHeaders = headers
    }
  }
  return bestScore >= 2 ? { headerRow: bestRow, headers: bestHeaders } : null
}

function appendAliasToColumn(
  sheet: ExcelJS.Worksheet,
  headers: Map<string, number>,
  headerRow: number,
  rowNumber: number,
  columnName: string,
  alias: string,
  standardValues: string[],
): { updated: boolean; message: string; aliases: string[] } {
  const cleanAlias = cellText(alias)
  if (!cleanAlias) return { updated: false, message: '', aliases: [] }
  let aliasCol = headers.get(columnName)
  if (!aliasCol) {
    aliasCol = sheet.columnCount + 1
    sheet.getRow(headerRow).getCell(aliasCol).value = columnName
    headers.set(columnName, aliasCol)
  }
  const cell = sheet.getRow(rowNumber).getCell(aliasCol)
  const existingAliases = splitAliases(cell.value)
  const existingNorms = new Set(existingAliases.map((value) => compactText(value)))
  const aliasNorm = compactText(cleanAlias)
  const standardNorms = new Set(standardValues.filter((value) => cellText(value)).map((value) => compactText(value)))
  if (!aliasNorm || existingNorms.has(aliasNorm) || standardNorms.has(aliasNorm)) {
    return { updated: false, message: `${columnName}已存在或无需写入`, aliases: existingAliases }
  }
  const aliases = [...existingAliases, cleanAlias]
  cell.value = aliases.join('/')
  return { updated: true, message: `${columnName}已写入`, aliases }
}

function resolvePriceItemRow(sheet: ExcelJS.Worksheet, headers: Map<string, number>, headerRow: number, item: QuotePriceItem): number | null {
  if (rowNumberMatchesPriceItem(sheet, headers, item.row_number, item)) return item.row_number
  let lastCategory = ''
  let lastMaterial = ''
  for (let rowIndex = headerRow + 1; rowIndex <= sheet.rowCount; rowIndex += 1) {
    const categoryCol = headers.get('检测项目')
    const materialCol = headers.get('检测材料')
    const category = categoryCol ? cellText(sheet.getRow(rowIndex).getCell(categoryCol).value) : ''
    const material = materialCol ? cellText(sheet.getRow(rowIndex).getCell(materialCol).value) : ''
    if (category) lastCategory = category
    if (material) lastMaterial = material
    if (rowNumberMatchesPriceItem(sheet, headers, rowIndex, item, lastCategory, lastMaterial)) return rowIndex
  }
  return null
}

function rowNumberMatchesPriceItem(
  sheet: ExcelJS.Worksheet,
  headers: Map<string, number>,
  rowNumber: number,
  item: QuotePriceItem,
  category?: string,
  material?: string,
): boolean {
  if (rowNumber < 1 || rowNumber > sheet.rowCount) return false
  const row = sheet.getRow(rowNumber)
  const parameterCol = headers.get('检测参数')
  if (!parameterCol) return false
  if (compactText(row.getCell(parameterCol).value) !== compactText(item.parameter)) return false
  const seqCol = headers.get('序号')
  if (seqCol && cellText(item.seq)) {
    const seq = cellText(row.getCell(seqCol).value)
    if (seq && compactText(seq) !== compactText(item.seq)) return false
  }
  const priceCol = headers.get('单价（元）') || headers.get('单价')
  if (priceCol && cellText(item.raw_price)) {
    const rawPrice = cellText(row.getCell(priceCol).value)
    if (rawPrice && compactText(rawPrice) !== compactText(item.raw_price)) return false
  }
  if (category !== undefined && cellText(item.category) && compactText(category) !== compactText(item.category)) return false
  if (material !== undefined && cellText(item.material) && compactText(material) !== compactText(item.material)) return false
  return true
}

function normalizeText(value: unknown, dropStopWords = false): string {
  let text = cellText(value).normalize('NFKC').trim().toLowerCase()
  text = text.replaceAll('（', '(').replaceAll('）', ')')
  text = text.replace(/\s+/g, '')
  text = text.replace(/[，,、;；/\\|+＋]/g, ' ')
  text = text.replace(/[()（）【】[\]{}:：\-—_]/g, '')
  text = text.replace(/\s+/g, ' ').trim()
  if (dropStopWords) {
    for (const word of STOP_WORDS) text = text.replaceAll(word, '')
  }
  return text
}

function compactText(value: unknown): string {
  return normalizeText(value, true).replaceAll(' ', '')
}

function cellText(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'number' && Number.isNaN(value)) return ''
  if (value instanceof Date) return value.toISOString().slice(0, 10)
  const text = String(value).trim()
  return text.endsWith('.0') && /^\d+\.0$/.test(text) ? text.slice(0, -2) : text
}

function toNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  const text = cellText(value).replaceAll(',', '')
  return /^-?\d+(?:\.\d+)?$/.test(text) ? Number(text) : null
}

function splitAliases(value: unknown): string[] {
  const text = cellText(value)
  return text ? text.split(/[,，、;；/|｜\n]+/).map((item) => item.trim()).filter(Boolean) : []
}

function similarity(left: string, right: string): number {
  const leftText = compactText(left)
  const rightText = compactText(right)
  if (!leftText || !rightText) return 0
  let ratio = sequenceMatcherRatio(leftText, rightText)
  if (leftText.includes(rightText) || rightText.includes(leftText)) {
    ratio = Math.max(ratio, Math.min(leftText.length, rightText.length) / Math.max(leftText.length, rightText.length) + 0.18)
  }
  return Math.min(ratio * 100, 100)
}

function sequenceMatcherRatio(left: string, right: string): number {
  const matches = matchingBlockSize(left, 0, left.length, right, 0, right.length)
  return left.length + right.length ? (2 * matches) / (left.length + right.length) : 1
}

function matchingBlockSize(left: string, leftStart: number, leftEnd: number, right: string, rightStart: number, rightEnd: number): number {
  let bestLeft = leftStart
  let bestRight = rightStart
  let bestSize = 0
  const lengths = new Map<number, number>()
  for (let i = leftStart; i < leftEnd; i += 1) {
    const nextLengths = new Map<number, number>()
    for (let j = rightStart; j < rightEnd; j += 1) {
      if (left[i] !== right[j]) continue
      const size = (lengths.get(j - 1) || 0) + 1
      nextLengths.set(j, size)
      if (size > bestSize) {
        bestLeft = i - size + 1
        bestRight = j - size + 1
        bestSize = size
      }
    }
    lengths.clear()
    nextLengths.forEach((value, key) => lengths.set(key, value))
  }
  if (!bestSize) return 0
  return (
    matchingBlockSize(left, leftStart, bestLeft, right, rightStart, bestRight)
    + bestSize
    + matchingBlockSize(left, bestLeft + bestSize, leftEnd, right, bestRight + bestSize, rightEnd)
  )
}

function searchBlob(item: QuotePriceItem): string {
  return [item.code, item.category, item.material, item.parameter, item.remark, item.aliases.join(' ')].join(' ')
}

function projectBlob(item: QuotePriceItem): string {
  return [item.category, item.material, ...item.project_aliases].join(' ')
}

function parameterBlob(item: QuotePriceItem): string {
  return [item.parameter, ...item.parameter_aliases].join(' ')
}

function makeLineAlias(sampleName: string, parameter: string): string {
  return sampleName && parameter ? `${sampleName}:${parameter}` : parameter || sampleName
}

function basename(path: string): string {
  return path.split(/[\\/]/).pop() || path
}

function round(value: number): number {
  return Math.round(value * 100) / 100
}

function quoteStatusFill(status: QuoteStatus | '手动确认'): ExcelJS.Fill {
  const colors: Record<QuoteStatus | '手动确认', string> = {
    自动匹配: 'FFDCFCE7',
    手动确认: 'FFDCFCE7',
    待确认: 'FFFEF3C7',
    未匹配: 'FFFEE2E2',
  }
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: colors[status] } }
}

function quoteStatusFont(status: QuoteStatus | '手动确认'): string {
  const colors: Record<QuoteStatus | '手动确认', string> = {
    自动匹配: 'FF166534',
    手动确认: 'FF166534',
    待确认: 'FF92400E',
    未匹配: 'FF991B1B',
  }
  return colors[status]
}
