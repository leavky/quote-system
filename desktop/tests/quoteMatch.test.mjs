import assert from 'node:assert/strict'
import { copyFile, mkdtemp, readFile as nodeReadFile, rm, writeFile as nodeWriteFile } from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { mock, test } from 'node:test'

import ExcelJS from 'exceljs'

mock.module('@tauri-apps/plugin-fs', {
  namedExports: {
    readFile: async (filePath) => new Uint8Array(await nodeReadFile(filePath)),
    writeFile: async (filePath, data) => nodeWriteFile(filePath, data),
  },
})

const {
  appendQuotePriceAliases,
  backfillQuoteWorkbook,
  buildQuoteMatchPayload,
  loadQuotePriceBook,
  markQuoteAliasLearned,
} = await import('../src/quoteMatch.ts')

const dataDirectory = path.resolve(import.meta.dirname, '../../data')
const sourcePriceBook = path.join(dataDirectory, '检测报价表.xlsx')

for (const quoteName of ['报价清单1.xlsx', '报价清单3.xlsx']) {
  test(`matches and backfills ${quoteName}`, async (context) => {
    const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), 'quote-match-'))
    context.after(() => rm(temporaryDirectory, { recursive: true, force: true }))
    const quotePath = path.join(dataDirectory, quoteName)
    const outputPath = path.join(temporaryDirectory, quoteName)
    const payload = await buildQuoteMatchPayload(quotePath, sourcePriceBook)

    assert.ok(payload.matches.length > 0)
    assert.equal(payload.matches.length, payload.summary.quote_lines)
    assert.ok(payload.matches.every((match) => match.matched))
    assert.equal(await backfillQuoteWorkbook(quotePath, payload, outputPath), payload.matches.length)

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(await nodeReadFile(outputPath))
    const firstMatch = payload.matches[0]
    const sheet = workbook.getWorksheet(firstMatch.sheet)
    assert.ok(sheet)
    const headerRow = Array.from({ length: Math.min(sheet.rowCount, 15) }, (_, index) => index + 1)
      .find((rowNumber) => sheet.getRow(rowNumber).values.some((value) => ['检测项目', '具体检测项目'].includes(String(value ?? '').trim())))
    assert.ok(headerRow)
    const headers = new Map()
    sheet.getRow(headerRow).eachCell((cell, column) => headers.set(String(cell.value ?? '').trim(), column))
    const unitPriceColumn = headers.get('单价（元）') || headers.get('单价')
    const totalColumn = headers.get('合价（元）') || headers.get('合价')
    assert.equal(sheet.getCell(firstMatch.row_number, unitPriceColumn).value, firstMatch.matched_price ?? firstMatch.matched_price_text)
    assert.equal(sheet.getCell(firstMatch.row_number, totalColumn).value, firstMatch.calculated_total)
    assert.equal(sheet.getCell(firstMatch.row_number, headers.get('报价编号')).value, firstMatch.matched_code)
    assert.equal(sheet.getCell(firstMatch.row_number, headers.get('匹配状态')).value, firstMatch.match_status)
    assert.equal(sheet.getCell(firstMatch.row_number, headers.get('匹配方法')).value, firstMatch.match_method)
  })
}

test('learns aliases in the price book copy and refreshes the current payload', async (context) => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), 'quote-alias-'))
  context.after(() => rm(temporaryDirectory, { recursive: true, force: true }))
  const priceBookPath = path.join(temporaryDirectory, '检测报价表.xlsx')
  await copyFile(sourcePriceBook, priceBookPath)
  const quotePath = path.join(dataDirectory, '报价清单1.xlsx')
  const payload = await buildQuoteMatchPayload(quotePath, priceBookPath)
  const match = payload.matches[0]
  assert.ok(match.matched)
  const projectAlias = `回归项目别名-${Date.now()}`
  const parameterAlias = `回归参数别名-${Date.now()}`
  const result = await appendQuotePriceAliases(priceBookPath, match.matched, projectAlias, parameterAlias)

  assert.equal(result.updated, true)
  assert.ok(result.project_aliases.includes(projectAlias))
  assert.ok(result.parameter_aliases.includes(parameterAlias))
  const refreshedPayload = markQuoteAliasLearned(payload, match.id, match.matched, result)
  const refreshedMatch = refreshedPayload.matches.find((item) => item.id === match.id)
  assert.equal(refreshedMatch.alias_learned, true)
  assert.ok(refreshedMatch.matched.project_aliases.includes(projectAlias))
  assert.ok(refreshedPayload.price_items.some((item) => item.project_aliases.includes(projectAlias)))

  const reloadedItems = await loadQuotePriceBook(priceBookPath)
  const reloaded = reloadedItems.find((item) => item.sheet === match.matched.sheet && item.row_number === match.matched.row_number)
  assert.ok(reloaded.project_aliases.includes(projectAlias))
  assert.ok(reloaded.parameter_aliases.includes(parameterAlias))
})
