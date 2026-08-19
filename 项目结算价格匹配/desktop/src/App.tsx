import { useEffect, useMemo, useState } from 'react'
import {
  Button,
  Card,
  Chip,
  DateRangePicker,
  DateField,
  Input,
  RangeCalendar,
} from '@heroui/react'
import type { CalendarDate } from '@internationalized/date'
import { invoke } from '@tauri-apps/api/core'
import { open, save } from '@tauri-apps/plugin-dialog'
import { readFile, writeFile } from '@tauri-apps/plugin-fs'
import type { RangeValue } from '@react-types/shared'
import {
  applyManualSettlementEdit,
  buildMatchPayload,
  exportResult,
  loadSettlementPriceBook,
  repricePayload,
  updateSettlementPriceBook,
  type MatchPayload,
  type ManualSettlementEdit,
  type PriceItem,
  type PriceSystem,
  type PriorityEntry,
  type RecordRow,
  type Summary,
} from './settlement'
import {
  applyQuoteManualMatch,
  appendQuotePriceAliases,
  buildQuoteMatchPayload,
  createQuoteImportTemplate,
  exportQuoteMatches,
  loadQuotePriceBook,
  updateQuotePriceBook,
  markQuoteAliasLearned,
  rankQuotePriceItemsForLine,
  type QuoteMatchPayload,
  type QuoteMatchRow,
  type QuotePriceItem,
  type QuoteStatus,
} from './quoteMatch'
import './App.css'

type DefaultPaths = { ledger_path: string; price_path: string; quote_price_path: string; quote_template_path: string }
type BusyAction = '' | 'initial' | 'match' | 'priority' | 'export' | 'quoteMatch' | 'quoteExport' | 'quoteTemplate'
type ModuleKey = 'projectPrice' | 'quoteMatch' | 'priceBook'

const statusClass: Record<string, string> = {
  已匹配: 'success',
  未匹配: 'danger',
  字段缺失: 'warning',
  价格表重复: 'warning',
  无可用价格: 'secondary',
}

const fmtMoney = (value: number | null | undefined) =>
  value === null || value === undefined ? '' : value.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })

const basename = (path: string) => path.split(/[\\/]/).pop() || path

type SettlementEditDraft = {
  unit_price: string
  quantity: string
  discount: string
  subtotal: string
  total: string
  unit_billing: string
}

function App() {
  const [ledgerPath, setLedgerPath] = useState('')
  const [pricePath, setPricePath] = useState('')
  const [quotePath, setQuotePath] = useState('')
  const [quotePricePath, setQuotePricePath] = useState('')
  const [quoteTemplatePath, setQuoteTemplatePath] = useState('')
  const [data, setData] = useState<MatchPayload | null>(null)
  const [quoteData, setQuoteData] = useState<QuoteMatchPayload | null>(null)
  const [priorityTopName, setPriorityTopName] = useState('')
  const [error, setError] = useState('')
  const [notice, setNotice] = useState('')
  const [busyAction, setBusyAction] = useState<BusyAction>('initial')
  const [activeModule, setActiveModule] = useState<ModuleKey>('projectPrice')
  const [query, setQuery] = useState('')
  const [statusFilter, setStatusFilter] = useState('')
  const [categoryFilter, setCategoryFilter] = useState('')
  const [dateRange, setDateRange] = useState<RangeValue<CalendarDate> | null>(null)
  const [projectFilter, setProjectFilter] = useState('')
  const [itemFilter, setItemFilter] = useState('')
  const [sortOrder, setSortOrder] = useState('source')
  const [pageSize, setPageSize] = useState(100)
  const [page, setPage] = useState(1)
  const [quoteStatusFilter, setQuoteStatusFilter] = useState<'all' | 'auto' | 'review' | 'unmatched' | 'manual'>('all')
  const [quotePickId, setQuotePickId] = useState('')
  const [quotePickQuery, setQuotePickQuery] = useState('')
  const [quotePickTab, setQuotePickTab] = useState('推荐')
  const [quotePickMode, setQuotePickMode] = useState<'fuzzy' | 'exact'>('fuzzy')
  const [priceBookTab, setPriceBookTab] = useState<'settlement' | 'quote'>('settlement')
  const [settlementBook, setSettlementBook] = useState<{ items: PriceItem[]; systems: PriceSystem[] } | null>(null)
  const [quoteBook, setQuoteBook] = useState<QuotePriceItem[]>([])
  const [priceBookQueries, setPriceBookQueries] = useState({ settlement: '', quote: '' })
  const [priceBookBusy, setPriceBookBusy] = useState(false)
  const [settlementEditId, setSettlementEditId] = useState('')
  const [settlementEditDraft, setSettlementEditDraft] = useState<SettlementEditDraft | null>(null)
  const loading = busyAction !== ''

  async function pickFile(kind: 'ledger' | 'price' | 'quote' | 'quotePrice') {
    const selected = await open({
      multiple: false,
      filters: [{ name: 'Excel', extensions: kind === 'quote' || kind === 'quotePrice' ? ['xlsx'] : ['xls', 'xlsx'] }],
    })
    if (typeof selected !== 'string') return
    if (kind === 'ledger') setLedgerPath(selected)
    else if (kind === 'price') {
      setPricePath(selected)
      setSettlementBook(null)
      if (activeModule === 'priceBook') {
        setPriceBookBusy(true)
        try { setSettlementBook(await loadSettlementPriceBook(selected)) } catch (err) { setError(readError(err)) } finally { setPriceBookBusy(false) }
      }
    } else if (kind === 'quotePrice') {
      setQuotePricePath(selected)
      setQuoteBook([])
      if (activeModule === 'priceBook') {
        setPriceBookBusy(true)
        try { setQuoteBook(await loadQuotePriceBook(selected)) } catch (err) { setError(readError(err)) } finally { setPriceBookBusy(false) }
      }
    } else setQuotePath(selected)
  }

  async function loadMatch(nextPriority?: PriorityEntry[], action: BusyAction = 'match') {
    if (!ledgerPath || !pricePath) {
      setError('请先选择台账明细表和计费价格汇总表')
      return
    }
    setBusyAction(action)
    setError('')
    try {
      const basePayload = nextPriority && data ? repricePayload(data, nextPriority) : await buildMatchPayload(ledgerPath, pricePath)
      const activePriority = nextPriority || buildPriorityOrder(basePayload.result.systems, priorityTopName)
      const payload = nextPriority && data ? basePayload : activePriority.length ? repricePayload(basePayload, activePriority) : basePayload
      setData(payload)
      if (payload.priority.length) setPriorityTopName(payload.priority[0].name)
    } catch (err) {
      setError(readError(err))
    } finally {
      setBusyAction('')
    }
  }

  async function exportRows() {
    if (!data) return
    const target = await save({
      defaultPath: '项目结算价格匹配结果.xlsx',
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    })
    if (!target) return
    setBusyAction('export')
    setError('')
    try {
      await exportResult(data, target, filtered.map((row) => row.id))
    } catch (err) {
      setError(readError(err))
    } finally {
      setBusyAction('')
    }
  }

  async function runQuoteMatch() {
    if (!quotePath) {
      setError('请先选择外部报价清单')
      return
    }
    if (!quotePricePath) {
      setError('请先选择报价库')
      return
    }
    setBusyAction('quoteMatch')
    setError('')
    try {
      const payload = await buildQuoteMatchPayload(quotePath, quotePricePath)
      setQuoteData(payload)
      setQuoteStatusFilter('all')
    } catch (err) {
      setError(readError(err))
    } finally {
      setBusyAction('')
    }
  }

  async function exportQuoteRows() {
    if (!quoteData) return
    const target = await save({
      defaultPath: '报价匹配结果.xlsx',
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    })
    if (!target) return
    setBusyAction('quoteExport')
    setError('')
    try {
      await exportQuoteMatches(quoteData, target)
    } catch (err) {
      setError(readError(err))
    } finally {
      setBusyAction('')
    }
  }

  async function saveQuoteTemplate() {
    const target = await save({
      defaultPath: '报价清单导入模板.xlsx',
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    })
    if (!target) return
    setBusyAction('quoteTemplate')
    setError('')
    try {
      if (quoteTemplatePath) {
        const bytes = await readFile(quoteTemplatePath)
        await writeFile(target, bytes)
      } else {
        await createQuoteImportTemplate(target)
      }
    } catch (err) {
      setError(readError(err))
    } finally {
      setBusyAction('')
    }
  }

  async function openPriceBook() {
    setActiveModule('priceBook')
    const loaders: Promise<void>[] = []
    if (!settlementBook && pricePath) loaders.push(loadSettlementPriceBook(pricePath).then(setSettlementBook))
    if (!quoteBook.length && quotePricePath) loaders.push(loadQuotePriceBook(quotePricePath).then(setQuoteBook))
    if (!loaders.length) return
    setPriceBookBusy(true)
    try {
      await Promise.all(loaders)
    } catch (err) {
      setError(readError(err))
    } finally {
      setPriceBookBusy(false)
    }
  }

  async function savePriceBook() {
    setPriceBookBusy(true)
    setError('')
    setNotice('')
    try {
      if (priceBookTab === 'settlement' && settlementBook) {
        const changed = await updateSettlementPriceBook(pricePath, settlementBook.items.map((item) => ({ sheet: item.sheet, row_number: item.row_number, prices: item.prices })))
        setNotice(changed ? `已保存 ${changed} 行结算价格，返回项目价格匹配后重新匹配即可生效` : '没有检测到价格变更')
      } else {
        const changed = await updateQuotePriceBook(quotePricePath, quoteBook.map((item) => ({ sheet: item.sheet, row_number: item.row_number, price: item.price, remark: item.remark })))
        setNotice(changed ? `已保存 ${changed} 行报价库数据，下一次报价匹配将使用新价格` : '没有检测到价格变更')
      }
    } catch (err) {
      setError(readError(err))
    } finally {
      setPriceBookBusy(false)
    }
  }

  function openSettlementEditor(row: RecordRow) {
    const manualEdit = data?.result.manual_edits.find((item) => item.row_id === row.id)
    const unitPrice = manualEdit?.unit_price ?? row.selected_price ?? row.source_unit_price
    const quantity = manualEdit?.quantity ?? row.quantity
    const discount = manualEdit?.discount ?? parseNumeric(row.original['折扣(元)']) ?? 0
    const subtotal = manualEdit?.subtotal ?? (unitPrice !== null && unitPrice !== undefined && quantity !== null && quantity !== undefined ? roundCurrency(unitPrice * quantity) : null)
    const total = manualEdit?.total ?? (subtotal !== null ? roundCurrency(subtotal - discount) : null)
    setSettlementEditId(row.id)
    setSettlementEditDraft({
      unit_price: valueDraft(unitPrice),
      quantity: valueDraft(quantity),
      discount: valueDraft(discount),
      subtotal: valueDraft(subtotal),
      total: valueDraft(total),
      unit_billing: valueDraft(manualEdit?.unit_billing ?? total),
    })
    setError('')
  }

  function closeSettlementEditor() {
    setSettlementEditId('')
    setSettlementEditDraft(null)
  }

  function saveSettlementEdit() {
    if (!data || !settlementEditId || !settlementEditDraft) return
    const unitPrice = parseNumeric(settlementEditDraft.unit_price)
    const quantity = parseNumeric(settlementEditDraft.quantity)
    if (unitPrice === null || quantity === null) {
      setError('请填写有效的单价和数量')
      return
    }
    const subtotal = parseNumeric(settlementEditDraft.subtotal) ?? roundCurrency(unitPrice * quantity)
    const discount = parseNumeric(settlementEditDraft.discount) ?? 0
    const total = parseNumeric(settlementEditDraft.total) ?? roundCurrency(subtotal - discount)
    const edit: ManualSettlementEdit = {
      row_id: settlementEditId,
      unit_price: unitPrice,
      quantity,
      discount,
      subtotal,
      total,
      unit_billing: parseNumeric(settlementEditDraft.unit_billing) ?? total,
    }
    setData(applyManualSettlementEdit(data, edit))
    closeSettlementEditor()
  }

  function openQuotePicker(row: QuoteMatchRow) {
    setQuotePickId(row.id)
    setQuotePickQuery(row.parameter || row.project_name || row.sample_name)
    setQuotePickTab('推荐')
    setQuotePickMode('fuzzy')
    setError('')
  }

  function closeQuotePicker() {
    setQuotePickId('')
    setQuotePickQuery('')
    setQuotePickTab('推荐')
  }

  async function chooseQuoteMatch(item: QuotePriceItem) {
    if (!quoteData || !quotePickId) return
    const activeMatchId = quotePickId
    const row = quotePickRow
    let payload = applyQuoteManualMatch(quoteData, activeMatchId, item)
    if (row && quotePricePath) {
      try {
        const learnResult = await appendQuotePriceAliases(quotePricePath, item, row.sample_name, row.parameter || row.project_name)
        payload = markQuoteAliasLearned(payload, activeMatchId, learnResult)
      } catch (err) {
        payload = markQuoteAliasLearned(payload, activeMatchId, { updated: false, message: readError(err) })
      }
    }
    setQuoteData(payload)
    closeQuotePicker()
  }

  useEffect(() => {
    async function loadDefaults() {
      setBusyAction('initial')
      try {
        const defaults = await invoke<DefaultPaths>('default_paths')
        setLedgerPath(defaults.ledger_path)
        setPricePath(defaults.price_path)
        setQuotePricePath(defaults.quote_price_path)
        setQuoteTemplatePath(defaults.quote_template_path)
        if (defaults.ledger_path && defaults.price_path) {
          const payload = await buildMatchPayload(defaults.ledger_path, defaults.price_path)
          setData(payload)
          if (payload.priority.length) setPriorityTopName(payload.priority[0].name)
        }
      } catch (err) {
        setError(readError(err))
      } finally {
        setBusyAction('')
      }
    }
    loadDefaults()
  }, [])

  const categories = useMemo(
    () => Array.from(new Set((data?.records || []).map((row) => row.report_category).filter(Boolean))).sort(),
    [data],
  )
  const statuses = useMemo(
    () => Array.from(new Set((data?.records || []).map((row) => row.status).filter(Boolean))),
    [data],
  )
  const prioritySystems = useMemo(() => data?.systems || [], [data])
  const filtered = useMemo(() => {
    const startDate = dateRange?.start.toString() || ''
    const endDate = dateRange?.end.toString() || ''
    const keyword = query.trim().toLowerCase()
    const project = projectFilter.trim().toLowerCase()
    const item = itemFilter.trim().toLowerCase()
    const rows = (data?.records || []).filter((row) => {
      if (startDate && (!row.date || row.date < startDate)) return false
      if (endDate && (!row.date || row.date > endDate)) return false
      if (statusFilter && row.status !== statusFilter) return false
      if (categoryFilter && row.report_category !== categoryFilter) return false
      if (project && !row.project_name.toLowerCase().includes(project)) return false
      if (item && !row.billing_item.toLowerCase().includes(item)) return false
      if (!keyword) return true
      return [row.report_number, row.project_name, row.unit_project, row.client_name, row.report_category, row.billing_item, row.matched_code, row.status]
        .join(' ')
        .toLowerCase()
        .includes(keyword)
    })
    if (sortOrder === 'date_desc') rows.sort((left, right) => right.date.localeCompare(left.date))
    if (sortOrder === 'date_asc') rows.sort((left, right) => left.date.localeCompare(right.date))
    if (sortOrder === 'amount_desc') rows.sort((left, right) => (right.settlement_amount ?? -Infinity) - (left.settlement_amount ?? -Infinity))
    return rows
  }, [data, categoryFilter, dateRange, itemFilter, projectFilter, query, sortOrder, statusFilter])
  const pageCount = Math.max(1, Math.ceil(filtered.length / pageSize))
  const currentPage = Math.min(page, pageCount)
  const pagedRows = useMemo(
    () => filtered.slice((currentPage - 1) * pageSize, currentPage * pageSize),
    [currentPage, filtered, pageSize],
  )
  useEffect(() => {
    setPage(1)
  }, [categoryFilter, dateRange, itemFilter, projectFilter, query, sortOrder, statusFilter, pageSize])
  const visibleQuoteRows = useMemo(() => {
    const rows = quoteData?.matches || []
    if (quoteStatusFilter === 'auto') return rows.filter((row) => row.match_status === '自动匹配')
    if (quoteStatusFilter === 'review') return rows.filter((row) => row.match_status === '待确认')
    if (quoteStatusFilter === 'unmatched') return rows.filter((row) => row.match_status === '未匹配')
    if (quoteStatusFilter === 'manual') return rows.filter((row) => row.manual_confirmed)
    return rows
  }, [quoteData, quoteStatusFilter])
  const quotePickRow = useMemo(() => quoteData?.matches.find((row) => row.id === quotePickId) || null, [quoteData, quotePickId])
  const settlementEditRow = useMemo(() => data?.records.find((row) => row.id === settlementEditId) || null, [data, settlementEditId])
  const quoteCandidates = useMemo(() => {
    if (!quoteData || !quotePickRow) return []
    return rankQuotePriceItemsForLine(quoteData.price_items, quotePickRow, {
      query: quotePickQuery,
      tab: quotePickTab,
      mode: quotePickMode,
      limit: null,
    })
  }, [quoteData, quotePickMode, quotePickQuery, quotePickRow, quotePickTab])

  const summary = summarize(filtered)

  function resetFilters() {
    setDateRange(null)
    setProjectFilter('')
    setCategoryFilter('')
    setItemFilter('')
    setStatusFilter('')
    setQuery('')
    setSortOrder('source')
    setPageSize(100)
    setPage(1)
  }

  return (
    <main className="shell">
      <div className="app-body">
        <aside className="sidebar">
          <div className="sidebar-brand">
            <strong>报价系统</strong>
          </div>
          <nav className="sidebar-nav" aria-label="功能模块">
          <button
            aria-label="项目价格匹配"
            className={activeModule === 'projectPrice' ? 'active' : ''}
            title="项目价格匹配"
            type="button"
            onClick={() => setActiveModule('projectPrice')}
          >
            <span aria-hidden="true" className="sidebar-menu-icon">¥</span>
          </button>
          <button
            aria-label="价格维护"
            className={activeModule === 'priceBook' ? 'active' : ''}
            title="价格维护"
            type="button"
            onClick={openPriceBook}
          >
            <span aria-hidden="true" className="sidebar-menu-icon">▤</span>
          </button>
          <button
            aria-label="报价匹配"
            className={activeModule === 'quoteMatch' ? 'active' : ''}
            title="报价匹配"
            type="button"
            onClick={() => setActiveModule('quoteMatch')}
          >
            <span aria-hidden="true" className="sidebar-menu-icon">≋</span>
          </button>
          </nav>
        </aside>

        <section className="module-content">
      {error && <div className="message error">{error}</div>}
      {notice && <div className="message success">{notice}</div>}
      {activeModule === 'projectPrice' ? (
        <>
      <Card className="panel summary-panel">
        <h2 className="summary-title">项目价格匹配</h2>
        <section className="summary-left">
          <div className="summary-files">
            <FilePick
              className="summary-control summary-file"
              label="台账明细表"
              value={ledgerPath ? basename(ledgerPath) : data?.ledger_name || '使用默认台账'}
              onClick={() => pickFile('ledger')}
            />
            <FilePick
              className="summary-control summary-file"
              label="计费价格汇总表"
              value={pricePath ? basename(pricePath) : data?.price_name || '使用默认价格表'}
              onClick={() => pickFile('price')}
            />
          </div>
        </section>

        <section className="summary-right">
          <div className="summary-actions-row">
            <div className="priority-select-box summary-control summary-priority">
              <select
                id="priorityTop"
                aria-label="价格体系优先级"
                className="native-select"
                value={priorityTopName}
                onChange={async (event) => {
                  const name = event.currentTarget.value
                  setPriorityTopName(name)
                  if (!data?.systems.length || !name) return
                  await loadMatch(buildPriorityOrder(data.systems, name), 'priority')
                }}
              >
                <option value="">请选择优先级</option>
                {prioritySystems.map((system) => (
                  <option key={system.name} value={system.name}>
                    {system.name}（{system.non_empty} 条有价）
                  </option>
                ))}
              </select>
            </div>
            <div className="module-actions">
              <Button type="button" size="sm" variant="secondary" onPress={() => loadMatch()} isDisabled={loading}>
                {busyAction === 'match' ? '匹配中' : '重新匹配'}
              </Button>
              <Button type="button" size="sm" variant="primary" onPress={exportRows} isDisabled={!data || loading}>
                {busyAction === 'export' ? '导出中' : '导出当前结果'}
              </Button>
            </div>
          </div>
          <section className="metrics">
            <Metric label="当前记录" value={summary.total.toString()} />
            <Metric label="已匹配" value={summary.matched.toString()} />
            <Metric label="待处理" value={summary.unresolved.toString()} />
            <Metric label="结算金额" value={fmtMoney(summary.total_amount)} money />
          </section>
        </section>
      </Card>

      <Card className="panel result-panel">
        <div className="filters">
          <div className="filters-row">
            <DateRangePicker
              aria-label="委托日期范围"
              className="date-range-filter"
              value={dateRange}
              onChange={setDateRange}
            >
              <DateField.Group fullWidth>
                <DateField.InputContainer>
                  <DateField.Input slot="start">
                    {(segment) => <DateField.Segment segment={segment} />}
                  </DateField.Input>
                  <DateRangePicker.RangeSeparator>至</DateRangePicker.RangeSeparator>
                  <DateField.Input slot="end">
                    {(segment) => <DateField.Segment segment={segment} />}
                  </DateField.Input>
                </DateField.InputContainer>
                <DateField.Suffix>
                  <DateRangePicker.Trigger aria-label="打开日期选择">
                    <DateRangePicker.TriggerIndicator />
                  </DateRangePicker.Trigger>
                </DateField.Suffix>
              </DateField.Group>
              <DateRangePicker.Popover>
                <RangeCalendar firstDayOfWeek="mon" visibleDuration={{ months: 1 }} weeksInMonth={6}>
                  <RangeCalendar.Header>
                    <RangeCalendar.YearPickerTrigger>
                      <RangeCalendar.YearPickerTriggerHeading />
                      <RangeCalendar.YearPickerTriggerIndicator />
                    </RangeCalendar.YearPickerTrigger>
                    <RangeCalendar.NavButton slot="previous" />
                    <RangeCalendar.NavButton slot="next" />
                  </RangeCalendar.Header>
                  <RangeCalendar.Grid>
                  <RangeCalendar.GridHeader>
                    {(day) => <RangeCalendar.HeaderCell>{day}</RangeCalendar.HeaderCell>}
                  </RangeCalendar.GridHeader>
                  <RangeCalendar.GridBody>
                    {(date) => <RangeCalendar.Cell date={date} />}
                  </RangeCalendar.GridBody>
                  </RangeCalendar.Grid>
                </RangeCalendar>
              </DateRangePicker.Popover>
            </DateRangePicker>
            <Input className="filter-project" value={projectFilter} onChange={(event) => setProjectFilter(event.currentTarget.value)} placeholder="工程名称" />
            <select className="native-select filter-category" value={categoryFilter} onChange={(event) => setCategoryFilter(event.currentTarget.value)}>
              <option value="">全部报告类别</option>
              {categories.map((item) => <option key={item} value={item}>{item}</option>)}
            </select>
            <Input className="filter-item" value={itemFilter} onChange={(event) => setItemFilter(event.currentTarget.value)} placeholder="计费项目" />
          </div>
          <div className="filters-row">
            <select className="native-select filter-status" value={statusFilter} onChange={(event) => setStatusFilter(event.currentTarget.value)}>
              <option value="">全部状态</option>
              {statuses.map((item) => <option key={item} value={item}>{item}</option>)}
            </select>
            <select className="native-select filter-sort" value={sortOrder} onChange={(event) => setSortOrder(event.currentTarget.value)}>
              <option value="source">台账原顺序</option>
              <option value="date_desc">日期从新到旧</option>
              <option value="date_asc">日期从旧到新</option>
              <option value="amount_desc">结算金额从高到低</option>
            </select>
            <Input className="filter-query" value={query} onChange={(event) => setQuery(event.currentTarget.value)} placeholder="关键词：报告编号、委托单位、状态等" />
            <Button className="filter-reset" type="button" size="sm" variant="secondary" onPress={resetFilters}>重置筛选</Button>
          </div>
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>委托日期</th>
                <th>报告编号</th>
                <th>工程名称</th>
                <th>报告类别</th>
                <th>计费项目</th>
                <th className="num">数量</th>
                <th className="num">台账单价</th>
                <th>价格体系</th>
                <th className="num">结算单价</th>
                <th className="num">结算金额</th>
                <th>项目编号</th>
                <th>状态</th>
              </tr>
            </thead>
            <tbody>
              {pagedRows.map((row) => (
                <tr className="editable-result-row" key={row.id} onClick={() => openSettlementEditor(row)}>
                  <td>{row.date}</td>
                  <td>{row.report_number}</td>
                  <td className="wide">{row.project_name}</td>
                  <td>{row.report_category}</td>
                  <td className="wide">{row.billing_item}</td>
                  <td className="num">{row.quantity ?? ''}</td>
                  <td className="num">{fmtMoney(row.source_unit_price)}</td>
                  <td>{row.selected_system}</td>
                  <td className="num">{fmtMoney(row.selected_price)}</td>
                  <td className="num">{fmtMoney(row.settlement_amount)}</td>
                  <td>{row.matched_code}</td>
                  <td>
                    <Chip className={`status-chip status-chip-${statusClass[row.status] || 'default'}`} size="sm" color={(statusClass[row.status] || 'default') as never} variant="soft">
                      {row.manual_match && row.status === '已匹配' ? '手动匹配' : row.status}
                    </Chip>
                  </td>
                </tr>
              ))}
              {!filtered.length && (
                <tr><td className="empty" colSpan={12}>没有符合条件的记录</td></tr>
              )}
            </tbody>
          </table>
        </div>
        <div className="table-foot pager">
          <span>显示 {(currentPage - 1) * pageSize + (filtered.length ? 1 : 0)}-{Math.min(currentPage * pageSize, filtered.length)} / {filtered.length} 条，导出会使用完整筛选结果</span>
          <select className="native-select page-size-select" value={pageSize} onChange={(event) => setPageSize(Number(event.target.value))}>
            <option value={50}>每页 50 条</option>
            <option value={100}>每页 100 条</option>
            <option value={200}>每页 200 条</option>
          </select>
          <Button type="button" size="sm" variant="secondary" onPress={() => setPage((value) => Math.max(1, value - 1))} isDisabled={currentPage <= 1}>上一页</Button>
          <span>第 {currentPage} / {pageCount} 页</span>
          <Button type="button" size="sm" variant="secondary" onPress={() => setPage((value) => Math.min(pageCount, value + 1))} isDisabled={currentPage >= pageCount}>下一页</Button>
        </div>
      </Card>
        </>
      ) : activeModule === 'quoteMatch' ? (
        <section className="quote-module">
          <Card className="panel summary-panel quote-summary-panel">
            <h2 className="summary-title">报价匹配</h2>
            <section className="summary-left">
              <div className="summary-files">
                <FilePick
                  className="summary-control summary-file"
                  label="外部报价清单"
                  value={quotePath ? basename(quotePath) : '选择报价清单'}
                  onClick={() => pickFile('quote')}
                />
                <FilePick
                  className="summary-control summary-file"
                  label="报价库"
                  value={quotePricePath ? basename(quotePricePath) : '选择报价库'}
                  onClick={() => pickFile('quotePrice')}
                />
              </div>
            </section>

            <section className="summary-right">
              <div className="summary-actions-row">
                <div className="quote-toolbar-actions">
                  <Button type="button" size="sm" variant="primary" onPress={runQuoteMatch} isDisabled={!quotePath || !quotePricePath || loading}>
                    {busyAction === 'quoteMatch' ? '匹配中' : '开始匹配'}
                  </Button>
                  <Button type="button" size="sm" variant="secondary" onPress={saveQuoteTemplate} isDisabled={loading}>
                    {busyAction === 'quoteTemplate' ? '保存中' : '下载导入模板'}
                  </Button>
                  <Button type="button" size="sm" variant="primary" onPress={exportQuoteRows} isDisabled={!quoteData || loading}>
                    {busyAction === 'quoteExport' ? '导出中' : '下载结果'}
                  </Button>
                </div>
              </div>
              <section className="quote-metrics">
                <button className={quoteStatusFilter === 'all' ? 'active' : ''} type="button" onClick={() => setQuoteStatusFilter('all')}>
                  <span>总计</span><strong>{quoteData?.summary.quote_lines || 0}</strong>
                </button>
                <button className={quoteStatusFilter === 'auto' ? 'active' : ''} type="button" onClick={() => setQuoteStatusFilter('auto')}>
                  <span>自动匹配</span><strong>{quoteData?.summary.auto || 0}</strong>
                </button>
                <button className={quoteStatusFilter === 'review' ? 'active' : ''} type="button" onClick={() => setQuoteStatusFilter('review')}>
                  <span>待确认</span><strong>{quoteData?.summary.review || 0}</strong>
                </button>
                <button className={quoteStatusFilter === 'unmatched' ? 'active' : ''} type="button" onClick={() => setQuoteStatusFilter('unmatched')}>
                  <span>未匹配</span><strong>{quoteData?.summary.unmatched || 0}</strong>
                </button>
                <button className={quoteStatusFilter === 'manual' ? 'active' : ''} type="button" onClick={() => setQuoteStatusFilter('manual')}>
                  <span>手动确认</span><strong>{quoteData?.summary.manual || 0}</strong>
                </button>
              </section>
            </section>
          </Card>

          <Card className="panel quote-result-panel">
            <div className="quote-result-head">
              <div>
                <h2>报价匹配结果</h2>
                <p>{quoteData ? `${quoteData.quote_name} / 报价库：${quoteData.price_name}` : '请先选择外部报价清单并开始匹配'}</p>
              </div>
              <span>显示 {visibleQuoteRows.length} / {quoteData?.matches.length || 0} 条</span>
            </div>
            <div className="quote-table-wrap">
              <table className="quote-table">
                <thead>
                  <tr>
                    <th>序号</th>
                    <th>检测材料</th>
                    <th>检测参数</th>
                    <th className="num">组/点数</th>
                    <th className="num">单价</th>
                    <th className="num">合价</th>
                    <th>备注</th>
                    <th>状态</th>
                    <th>匹配项</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleQuoteRows.map((row) => (
                    <tr key={row.id}>
                      <td>{row.seq}</td>
                      <td>{row.sample_name}</td>
                      <td className="wide">{row.parameter || row.project_name}</td>
                      <td className="num">{row.quantity ?? ''}</td>
                      <td className="num">{row.matched_price_text || fmtMoney(row.matched_price)}</td>
                      <td className="num">{fmtMoney(row.calculated_total)}</td>
                      <td className="wide">{row.matched?.remark || row.remark}</td>
                      <td><QuoteStatusChip row={row} /></td>
                      <td>
                        <Button type="button" size="sm" variant={row.matched ? 'secondary' : 'primary'} onPress={() => openQuotePicker(row)}>
                          {row.matched_label || '选择匹配项'}
                        </Button>
                      </td>
                    </tr>
                  ))}
                  {!visibleQuoteRows.length && (
                    <tr><td className="empty" colSpan={9}>{quoteData ? '没有符合当前状态的记录' : '暂无报价匹配结果'}</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </Card>
        </section>
      ) : (
        <PriceBookModule
          activeTab={priceBookTab}
          settlementBook={settlementBook}
          quoteBook={quoteBook}
          pricePath={pricePath}
          quotePricePath={quotePricePath}
          query={priceBookQueries[priceBookTab]}
          busy={priceBookBusy}
          onTabChange={setPriceBookTab}
          onQueryChange={(query) => setPriceBookQueries((queries) => ({ ...queries, [priceBookTab]: query }))}
          onSave={savePriceBook}
          onSettlementChange={setSettlementBook}
          onQuoteChange={setQuoteBook}
          onPickSettlement={() => pickFile('price')}
          onPickQuote={() => pickFile('quotePrice')}
        />
      )}

        </section>
      </div>

      {quotePickRow && quoteData && (
        <QuoteMatchDialog
          candidates={quoteCandidates}
          mode={quotePickMode}
          query={quotePickQuery}
          row={quotePickRow}
          tab={quotePickTab}
          tabs={['推荐', ...quoteData.tabs]}
          onChangeMode={setQuotePickMode}
          onChangeQuery={setQuotePickQuery}
          onChangeTab={setQuotePickTab}
          onClose={closeQuotePicker}
          onSelect={chooseQuoteMatch}
        />
      )}
      {settlementEditRow && settlementEditDraft && (
        <SettlementEditDialog
          draft={settlementEditDraft}
          row={settlementEditRow}
          onChange={setSettlementEditDraft}
          onClose={closeSettlementEditor}
          onSave={saveSettlementEdit}
        />
      )}
    </main>
  )
}

function PriceBookModule({
  activeTab,
  settlementBook,
  quoteBook,
  pricePath,
  quotePricePath,
  query,
  busy,
  onTabChange,
  onQueryChange,
  onSave,
  onSettlementChange,
  onQuoteChange,
  onPickSettlement,
  onPickQuote,
}: {
  activeTab: 'settlement' | 'quote'
  settlementBook: { items: PriceItem[]; systems: PriceSystem[] } | null
  quoteBook: QuotePriceItem[]
  pricePath: string
  quotePricePath: string
  query: string
  busy: boolean
  onTabChange: (tab: 'settlement' | 'quote') => void
  onQueryChange: (query: string) => void
  onSave: () => void
  onSettlementChange: (value: { items: PriceItem[]; systems: PriceSystem[] } | null) => void
  onQuoteChange: (value: QuotePriceItem[]) => void
  onPickSettlement: () => void
  onPickQuote: () => void
}) {
  const [editing, setEditing] = useState<{ kind: 'settlement'; item: PriceItem } | { kind: 'quote'; item: QuotePriceItem } | null>(null)
  const [draftPrices, setDraftPrices] = useState<Record<string, string>>({})
  const [draftRemark, setDraftRemark] = useState('')
  const settlementRows = (settlementBook?.items || []).filter((item) => !query || [item.report_category, item.billing_item, item.code, item.category].join(' ').toLowerCase().includes(query.toLowerCase()))
  const quoteRows = quoteBook.filter((item) => !query || item.search_text.toLowerCase().includes(query.toLowerCase()))
  const parsePriceDraft = (raw: string) => {
    const value = raw.trim()
    if (!/^-?\d*(?:\.\d*)?$/.test(value)) return null
    if (value === '' || value === '-' || value === '.' || value === '-.') return null
    const parsed = Number(value)
    return Number.isFinite(parsed) ? parsed : null
  }
  function openSettlementEditor(item: PriceItem) {
    setEditing({ kind: 'settlement', item })
    setDraftPrices(Object.fromEntries((settlementBook?.systems || []).map((system) => [system.name, item.prices[system.name] === null ? '' : String(item.prices[system.name] ?? '')])))
  }
  function openQuoteEditor(item: QuotePriceItem) {
    setEditing({ kind: 'quote', item })
    setDraftPrices({ quote: item.price === null ? '' : String(item.price ?? '') })
    setDraftRemark(item.remark)
  }
  function closeEditor() { setEditing(null); setDraftPrices({}); setDraftRemark('') }
  function applyEditor() {
    if (!editing) return
    if (editing.kind === 'settlement' && settlementBook) {
      onSettlementChange({ ...settlementBook, items: settlementBook.items.map((item) => item.sheet === editing.item.sheet && item.row_number === editing.item.row_number ? { ...item, prices: Object.fromEntries(settlementBook.systems.map((system) => [system.name, parsePriceDraft(draftPrices[system.name] || '')])) } : item) })
    } else if (editing.kind === 'quote') {
      const raw = draftPrices.quote || ''
      onQuoteChange(quoteBook.map((item) => item.sheet === editing.item.sheet && item.row_number === editing.item.row_number ? { ...item, price: parsePriceDraft(raw), raw_price: raw, remark: draftRemark } : item))
    }
    closeEditor()
  }
  function editingTitle() {
    if (!editing) return ''
    if (editing.kind === 'settlement') return `${editing.item.category || editing.item.sheet} / ${editing.item.billing_item || '-'}`
    return `${editing.item.category || editing.item.sheet} / ${editing.item.parameter || '-'}`
  }
  return (
    <section className="price-book-module">
      <Card className="panel price-book-panel">
        <div className="price-book-tabs" role="tablist">
          <button className={activeTab === 'settlement' ? 'active' : ''} type="button" role="tab" aria-selected={activeTab === 'settlement'} onClick={() => onTabChange('settlement')}>结算价格体系</button>
          <button className={activeTab === 'quote' ? 'active' : ''} type="button" role="tab" aria-selected={activeTab === 'quote'} onClick={() => onTabChange('quote')}>报价库</button>
        </div>
        <div className="price-book-source">
          <span>{activeTab === 'settlement' ? (pricePath ? basename(pricePath) : '尚未选择结算价格表') : (quotePricePath ? basename(quotePricePath) : '尚未选择报价库')}</span>
          <Button type="button" size="sm" variant="secondary" onPress={activeTab === 'settlement' ? onPickSettlement : onPickQuote}>{activeTab === 'settlement' ? '选择结算价格表' : '选择报价库'}</Button>
        </div>
        {activeTab === 'settlement' && settlementBook ? (
          <div className="price-book-tab-panel" role="tabpanel">
            <div className="price-book-toolbar">
              <div className="price-book-search"><Input value={query} onChange={(event) => onQueryChange(event.currentTarget.value)} placeholder="搜索分类、报告类别、计费项目或项目编号" /></div>
              <Button type="button" size="sm" variant="primary" onPress={onSave} isDisabled={busy}>{busy ? '保存中' : '保存修改'}</Button>
            </div>
            <div className="price-book-table-wrap">
              <table className="price-book-table">
                <thead><tr><th>分类</th><th>报告类别</th><th>计费项目</th><th>项目编号</th>{settlementBook.systems.map((system) => <th className="num" key={system.name}>{system.name}</th>)}</tr></thead>
                <tbody>{settlementRows.map((item) => <tr className="price-book-row" key={`${item.sheet}-${item.row_number}`} onClick={() => openSettlementEditor(item)}><td>{item.category || '-'}</td><td>{item.report_category || '-'}</td><td>{item.billing_item || '-'}</td><td>{item.code || '-'}</td>{settlementBook.systems.map((system) => <td className="num" key={system.name}>{item.prices[system.name] ?? '-'}</td>)}</tr>)}{!settlementRows.length && <tr><td className="empty" colSpan={4 + settlementBook.systems.length}>没有符合条件的价格项目</td></tr>}</tbody>
              </table>
            </div>
          </div>
        ) : activeTab === 'quote' ? (
          <div className="price-book-tab-panel" role="tabpanel">
            <div className="price-book-toolbar">
              <div className="price-book-search"><Input value={query} onChange={(event) => onQueryChange(event.currentTarget.value)} placeholder="搜索分类、检测材料、检测参数或报价编号" /></div>
              <Button type="button" size="sm" variant="primary" onPress={onSave} isDisabled={busy || !quoteBook.length}>{busy ? '保存中' : '保存修改'}</Button>
            </div>
            <div className="price-book-table-wrap">
              <table className="price-book-table quote-price-book-table">
                <thead><tr><th>分类</th><th>检测材料</th><th>检测参数</th><th>单位</th><th className="num">单价（元）</th><th>备注</th><th>报价编号</th></tr></thead>
                <tbody>{quoteRows.map((item) => <tr className="price-book-row" key={`${item.sheet}-${item.row_number}`} onClick={() => openQuoteEditor(item)}><td>{item.category || item.sheet}</td><td>{item.material || '-'}</td><td>{item.parameter || '-'}</td><td>{item.unit || '-'}</td><td className="num">{item.price ?? '-'}</td><td>{item.remark || '-'}</td><td>{item.code || '-'}</td></tr>)}{!quoteRows.length && <tr><td className="empty" colSpan={7}>没有符合条件的价格项目</td></tr>}</tbody>
              </table>
            </div>
          </div>
        ) : <div className="price-book-loading">{busy ? '正在读取价格文件…' : '请先选择对应的 Excel 价格文件'}</div>}
      </Card>
      {editing && <div className="price-edit-overlay" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) closeEditor() }}><section className="price-edit-dialog" role="dialog" aria-modal="true" aria-label="编辑价格"><div className="price-edit-head"><div><h2>编辑价格</h2><p>{editingTitle()}</p></div><Button type="button" size="sm" variant="secondary" onPress={closeEditor}>关闭</Button></div><div className="price-edit-body">{editing.kind === 'settlement' && settlementBook ? <>{settlementBook.systems.map((system) => <label className="price-edit-field" key={system.name}><span>{system.name}</span><Input inputMode="decimal" value={draftPrices[system.name] ?? ''} onChange={(event) => { const value = event.currentTarget.value; if (/^-?\d*(?:\.\d*)?$/.test(value)) setDraftPrices((draft) => ({ ...draft, [system.name]: value })) }} placeholder="留空表示无价格" /></label>)}</> : <><label className="price-edit-field"><span>单价（元）</span><Input inputMode="decimal" value={draftPrices.quote ?? ''} onChange={(event) => { const value = event.currentTarget.value; if (/^-?\d*(?:\.\d*)?$/.test(value)) setDraftPrices((draft) => ({ ...draft, quote: value })) }} /></label><label className="price-edit-field"><span>备注</span><Input value={draftRemark} onChange={(event) => setDraftRemark(event.currentTarget.value)} /></label></>}<div className="price-edit-footer"><Button type="button" variant="secondary" onPress={closeEditor}>取消</Button><Button type="button" variant="primary" onPress={applyEditor}>确认修改</Button></div></div></section></div>}
    </section>
  )
}

function QuoteStatusChip({ row }: { row: QuoteMatchRow }) {
  const label = row.manual_confirmed ? '手动确认' : row.match_status
  const color = quoteStatusColor(row.manual_confirmed ? '手动确认' : row.match_status)
  return (
    <Chip className={`status-chip status-chip-${color}`} size="sm" color={color as never} variant="soft">
      {label}
    </Chip>
  )
}

function SettlementEditDialog({
  draft,
  row,
  onChange,
  onClose,
  onSave,
}: {
  draft: SettlementEditDraft
  row: RecordRow
  onChange: (draft: SettlementEditDraft) => void
  onClose: () => void
  onSave: () => void
}) {
  const setField = (field: keyof SettlementEditDraft, value: string) => {
    if (!/^-?\d*(?:\.\d*)?$/.test(value)) return
    const next = { ...draft, [field]: value }
    if (field === 'unit_price' || field === 'quantity' || field === 'discount') {
      const unitPrice = parseNumeric(next.unit_price)
      const quantity = parseNumeric(next.quantity)
      const discount = parseNumeric(next.discount) ?? 0
      if (unitPrice !== null && quantity !== null) {
        const subtotal = roundCurrency(unitPrice * quantity)
        const total = roundCurrency(subtotal - discount)
        next.subtotal = valueDraft(subtotal)
        next.total = valueDraft(total)
        next.unit_billing = valueDraft(total)
      }
    }
    onChange(next)
  }

  return (
    <div className="manual-overlay" role="presentation" onMouseDown={(event) => {
      if (event.target === event.currentTarget) onClose()
    }}>
      <section className="settlement-edit-dialog" role="dialog" aria-modal="true" aria-label="编辑项目匹配结果">
        <div className="manual-head">
          <div className="manual-head-main">
            <h2>编辑匹配结果</h2>
            <div className="manual-context">
              <span>报告编号：{row.report_number || '-'}</span>
              <span>报告类别：{row.report_category || '-'}</span>
              <span>计费项目：{row.billing_item || '-'}</span>
            </div>
          </div>
          <Button type="button" size="sm" variant="secondary" onPress={onClose}>关闭</Button>
        </div>
        <div className="settlement-edit-body">
          <label><span>单价(元)</span><input inputMode="decimal" value={draft.unit_price} onChange={(event) => setField('unit_price', event.currentTarget.value)} /></label>
          <label><span>数量</span><input inputMode="decimal" value={draft.quantity} onChange={(event) => setField('quantity', event.currentTarget.value)} /></label>
          <label><span>小计</span><input inputMode="decimal" value={draft.subtotal} onChange={(event) => setField('subtotal', event.currentTarget.value)} /></label>
          <label><span>折扣(元)</span><input inputMode="decimal" value={draft.discount} onChange={(event) => setField('discount', event.currentTarget.value)} /></label>
          <label><span>总价</span><input inputMode="decimal" value={draft.total} onChange={(event) => setField('total', event.currentTarget.value)} /></label>
          <label><span>单位计费</span><input inputMode="decimal" value={draft.unit_billing} onChange={(event) => setField('unit_billing', event.currentTarget.value)} /></label>
        </div>
        <div className="settlement-edit-actions">
          <Button type="button" size="sm" variant="secondary" onPress={onClose}>取消</Button>
          <Button type="button" size="sm" variant="primary" onPress={onSave}>保存</Button>
        </div>
      </section>
    </div>
  )
}

function QuoteMatchDialog({
  candidates,
  mode,
  query,
  row,
  tab,
  tabs,
  onChangeMode,
  onChangeQuery,
  onChangeTab,
  onClose,
  onSelect,
}: {
  candidates: Array<QuotePriceItem & { score: number; method: string }>
  mode: 'fuzzy' | 'exact'
  query: string
  row: QuoteMatchRow
  tab: string
  tabs: string[]
  onChangeMode: (value: 'fuzzy' | 'exact') => void
  onChangeQuery: (value: string) => void
  onChangeTab: (value: string) => void
  onClose: () => void
  onSelect: (item: QuotePriceItem) => void
}) {
  return (
    <div className="manual-overlay" role="presentation" onMouseDown={(event) => {
      if (event.target === event.currentTarget) onClose()
    }}>
      <section className="manual-dialog" role="dialog" aria-modal="true" aria-label="手动选择匹配项目">
        <div className="manual-head">
          <div className="manual-head-main">
            <h2>选择报价匹配项</h2>
            <div className="manual-context">
              <span>检测材料：{row.sample_name || '未填写'}</span>
              <span>检测参数：{row.parameter || row.project_name || '-'}</span>
              <span>组/点数：{row.quantity ?? '-'}</span>
              <span>原始单价：{fmtMoney(row.source_price) || '-'}</span>
              {row.remark && <span>备注：{row.remark}</span>}
            </div>
          </div>
          <Button type="button" size="sm" variant="secondary" onPress={onClose}>关闭</Button>
        </div>

        <div className="manual-body">
          <section className="manual-picker">
            <div className="quote-picker-search">
              <Input value={query} onChange={(event) => onChangeQuery(event.currentTarget.value)} placeholder="搜索检测项目、检测材料、检测参数或别名" autoFocus />
              <Button type="button" size="sm" variant={mode === 'fuzzy' ? 'primary' : 'secondary'} onPress={() => onChangeMode('fuzzy')}>模糊</Button>
              <Button type="button" size="sm" variant={mode === 'exact' ? 'primary' : 'secondary'} onPress={() => onChangeMode('exact')}>精确</Button>
            </div>
            <div className="quote-tabs">
              {tabs.map((item) => (
                <button className={tab === item ? 'active' : ''} type="button" key={item} onClick={() => onChangeTab(item)}>
                  {item}
                </button>
              ))}
            </div>
            <div className="manual-table-info">
              <span>检测报价表</span>
              <small>显示 {candidates.length} 项，点击行或“选择”完成匹配</small>
            </div>
            <div className="manual-price-table-wrap">
              <table className="manual-price-table">
                <thead>
                  <tr>
                    <th>序号</th>
                    <th>分类</th>
                    <th>检测项目</th>
                    <th>检测材料</th>
                    <th>检测参数</th>
                    <th>单位</th>
                    <th className="num">单价（元）</th>
                    <th>备注</th>
                    <th>报价编号</th>
                    <th className="num">分数</th>
                    <th>操作</th>
                  </tr>
                </thead>
                <tbody>
                  {candidates.map((item) => (
                    <tr className="manual-price-row" key={`${item.sheet}-${item.row_number}`} onClick={() => onSelect(item)}>
                      <td>{item.seq}</td>
                      <td>{item.sheet}</td>
                      <td>{item.category || '-'}</td>
                      <td>{item.material || '-'}</td>
                      <td className="wide">{item.parameter || '-'}</td>
                      <td>{item.unit || '-'}</td>
                      <td className="num manual-price-cell">{item.raw_price || fmtMoney(item.price)}</td>
                      <td className="wide">{item.remark || '-'}</td>
                      <td>{item.code || '-'}</td>
                      <td className="num">{item.score}</td>
                      <td>
                        <Button type="button" size="sm" variant="primary" onPress={() => onSelect(item)}>选择</Button>
                      </td>
                    </tr>
                  ))}
                  {!candidates.length && <tr><td className="manual-empty" colSpan={11}>没有找到匹配的报价项目</td></tr>}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      </section>
    </div>
  )
}

function FilePick({ label, value, onClick, className = '' }: { label: string; value: string; onClick: () => void; className?: string }) {
  return (
    <div className={['file-pick', 'file-pick-inline', className].filter(Boolean).join(' ')}>
      <span>{label}</span>
      <Button
        className="file-pick-trigger"
        fullWidth
        size="sm"
        type="button"
        variant="secondary"
        onPress={onClick}
      >
        {value}
      </Button>
    </div>
  )
}

function Metric({ label, value, money }: { label: string; value: string; money?: boolean }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong className={money ? 'money' : ''}>{value}</strong>
    </div>
  )
}

function summarize(rows: RecordRow[]): Summary {
  return rows.reduce(
    (acc, row) => {
      acc.total += 1
      if (row.status === '已匹配') acc.matched += 1
      else acc.unresolved += 1
      if (typeof row.settlement_amount === 'number') acc.total_amount += row.settlement_amount
      acc.status_counts[row.status] = (acc.status_counts[row.status] || 0) + 1
      return acc
    },
    { total: 0, matched: 0, unresolved: 0, total_amount: 0, status_counts: {} } as Summary,
  )
}

function quoteStatusColor(status: QuoteStatus | '手动确认') {
  if (status === '自动匹配') return 'success'
  if (status === '待确认' || status === '手动确认') return 'warning'
  return 'danger'
}

function buildPriorityOrder(systems: Array<{ name: string }>, topName: string): PriorityEntry[] {
  if (!systems.length) return []
  const ordered = [...systems]
  if (topName) {
    const topIndex = ordered.findIndex((system) => system.name === topName)
    if (topIndex > 0) {
      const [top] = ordered.splice(topIndex, 1)
      ordered.unshift(top)
    }
  }
  return ordered.map((system) => ({ name: system.name, enabled: true }))
}

function readError(err: unknown) {
  if (typeof err === 'string') {
    try {
      return JSON.parse(err).message || err
    } catch {
      return err
    }
  }
  return err instanceof Error ? err.message : '处理失败'
}

function parseNumeric(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  const text = String(value ?? '').trim().replaceAll(',', '')
  return /^-?\d+(?:\.\d+)?$/.test(text) ? Number(text) : null
}

function roundCurrency(value: number): number {
  return Math.round((value + Number.EPSILON) * 100) / 100
}

function valueDraft(value: number | null | undefined): string {
  return value === null || value === undefined ? '' : String(value)
}

export default App
