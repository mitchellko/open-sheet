import { col, div, ref, Sheet, sub, Table, Workbook } from '@open-sheet/core'

const quarters = [
  { quarter: 'Q1', revenue: 12_400_000, cogs: 5_100_000 },
  { quarter: 'Q2', revenue: 13_900_000, cogs: 5_600_000 },
  { quarter: 'Q3', revenue: 15_200_000, cogs: 6_050_000 },
  { quarter: 'Q4', revenue: 18_650_000, cogs: 7_100_000 },
]

// The generator lifts everything between these markers into the image's left
// panel, so what the picture shows is this file rather than a copy of it.
// #region shot
const columns = [
  col('quarter', { header: 'Quarter' }),
  col('revenue', { header: 'Revenue', format: 'currency', bar: true }),
  col('cogs', { header: 'COGS', format: 'currency' }),

  col('grossProfit', {
    header: 'Gross profit',
    format: 'currency',
    formula: (r) => sub(r.cell('revenue'), r.cell('cogs')),
  }),

  col('margin', {
    header: 'Margin',
    format: 'percent',
    scale: ['#fee2e2', '#ffffff', '#dcfce7'],
    formula: (r) => div(sub(r.cell('revenue'), r.cell('cogs')), r.cell('revenue')),
  }),

  col('qoq', {
    header: 'QoQ growth',
    format: 'percent',
    highlight: { above: 0.15, fill: '#dcfce7', bold: true },
    formula: (r) => (r.isFirst ? null : sub(div(r.cell('revenue'), r.prev().cell('revenue')), 1)),
  }),
]
// #endregion shot

export default Workbook({
  children: Sheet({
    name: 'P&L',
    children: Table({
      name: 'pl',
      data: quarters,
      columns,
      total: { revenue: 'sum', cogs: 'sum', grossProfit: 'sum' },
    }),
  }),
})

export { ref }
