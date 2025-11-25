import { NextRequest, NextResponse } from 'next/server'
import initSqlJs from 'sql.js'
import fs from 'fs'
import path from 'path'

async function loadDatabase() {
  const SQL = await initSqlJs({
    locateFile: (file) => path.join(process.cwd(), 'node_modules/sql.js/dist', file),
  })
  const dbPath = path.resolve(process.cwd(), '..', 'woocommerce_orders.db')
  const fileBuffer = fs.readFileSync(dbPath)
  return new SQL.Database(new Uint8Array(fileBuffer))
}

function toNumber(v: any) {
  const n = Number(v)
  return Number.isFinite(n) ? n : 0
}

export async function GET(request: NextRequest) {
  try {
    const url = new URL(request.url)
    const start = url.searchParams.get('start') || ''
    const db = await loadDatabase()

    const rows: any[] = []
    const query = `SELECT id, date_created, source, status, total, line_items FROM orders ${
      start ? `WHERE date_created >= '${start}'` : ''
    }`
    const result = db.exec(query)
    if (result.length === 0) {
      return NextResponse.json([])
    }

    const cols = result[0].columns
    const values = result[0].values

    const data = values.map((row) => {
      const obj: Record<string, any> = {}
      cols.forEach((c, i) => (obj[c] = row[i]))
      return obj
    })

    const grouped = new Map<string, any[]>()
    for (const r of data) {
      const key = `${String(r.date_created).slice(0, 7)}|${r.source}`
      const arr = grouped.get(key) || []
      arr.push(r)
      grouped.set(key, arr)
    }

    const out: any[] = []
    for (const [key, list] of grouped.entries()) {
      const [month, source] = key.split('|')
      let totalOrders = list.length
      let totalProducts = 0
      let totalAmount = 0
      let failedOrders = 0
      let failedProducts = 0
      let failedAmount = 0
      let cancelledOrders = 0
      let cancelledProducts = 0
      let cancelledAmount = 0
      let successOrders = 0
      let successProducts = 0
      let successAmount = 0

      for (const r of list) {
        let qty = 0
        try {
          if (r.line_items) {
            const items = JSON.parse(String(r.line_items))
            qty = items.reduce((s: number, it: any) => s + (Number(it.quantity) || 0), 0)
          }
        } catch {
          qty = 0
        }
        const amount = toNumber(r.total)
        const status = String(r.status).toLowerCase()
        totalProducts += qty
        totalAmount += amount
        if (status === 'failed') {
          failedOrders += 1
          failedProducts += qty
          failedAmount += amount
        } else if (status === 'cancelled') {
          cancelledOrders += 1
          cancelledProducts += qty
          cancelledAmount += amount
        } else {
          successOrders += 1
          successProducts += qty
          successAmount += amount
        }
      }

      const completion = Number(((totalProducts / 2000) * 100).toFixed(2))
      out.push({
        月份: month,
        来源网站: source,
        总订单数: totalOrders,
        总产品数量: totalProducts,
        总订单总额: Number(totalAmount.toFixed(2)),
        目标完成度: completion,
        失败订单总数: failedOrders,
        失败产品数量: failedProducts,
        失败订单总金额: Number(failedAmount.toFixed(2)),
        取消订单总数: cancelledOrders,
        取消产品数量: cancelledProducts,
        取消订单总金额: Number(cancelledAmount.toFixed(2)),
        成功订单数: successOrders,
        成功产品数量: successProducts,
        成功销售金额: Number(successAmount.toFixed(2)),
      })
    }

    // 按月份和网站排序
    out.sort((a, b) => (a.月份 === b.月份 ? a.来源网站.localeCompare(b.来源网站) : a.月份.localeCompare(b.月份)))
    return NextResponse.json(out)
  } catch (error) {
    console.error('Monthly stats error:', error)
    return NextResponse.json({ error: 'Failed to build monthly stats' }, { status: 500 })
  }
}

