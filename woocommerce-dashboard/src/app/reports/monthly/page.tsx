'use client'

import { useEffect, useState } from 'react'
import { useRouter } from 'next/navigation'

type Row = {
  月份: string
  来源网站: string
  总订单数: number
  总产品数量: number
  总订单总额: number
  目标完成度: number
  失败订单总数: number
  失败产品数量: number
  失败订单总金额: number
  取消订单总数: number
  取消产品数量: number
  取消订单总金额: number
  成功订单数: number
  成功产品数量: number
  成功销售金额: number
}

export default function MonthlyReportPage() {
  const [data, setData] = useState<Row[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const router = useRouter()

  useEffect(() => {
    const load = async () => {
      try {
        const res = await fetch('/api/stats/monthly?start=2025-10-01')
        if (!res.ok) {
          throw new Error('加载统计失败')
        }
        const json = await res.json()
        setData(json)
      } catch (e: any) {
        setError(e.message || '加载失败')
      } finally {
        setLoading(false)
      }
    }
    load()
  }, [])

  if (loading) {
    return (
      <div className="p-8">加载中...</div>
    )
  }

  if (error) {
    return (
      <div className="p-8 text-red-600">{error}</div>
    )
  }

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-4">
        <h1 className="text-xl font-semibold">月度统计（来源：SQLite）</h1>
        <button className="px-3 py-2 bg-primary-600 text-white rounded" onClick={() => router.refresh()}>刷新</button>
      </div>
      <div className="overflow-auto border rounded">
        <table className="min-w-full text-sm">
          <thead className="bg-gray-100">
            <tr>
              {['月份','来源网站','总订单数','总产品数量','总订单总额','目标完成度(%)','失败订单总数','失败产品数量','失败订单总金额','取消订单总数','取消产品数量','取消订单总金额','成功订单数','成功产品数量','成功销售金额'].map((h) => (
                <th key={h} className="px-3 py-2 text-left whitespace-nowrap">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, idx) => (
              <tr key={idx} className={idx % 2 ? 'bg-white' : 'bg-gray-50'}>
                <td className="px-3 py-2">{row.月份}</td>
                <td className="px-3 py-2">{row.来源网站}</td>
                <td className="px-3 py-2">{row.总订单数}</td>
                <td className="px-3 py-2">{row.总产品数量}</td>
                <td className="px-3 py-2">{row.总订单总额.toFixed(2)}</td>
                <td className="px-3 py-2">{row.目标完成度.toFixed(2)}</td>
                <td className="px-3 py-2">{row.失败订单总数}</td>
                <td className="px-3 py-2">{row.失败产品数量}</td>
                <td className="px-3 py-2">{row.失败订单总金额.toFixed(2)}</td>
                <td className="px-3 py-2">{row.取消订单总数}</td>
                <td className="px-3 py-2">{row.取消产品数量}</td>
                <td className="px-3 py-2">{row.取消订单总金额.toFixed(2)}</td>
                <td className="px-3 py-2">{row.成功订单数}</td>
                <td className="px-3 py-2">{row.成功产品数量}</td>
                <td className="px-3 py-2">{row.成功销售金额.toFixed(2)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

