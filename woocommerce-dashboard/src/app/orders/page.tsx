'use client'

import { useState, useEffect } from 'react'
import { useSimpleAuth } from '@/hooks/useSimpleAuth'
import { FunnelIcon, ArrowDownTrayIcon } from '@heroicons/react/24/outline'
import DataTable, { Column } from '@/components/ui/data-table'
import { Button, Input, Select } from '@/components/ui/form'
import { LineChartComponent, BarChartComponent } from '@/components/ui/charts'
import { ordersService, sitesService } from '@/lib/database'
import { Database } from '@/types/database'

type Order = Database['public']['Tables']['orders']['Row']
type Site = Database['public']['Tables']['sites']['Row']

export default function OrdersPage() {
  const { session } = useSimpleAuth()
  const [orders, setOrders] = useState<Order[]>([])
  const [sites, setSites] = useState<Site[]>([])
  const [loading, setLoading] = useState(true)
  const [stats, setStats] = useState({
    totalOrders: 0,
    totalRevenue: 0,
    averageOrderValue: 0,
    statusBreakdown: {} as Record<string, number>
  })
  
  // 筛选状态
  const [filters, setFilters] = useState({
    siteId: '',
    status: '',
    dateFrom: '',
    dateTo: '',
    search: ''
  })

  useEffect(() => {
    loadData()
  }, [session])

  useEffect(() => {
    loadOrders()
  }, [filters])

  const loadData = async () => {
    if (!session?.user?.id) return
    
    try {
      setLoading(true)
      const [sitesData] = await Promise.all([
        sitesService.getAll(session.user.id)
      ])
      setSites(sitesData)
      await loadOrders()
    } catch (error) {
      console.error('Failed to load data:', error)
    } finally {
      setLoading(false)
    }
  }

  const loadOrders = async () => {
    try {
      const ordersData = await ordersService.getAll(filters.siteId || undefined)
      setOrders(ordersData)
      
      const statsData = await ordersService.getStats(filters.siteId || undefined)
      setStats(statsData)
    } catch (error) {
      console.error('Failed to load orders:', error)
    }
  }

  const handleExport = () => {
    // 导出功能实现
    const csvContent = [
      ['订单ID', '站点', '状态', '总金额', '货币', '客户邮箱', '创建时间'].join(','),
      ...orders.map(order => [
        order.order_id,
        order.site_id,
        order.status,
        order.total,
        order.currency,
        order.customer_email,
        new Date(order.date_created).toLocaleString()
      ].join(','))
    ].join('\n')

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
    const link = document.createElement('a')
    link.href = URL.createObjectURL(blob)
    link.download = `orders_${new Date().toISOString().split('T')[0]}.csv`
    link.click()
  }

  const columns: Column<Order>[] = [
    {
      key: 'order_id',
      label: '订单ID',
      sortable: true
    },
    {
      key: 'customer_name',
      label: '客户姓名',
      sortable: true
    },
    {
      key: 'customer_email',
      label: '客户邮箱',
      sortable: true
    },
    {
      key: 'status',
      label: '状态',
      sortable: true,
      render: (value) => {
        const statusColors: Record<string, string> = {
          'completed': 'bg-green-100 text-green-800',
          'processing': 'bg-blue-100 text-blue-800',
          'pending': 'bg-yellow-100 text-yellow-800',
          'cancelled': 'bg-red-100 text-red-800',
          'refunded': 'bg-gray-100 text-gray-800'
        }
        
        return (
          <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
            statusColors[value] || 'bg-gray-100 text-gray-800'
          }`}>
            {value}
          </span>
        )
      }
    },
    {
      key: 'total',
      label: '总金额',
      sortable: true,
      render: (value, row) => `${row.currency} ${value.toFixed(2)}`
    },
    {
      key: 'date_created',
      label: '创建时间',
      sortable: true,
      render: (value) => new Date(value).toLocaleString()
    }
  ]

  // 模拟图表数据
  const salesTrendData = [
    { name: '1月', value: 12000 },
    { name: '2月', value: 19000 },
    { name: '3月', value: 15000 },
    { name: '4月', value: 25000 },
    { name: '5月', value: 22000 },
    { name: '6月', value: 30000 }
  ]

  const statusData = Object.entries(stats.statusBreakdown).map(([status, count]) => ({
    name: status,
    value: count
  }))

  return (
    <div className="p-6">
      <div className="flex justify-between items-center mb-6">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">订单分析</h1>
          <p className="text-gray-600">查看和分析您的订单数据</p>
        </div>
        <Button
          onClick={handleExport}
          variant="outline"
          className="flex items-center space-x-2"
        >
          <ArrowDownTrayIcon className="h-4 w-4" />
          <span>导出数据</span>
        </Button>
      </div>

      {/* 统计卡片 */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-sm font-medium text-gray-500">总订单数</h3>
          <p className="text-2xl font-bold text-gray-900">{stats.totalOrders}</p>
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-sm font-medium text-gray-500">总收入</h3>
          <p className="text-2xl font-bold text-gray-900">¥{stats.totalRevenue.toFixed(2)}</p>
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-sm font-medium text-gray-500">平均订单价值</h3>
          <p className="text-2xl font-bold text-gray-900">¥{stats.averageOrderValue.toFixed(2)}</p>
        </div>
      </div>

      {/* 图表区域 */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-lg font-medium text-gray-900 mb-4">销售趋势</h3>
          <LineChartComponent data={salesTrendData} height={300} />
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-lg font-medium text-gray-900 mb-4">订单状态分布</h3>
          <BarChartComponent data={statusData} height={300} />
        </div>
      </div>

      {/* 筛选器 */}
      <div className="bg-white p-6 rounded-lg shadow mb-6">
        <div className="flex items-center space-x-2 mb-4">
          <FunnelIcon className="h-5 w-5 text-gray-400" />
          <h3 className="text-lg font-medium text-gray-900">筛选器</h3>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
          <Select
            label="站点"
            value={filters.siteId}
            onChange={(e) => setFilters({ ...filters, siteId: e.target.value })}
            options={[
              { value: '', label: '所有站点' },
              ...sites.map(site => ({ value: site.id, label: site.name }))
            ]}
          />
          <Select
            label="状态"
            value={filters.status}
            onChange={(e) => setFilters({ ...filters, status: e.target.value })}
            options={[
              { value: '', label: '所有状态' },
              { value: 'completed', label: '已完成' },
              { value: 'processing', label: '处理中' },
              { value: 'pending', label: '待付款' },
              { value: 'cancelled', label: '已取消' }
            ]}
          />
          <Input
            label="开始日期"
            type="date"
            value={filters.dateFrom}
            onChange={(e) => setFilters({ ...filters, dateFrom: e.target.value })}
          />
          <Input
            label="结束日期"
            type="date"
            value={filters.dateTo}
            onChange={(e) => setFilters({ ...filters, dateTo: e.target.value })}
          />
          <Input
            label="搜索"
            placeholder="搜索订单..."
            value={filters.search}
            onChange={(e) => setFilters({ ...filters, search: e.target.value })}
          />
        </div>
      </div>

      {/* 订单表格 */}
      <DataTable
        data={orders}
        columns={columns}
        loading={loading}
        emptyMessage="暂无订单数据"
      />
    </div>
  )
}