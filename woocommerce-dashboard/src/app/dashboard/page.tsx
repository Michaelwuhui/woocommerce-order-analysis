'use client'

import { useState, useEffect } from 'react'
import { useSimpleAuth } from '@/hooks/useSimpleAuth'
import { 
  CurrencyDollarIcon, 
  ShoppingCartIcon, 
  UsersIcon, 
  ArrowTrendingUpIcon 
} from '@heroicons/react/24/outline'
import { LineChartComponent, BarChartComponent, PieChartComponent } from '@/components/ui/charts'

// 模拟数据
const mockStats = {
  totalRevenue: 125430.50,
  totalOrders: 1234,
  totalCustomers: 567,
  conversionRate: 3.2
}

const mockSalesData = [
  { name: '1月', value: 12000 },
  { name: '2月', value: 19000 },
  { name: '3月', value: 15000 },
  { name: '4月', value: 25000 },
  { name: '5月', value: 22000 },
  { name: '6月', value: 30000 }
]

const mockOrderStatusData = [
  { name: '已完成', value: 45 },
  { name: '处理中', value: 25 },
  { name: '待付款', value: 20 },
  { name: '已取消', value: 10 }
]

const mockTopProducts = [
  { name: '产品A', value: 150 },
  { name: '产品B', value: 120 },
  { name: '产品C', value: 100 },
  { name: '产品D', value: 80 },
  { name: '产品E', value: 60 }
]

interface StatCardProps {
  title: string
  value: string
  icon: React.ComponentType<{ className?: string }>
  color: 'blue' | 'green' | 'purple' | 'orange'
  trend?: string
}

function StatCard({ title, value, icon: Icon, color, trend }: StatCardProps) {
  const colorClasses = {
    blue: 'bg-blue-500',
    green: 'bg-green-500',
    purple: 'bg-purple-500',
    orange: 'bg-orange-500'
  }

  return (
    <div className="bg-white overflow-hidden shadow rounded-lg">
      <div className="p-5">
        <div className="flex items-center">
          <div className="flex-shrink-0">
            <div className={`p-3 rounded-md ${colorClasses[color]}`}>
              <Icon className="h-6 w-6 text-white" />
            </div>
          </div>
          <div className="ml-5 w-0 flex-1">
            <dl>
              <dt className="text-sm font-medium text-gray-500 truncate">{title}</dt>
              <dd className="flex items-baseline">
                <div className="text-2xl font-semibold text-gray-900">{value}</div>
                {trend && (
                  <div className="ml-2 flex items-baseline text-sm font-semibold text-green-600">
                    {trend}
                  </div>
                )}
              </dd>
            </dl>
          </div>
        </div>
      </div>
    </div>
  )
}

export default function DashboardPage() {
  const { session } = useSimpleAuth()
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    // 模拟数据加载
    const timer = setTimeout(() => {
      setLoading(false)
    }, 1000)

    return () => clearTimeout(timer)
  }, [])

  if (loading) {
    return (
      <div className="p-6">
        <div className="animate-pulse">
          <div className="h-8 bg-gray-200 rounded w-1/4 mb-6"></div>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 mb-8">
            {[...Array(4)].map((_, i) => (
              <div key={i} className="bg-white p-6 rounded-lg shadow">
                <div className="h-4 bg-gray-200 rounded w-3/4 mb-2"></div>
                <div className="h-8 bg-gray-200 rounded w-1/2"></div>
              </div>
            ))}
          </div>
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="bg-white p-6 rounded-lg shadow h-96"></div>
            <div className="bg-white p-6 rounded-lg shadow h-96"></div>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="p-6">
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-gray-900">
          欢迎回来，{session?.user?.name || '用户'}
        </h1>
        <p className="text-gray-600">这是您的数据分析仪表板概览</p>
      </div>

      {/* 统计卡片 */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 mb-8">
        <StatCard
          title="总收入"
          value={`¥${mockStats.totalRevenue.toLocaleString()}`}
          icon={CurrencyDollarIcon}
          color="blue"
          trend="+12.5%"
        />
        <StatCard
          title="总订单"
          value={mockStats.totalOrders.toLocaleString()}
          icon={ShoppingCartIcon}
          color="green"
          trend="+8.2%"
        />
        <StatCard
          title="客户数量"
          value={mockStats.totalCustomers.toLocaleString()}
          icon={UsersIcon}
          color="purple"
          trend="+5.1%"
        />
        <StatCard
          title="转化率"
          value={`${mockStats.conversionRate}%`}
          icon={ArrowTrendingUpIcon}
          color="orange"
          trend="+0.8%"
        />
      </div>

      {/* 图表区域 */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-lg font-medium text-gray-900 mb-4">销售趋势</h3>
          <LineChartComponent data={mockSalesData} height={300} />
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-lg font-medium text-gray-900 mb-4">订单状态分布</h3>
          <PieChartComponent data={mockOrderStatusData} height={300} />
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-lg font-medium text-gray-900 mb-4">热销产品</h3>
          <BarChartComponent data={mockTopProducts} height={300} />
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-lg font-medium text-gray-900 mb-4">最近活动</h3>
          <div className="space-y-4">
            {[
              { action: '新订单', description: '订单 #1234 已创建', time: '2分钟前' },
              { action: '数据同步', description: '站点A数据同步完成', time: '15分钟前' },
              { action: '新客户', description: '客户张三已注册', time: '1小时前' },
              { action: '订单完成', description: '订单 #1230 已完成', time: '2小时前' }
            ].map((activity, index) => (
              <div key={index} className="flex items-center space-x-3">
                <div className="flex-shrink-0">
                  <div className="h-2 w-2 bg-blue-500 rounded-full"></div>
                </div>
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-medium text-gray-900">{activity.action}</p>
                  <p className="text-sm text-gray-500">{activity.description}</p>
                </div>
                <div className="flex-shrink-0 text-sm text-gray-500">
                  {activity.time}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  )
}