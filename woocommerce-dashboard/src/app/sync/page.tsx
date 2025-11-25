'use client'

import { useState, useEffect } from 'react'
import { useSimpleAuth } from '@/hooks/useSimpleAuth'
import { 
  ArrowPathIcon, 
  PlayIcon, 
  StopIcon,
  ClockIcon,
  CheckCircleIcon,
  XCircleIcon
} from '@heroicons/react/24/outline'
import DataTable, { Column } from '@/components/ui/data-table'
import { Button, Select } from '@/components/ui/form'
import { sitesService, syncLogsService } from '@/lib/database'
import { Database } from '@/types/database'

type Site = Database['public']['Tables']['sites']['Row']
type SyncLog = Database['public']['Tables']['sync_logs']['Row']

export default function SyncPage() {
  const { session } = useSimpleAuth()
  const [sites, setSites] = useState<Site[]>([])
  const [syncLogs, setSyncLogs] = useState<SyncLog[]>([])
  const [loading, setLoading] = useState(true)
  const [syncing, setSyncing] = useState<Record<string, boolean>>({})
  const [selectedSite, setSelectedSite] = useState('')

  useEffect(() => {
    loadData()
  }, [session])

  const loadData = async () => {
    if (!session?.user?.id) return
    
    try {
      setLoading(true)
      const [sitesData, logsData] = await Promise.all([
        sitesService.getAll(session.user.id),
        syncLogsService.getAll()
      ])
      setSites(sitesData)
      setSyncLogs(logsData)
    } catch (error) {
      console.error('Failed to load data:', error)
    } finally {
      setLoading(false)
    }
  }

  const handleSync = async (siteId: string, type: 'orders' | 'products' | 'customers') => {
    try {
      setSyncing(prev => ({ ...prev, [`${siteId}-${type}`]: true }))
      
      // 创建同步日志
      await syncLogsService.create({
        site_id: siteId,
        type,
        status: 'running',
        records_processed: 0,
        records_total: 0,
        started_at: new Date().toISOString()
      })

      // 调用同步API
      const response = await fetch('/api/sync', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ siteId, type })
      })

      if (!response.ok) {
        throw new Error('Sync failed')
      }

      await loadData()
    } catch (error) {
      console.error('Sync error:', error)
    } finally {
      setSyncing(prev => ({ ...prev, [`${siteId}-${type}`]: false }))
    }
  }

  const handleSyncAll = async (siteId: string) => {
    const types: ('orders' | 'products' | 'customers')[] = ['orders', 'products', 'customers']
    for (const type of types) {
      await handleSync(siteId, type)
    }
  }

  const getStatusIcon = (status: string) => {
    switch (status) {
      case 'completed':
        return <CheckCircleIcon className="h-5 w-5 text-green-500" />
      case 'failed':
        return <XCircleIcon className="h-5 w-5 text-red-500" />
      case 'running':
        return <ArrowPathIcon className="h-5 w-5 text-blue-500 animate-spin" />
      default:
        return <ClockIcon className="h-5 w-5 text-gray-500" />
    }
  }

  const getStatusText = (status: string) => {
    const statusMap: Record<string, string> = {
      pending: '等待中',
      running: '运行中',
      completed: '已完成',
      failed: '失败'
    }
    return statusMap[status] || status
  }

  const columns: Column<SyncLog>[] = [
    {
      key: 'site_id',
      label: '站点',
      render: (value) => {
        const site = sites.find(s => s.id === value)
        return site?.name || value
      }
    },
    {
      key: 'type',
      label: '同步类型',
      render: (value) => {
        const typeMap: Record<string, string> = {
          orders: '订单',
          products: '产品',
          customers: '客户'
        }
        return typeMap[value] || value
      }
    },
    {
      key: 'status',
      label: '状态',
      render: (value) => (
        <div className="flex items-center space-x-2">
          {getStatusIcon(value)}
          <span>{getStatusText(value)}</span>
        </div>
      )
    },
    {
      key: 'records_processed',
      label: '进度',
      render: (value, row) => {
        const percentage = row.records_total > 0 
          ? Math.round((value / row.records_total) * 100) 
          : 0
        return (
          <div className="flex items-center space-x-2">
            <div className="w-20 bg-gray-200 rounded-full h-2">
              <div 
                className="bg-blue-600 h-2 rounded-full" 
                style={{ width: `${percentage}%` }}
              ></div>
            </div>
            <span className="text-sm text-gray-600">
              {value}/{row.records_total} ({percentage}%)
            </span>
          </div>
        )
      }
    },
    {
      key: 'started_at',
      label: '开始时间',
      sortable: true,
      render: (value) => new Date(value).toLocaleString()
    },
    {
      key: 'completed_at',
      label: '完成时间',
      render: (value) => value ? new Date(value).toLocaleString() : '-'
    },
    {
      key: 'message',
      label: '消息',
      render: (value) => value || '-'
    }
  ]

  const filteredLogs = selectedSite 
    ? syncLogs.filter(log => log.site_id === selectedSite)
    : syncLogs

  return (
    <div className="p-6">
      <div className="flex justify-between items-center mb-6">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">数据同步</h1>
          <p className="text-gray-600">管理和监控数据同步任务</p>
        </div>
      </div>

      {/* 站点同步控制 */}
      <div className="bg-white p-6 rounded-lg shadow mb-6">
        <h3 className="text-lg font-medium text-gray-900 mb-4">站点同步控制</h3>
        <div className="space-y-4">
          {sites.map(site => (
            <div key={site.id} className="flex items-center justify-between p-4 border rounded-lg">
              <div>
                <h4 className="font-medium text-gray-900">{site.name}</h4>
                <p className="text-sm text-gray-500">{site.url}</p>
                <p className="text-xs text-gray-400">
                  最后同步: {site.last_sync ? new Date(site.last_sync).toLocaleString() : '从未同步'}
                </p>
              </div>
              <div className="flex items-center space-x-2">
                <Button
                  size="sm"
                  onClick={() => handleSync(site.id, 'orders')}
                  loading={syncing[`${site.id}-orders`]}
                  disabled={site.status === 'inactive'}
                >
                  同步订单
                </Button>
                <Button
                  size="sm"
                  variant="secondary"
                  onClick={() => handleSync(site.id, 'products')}
                  loading={syncing[`${site.id}-products`]}
                  disabled={site.status === 'inactive'}
                >
                  同步产品
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => handleSync(site.id, 'customers')}
                  loading={syncing[`${site.id}-customers`]}
                  disabled={site.status === 'inactive'}
                >
                  同步客户
                </Button>
                <Button
                  size="sm"
                  onClick={() => handleSyncAll(site.id)}
                  disabled={site.status === 'inactive'}
                  className="flex items-center space-x-1"
                >
                  <ArrowPathIcon className="h-4 w-4" />
                  <span>全部同步</span>
                </Button>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* 同步日志筛选 */}
      <div className="bg-white p-6 rounded-lg shadow mb-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-lg font-medium text-gray-900">同步日志</h3>
          <div className="flex items-center space-x-4">
            <Select
              value={selectedSite}
              onChange={(e) => setSelectedSite(e.target.value)}
              options={[
                { value: '', label: '所有站点' },
                ...sites.map(site => ({ value: site.id, label: site.name }))
              ]}
            />
            <Button
              variant="outline"
              onClick={loadData}
              className="flex items-center space-x-1"
            >
              <ArrowPathIcon className="h-4 w-4" />
              <span>刷新</span>
            </Button>
          </div>
        </div>
      </div>

      {/* 同步日志表格 */}
      <DataTable
        data={filteredLogs}
        columns={columns}
        loading={loading}
        emptyMessage="暂无同步日志"
      />
    </div>
  )
}