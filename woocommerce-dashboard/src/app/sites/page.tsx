'use client'

import { useState } from 'react'
import { useSimpleAuth } from '@/hooks/useSimpleAuth'
import { PlusIcon, PencilIcon, TrashIcon } from '@heroicons/react/24/outline'
import DataTable, { Column } from '@/components/ui/data-table'
import Modal, { ConfirmDialog } from '@/components/ui/modal'
import { Button, Input, Select, Form } from '@/components/ui/form'
import { useSites, useCreateSite, useUpdateSite, useDeleteSite } from '@/hooks/api'
import { Database } from '@/types/database'

type Site = Database['public']['Tables']['sites']['Row']

export default function SitesPage() {
  const { session, loading: authLoading } = useSimpleAuth()
  const { data: sites = [], isLoading: loading } = useSites()
  const createSiteMutation = useCreateSite()
  const updateSiteMutation = useUpdateSite()
  const deleteSiteMutation = useDeleteSite()
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [isDeleteDialogOpen, setIsDeleteDialogOpen] = useState(false)
  const [selectedSite, setSelectedSite] = useState<Site | null>(null)
  const [formData, setFormData] = useState({
    name: '',
    url: '',
    woo_url: '',
    consumer_key: '',
    consumer_secret: '',
    status: 'active' as 'active' | 'inactive'
  })

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!session?.user?.id) return

    try {
      if (selectedSite) {
        // 更新站点
        await updateSiteMutation.mutateAsync({
          id: selectedSite.id,
          data: formData
        })
      } else {
        // 创建新站点
        await createSiteMutation.mutateAsync({
          ...formData,
          user_id: session.user.id
        })
      }
      
      handleCloseModal()
    } catch (error) {
      console.error('Failed to save site:', error)
    }
  }

  const handleEdit = (site: Site) => {
    setSelectedSite(site)
    setFormData({
      name: site.name,
      url: site.url,
      woo_url: site.woo_url,
      consumer_key: site.consumer_key,
      consumer_secret: site.consumer_secret,
      status: site.status
    })
    setIsModalOpen(true)
  }

  const handleDelete = async () => {
    if (!selectedSite) return

    try {
      await deleteSiteMutation.mutateAsync(selectedSite.id)
      setSelectedSite(null)
      setIsDeleteDialogOpen(false)
    } catch (error) {
      console.error('Failed to delete site:', error)
    }
  }

  const handleCloseModal = () => {
    setIsModalOpen(false)
    setSelectedSite(null)
    setFormData({
      name: '',
      url: '',
      woo_url: '',
      consumer_key: '',
      consumer_secret: '',
      status: 'active'
    })
  }

  const columns: Column<Site>[] = [
    {
      key: 'name',
      label: '站点名称',
      sortable: true
    },
    {
      key: 'url',
      label: '站点URL',
      sortable: true,
      render: (value) => (
        <a href={value} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:text-blue-800">
          {value}
        </a>
      )
    },
    {
      key: 'status',
      label: '状态',
      sortable: true,
      render: (value) => (
        <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
          value === 'active' 
            ? 'bg-green-100 text-green-800' 
            : 'bg-red-100 text-red-800'
        }`}>
          {value === 'active' ? '活跃' : '停用'}
        </span>
      )
    },
    {
      key: 'last_sync',
      label: '最后同步',
      sortable: true,
      render: (value) => value ? new Date(value).toLocaleString() : '从未同步'
    },
    {
      key: 'created_at',
      label: '创建时间',
      sortable: true,
      render: (value) => new Date(value).toLocaleString()
    },
    {
      key: 'id',
      label: '操作',
      render: (_, row) => (
        <div className="flex space-x-2">
          <button
            onClick={() => handleEdit(row)}
            className="text-blue-600 hover:text-blue-800"
          >
            <PencilIcon className="h-4 w-4" />
          </button>
          <button
            onClick={() => {
              setSelectedSite(row)
              setIsDeleteDialogOpen(true)
            }}
            className="text-red-600 hover:text-red-800"
          >
            <TrashIcon className="h-4 w-4" />
          </button>
        </div>
      )
    }
  ]

  // 如果正在加载认证状态，显示加载中
  if (authLoading) {
    return (
      <div className="p-6 flex justify-center items-center min-h-64">
        <div className="text-gray-500">加载中...</div>
      </div>
    )
  }

  // 如果未登录，显示登录提示
  if (!session) {
    return (
      <div className="p-6 flex flex-col justify-center items-center min-h-64">
        <div className="text-gray-500 mb-4">请先登录以查看站点数据</div>
        <Button onClick={() => window.location.href = '/simple-login'}>
          前往登录
        </Button>
      </div>
    )
  }

  return (
    <div className="p-6">
      <div className="flex justify-between items-center mb-6">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">站点管理</h1>
          <p className="text-gray-600">管理您的WooCommerce站点</p>
        </div>
        <Button
          onClick={() => setIsModalOpen(true)}
          className="flex items-center space-x-2"
        >
          <PlusIcon className="h-4 w-4" />
          <span>添加站点</span>
        </Button>
      </div>

      <DataTable
        data={sites}
        columns={columns}
        loading={loading}
        emptyMessage="暂无站点数据"
      />

      {/* 添加/编辑站点模态框 */}
      <Modal
        isOpen={isModalOpen}
        onClose={handleCloseModal}
        title={selectedSite ? '编辑站点' : '添加站点'}
        size="lg"
      >
        <Form onSubmit={handleSubmit}>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <Input
              label="站点名称"
              value={formData.name}
              onChange={(e) => setFormData({ ...formData, name: e.target.value })}
              required
            />
            <Input
              label="站点URL"
              type="url"
              value={formData.url}
              onChange={(e) => setFormData({ ...formData, url: e.target.value })}
              placeholder="https://example.com"
              required
            />
            <Input
              label="WooCommerce API URL"
              type="url"
              value={formData.woo_url}
              onChange={(e) => setFormData({ ...formData, woo_url: e.target.value })}
              placeholder="https://example.com/wp-json/wc/v3"
              required
            />
            <Select
              label="状态"
              value={formData.status}
              onChange={(e) => setFormData({ ...formData, status: e.target.value as 'active' | 'inactive' })}
              options={[
                { value: 'active', label: '活跃' },
                { value: 'inactive', label: '停用' }
              ]}
            />
            <Input
              label="Consumer Key"
              value={formData.consumer_key}
              onChange={(e) => setFormData({ ...formData, consumer_key: e.target.value })}
              required
            />
            <Input
              label="Consumer Secret"
              type="password"
              value={formData.consumer_secret}
              onChange={(e) => setFormData({ ...formData, consumer_secret: e.target.value })}
              required
            />
          </div>
          <div className="flex justify-end space-x-3 mt-6">
            <Button
              type="button"
              variant="outline"
              onClick={handleCloseModal}
            >
              取消
            </Button>
            <Button type="submit">
              {selectedSite ? '更新' : '创建'}
            </Button>
          </div>
        </Form>
      </Modal>

      {/* 删除确认对话框 */}
      <ConfirmDialog
        isOpen={isDeleteDialogOpen}
        onClose={() => setIsDeleteDialogOpen(false)}
        onConfirm={handleDelete}
        title="删除站点"
        message={`确定要删除站点 "${selectedSite?.name}" 吗？此操作不可撤销。`}
        confirmText="删除"
        type="danger"
      />
    </div>
  )
}