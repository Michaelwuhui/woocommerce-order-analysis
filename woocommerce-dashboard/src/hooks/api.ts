import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { queryKeys } from '@/lib/react-query'
import { Site, Order, SyncLog } from '@/types/database'

// API functions
const api = {
  // Sites
  sites: {
    getAll: async (): Promise<Site[]> => {
      const response = await fetch('/api/sites')
      if (!response.ok) throw new Error('Failed to fetch sites')
      return response.json()
    },
    create: async (data: Omit<Site, 'id' | 'created_at' | 'updated_at'>): Promise<Site> => {
      const response = await fetch('/api/sites', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      })
      if (!response.ok) throw new Error('Failed to create site')
      return response.json()
    },
    update: async (id: string, data: Partial<Site>): Promise<Site> => {
      const response = await fetch('/api/sites', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id, ...data }),
      })
      if (!response.ok) throw new Error('Failed to update site')
      return response.json()
    },
    delete: async (id: string): Promise<void> => {
      const response = await fetch(`/api/sites?id=${id}`, {
        method: 'DELETE',
      })
      if (!response.ok) throw new Error('Failed to delete site')
    },
  },
  
  // Orders
  orders: {
    getAll: async (siteId?: string, limit = 50, offset = 0): Promise<Order[]> => {
      const params = new URLSearchParams({
        limit: limit.toString(),
        offset: offset.toString(),
      })
      if (siteId) params.append('siteId', siteId)
      
      const response = await fetch(`/api/orders?${params}`)
      if (!response.ok) throw new Error('Failed to fetch orders')
      return response.json()
    },
  },
  
  // Sync
  sync: {
    start: async (siteId: string) => {
      const response = await fetch('/api/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ siteId }),
      })
      if (!response.ok) throw new Error('Failed to start sync')
      return response.json()
    },
    getLogs: async (siteId?: string): Promise<SyncLog[]> => {
      const params = siteId ? `?siteId=${siteId}` : ''
      const response = await fetch(`/api/sync${params}`)
      if (!response.ok) throw new Error('Failed to fetch sync logs')
      return response.json()
    },
  },
  
  // Reports
  reports: {
    generate: async (type: string, filters: Record<string, any> = {}) => {
      const params = new URLSearchParams({ type, ...filters })
      const response = await fetch(`/api/reports?${params}`)
      if (!response.ok) throw new Error('Failed to generate report')
      return response.json()
    },
  },
  
  // WooCommerce
  woocommerce: {
    testConnection: async (siteId: string) => {
      const response = await fetch('/api/woocommerce/test', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ siteId }),
      })
      if (!response.ok) throw new Error('Failed to test connection')
      return response.json()
    },
  },
}

// Hooks
export const useSites = () => {
  return useQuery({
    queryKey: queryKeys.sites.all,
    queryFn: api.sites.getAll,
    retry: false, // 不重试，避免多次401错误
  })
}

export const useCreateSite = () => {
  const queryClient = useQueryClient()
  
  return useMutation({
    mutationFn: api.sites.create,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: queryKeys.sites.all })
    },
  })
}

export const useUpdateSite = () => {
  const queryClient = useQueryClient()
  
  return useMutation({
    mutationFn: ({ id, data }: { id: string; data: Partial<Site> }) => api.sites.update(id, data),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: queryKeys.sites.all })
    },
  })
}

export const useDeleteSite = () => {
  const queryClient = useQueryClient()
  
  return useMutation({
    mutationFn: api.sites.delete,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: queryKeys.sites.all })
    },
  })
}

export const useOrders = (siteId?: string, limit = 50, offset = 0) => {
  return useQuery({
    queryKey: siteId 
      ? queryKeys.orders.bySite(siteId)
      : queryKeys.orders.all,
    queryFn: () => api.orders.getAll(siteId, limit, offset),
  })
}

export const useStartSync = () => {
  const queryClient = useQueryClient()
  
  return useMutation({
    mutationFn: api.sync.start,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: queryKeys.sync.logs })
      queryClient.invalidateQueries({ queryKey: queryKeys.orders.all })
    },
  })
}

export const useSyncLogs = (siteId?: string) => {
  return useQuery({
    queryKey: siteId 
      ? queryKeys.sync.bySite(siteId)
      : queryKeys.sync.logs,
    queryFn: () => api.sync.getLogs(siteId),
  })
}

export const useReport = (type: string, filters: Record<string, any> = {}) => {
  return useQuery({
    queryKey: queryKeys.reports.filtered(type, filters),
    queryFn: () => api.reports.generate(type, filters),
    enabled: !!type,
  })
}

export const useTestWooConnection = () => {
  return useMutation({
    mutationFn: api.woocommerce.testConnection,
  })
}