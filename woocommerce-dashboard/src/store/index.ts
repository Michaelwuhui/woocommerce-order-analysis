import { create } from 'zustand'
import { devtools } from 'zustand/middleware'
import { Site, Order, SyncLog } from '@/types/database'

interface AppState {
  // Sites
  sites: Site[]
  selectedSite: Site | null
  setSites: (sites: Site[]) => void
  setSelectedSite: (site: Site | null) => void
  addSite: (site: Site) => void
  updateSite: (id: string, updates: Partial<Site>) => void
  removeSite: (id: string) => void

  // Orders
  orders: Order[]
  setOrders: (orders: Order[]) => void
  addOrders: (orders: Order[]) => void

  // Sync Logs
  syncLogs: SyncLog[]
  setSyncLogs: (logs: SyncLog[]) => void
  addSyncLog: (log: SyncLog) => void

  // UI State
  sidebarOpen: boolean
  setSidebarOpen: (open: boolean) => void
  loading: boolean
  setLoading: (loading: boolean) => void
}

export const useAppStore = create<AppState>()(
  devtools(
    (set, get) => ({
      // Sites
      sites: [],
      selectedSite: null,
      setSites: (sites) => set({ sites }),
      setSelectedSite: (site) => set({ selectedSite: site }),
      addSite: (site) => set((state) => ({ sites: [...state.sites, site] })),
      updateSite: (id, updates) =>
        set((state) => ({
          sites: state.sites.map((site) =>
            site.id === id ? { ...site, ...updates } : site
          ),
        })),
      removeSite: (id) =>
        set((state) => ({
          sites: state.sites.filter((site) => site.id !== id),
        })),

      // Orders
      orders: [],
      setOrders: (orders) => set({ orders }),
      addOrders: (orders) =>
        set((state) => ({ orders: [...state.orders, ...orders] })),

      // Sync Logs
      syncLogs: [],
      setSyncLogs: (logs) => set({ syncLogs: logs }),
      addSyncLog: (log) =>
        set((state) => ({ syncLogs: [log, ...state.syncLogs] })),

      // UI State
      sidebarOpen: true,
      setSidebarOpen: (open) => set({ sidebarOpen: open }),
      loading: false,
      setLoading: (loading) => set({ loading }),
    }),
    {
      name: 'woocommerce-dashboard-store',
    }
  )
)