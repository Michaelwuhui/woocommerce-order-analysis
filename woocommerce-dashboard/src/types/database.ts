export interface Database {
  public: {
    Tables: {
      sites: {
        Row: {
          id: string
          name: string
          url: string
          woo_url: string
          consumer_key: string
          consumer_secret: string
          status: 'active' | 'inactive'
          last_sync: string | null
          created_at: string
          updated_at: string
          user_id: string
        }
        Insert: {
          id?: string
          name: string
          url: string
          woo_url: string
          consumer_key: string
          consumer_secret: string
          status?: 'active' | 'inactive'
          last_sync?: string | null
          created_at?: string
          updated_at?: string
          user_id: string
        }
        Update: {
          id?: string
          name?: string
          url?: string
          woo_url?: string
          consumer_key?: string
          consumer_secret?: string
          status?: 'active' | 'inactive'
          last_sync?: string | null
          created_at?: string
          updated_at?: string
          user_id?: string
        }
      }
      orders: {
        Row: {
          id: string
          site_id: string
          order_id: number
          status: string
          currency: string
          total: number
          date_created: string
          date_modified: string
          customer_id: number
          customer_email: string
          customer_name: string
          billing_country: string
          billing_state: string
          billing_city: string
          shipping_country: string
          shipping_state: string
          shipping_city: string
          line_items: any[]
          created_at: string
          updated_at: string
        }
        Insert: {
          id?: string
          site_id: string
          order_id: number
          status: string
          currency: string
          total: number
          date_created: string
          date_modified: string
          customer_id: number
          customer_email: string
          customer_name: string
          billing_country: string
          billing_state: string
          billing_city: string
          shipping_country: string
          shipping_state: string
          shipping_city: string
          line_items: any[]
          created_at?: string
          updated_at?: string
        }
        Update: {
          id?: string
          site_id?: string
          order_id?: number
          status?: string
          currency?: string
          total?: number
          date_created?: string
          date_modified?: string
          customer_id?: number
          customer_email?: string
          customer_name?: string
          billing_country?: string
          billing_state?: string
          billing_city?: string
          shipping_country?: string
          shipping_state?: string
          shipping_city?: string
          line_items?: any[]
          created_at?: string
          updated_at?: string
        }
      }
      sync_logs: {
        Row: {
          id: string
          site_id: string
          type: 'orders' | 'products' | 'customers'
          status: 'pending' | 'running' | 'completed' | 'failed'
          message: string | null
          records_processed: number
          records_total: number
          started_at: string
          completed_at: string | null
          created_at: string
        }
        Insert: {
          id?: string
          site_id: string
          type: 'orders' | 'products' | 'customers'
          status?: 'pending' | 'running' | 'completed' | 'failed'
          message?: string | null
          records_processed?: number
          records_total?: number
          started_at?: string
          completed_at?: string | null
          created_at?: string
        }
        Update: {
          id?: string
          site_id?: string
          type?: 'orders' | 'products' | 'customers'
          status?: 'pending' | 'running' | 'completed' | 'failed'
          message?: string | null
          records_processed?: number
          records_total?: number
          started_at?: string
          completed_at?: string | null
          created_at?: string
        }
      }
      users: {
        Row: {
          id: string
          email: string
          name: string | null
          role: 'admin' | 'user'
          created_at: string
          updated_at: string
        }
        Insert: {
          id: string
          email: string
          name?: string | null
          role?: 'admin' | 'user'
          created_at?: string
          updated_at?: string
        }
        Update: {
          id?: string
          email?: string
          name?: string | null
          role?: 'admin' | 'user'
          created_at?: string
          updated_at?: string
        }
      }
    }
    Views: {
      [_ in never]: never
    }
    Functions: {
      [_ in never]: never
    }
    Enums: {
      [_ in never]: never
    }
  }
}

// Type aliases for easier use
export type Site = Database['public']['Tables']['sites']['Row']
export type SiteInsert = Database['public']['Tables']['sites']['Insert']
export type SiteUpdate = Database['public']['Tables']['sites']['Update']

export type Order = Database['public']['Tables']['orders']['Row']
export type OrderInsert = Database['public']['Tables']['orders']['Insert']
export type OrderUpdate = Database['public']['Tables']['orders']['Update']

export type SyncLog = Database['public']['Tables']['sync_logs']['Row']
export type SyncLogInsert = Database['public']['Tables']['sync_logs']['Insert']
export type SyncLogUpdate = Database['public']['Tables']['sync_logs']['Update']

export type User = Database['public']['Tables']['users']['Row']
export type UserInsert = Database['public']['Tables']['users']['Insert']
export type UserUpdate = Database['public']['Tables']['users']['Update']