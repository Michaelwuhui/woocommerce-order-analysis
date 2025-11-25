import { NextRequest, NextResponse } from 'next/server'
import { getServerSession } from 'next-auth'
import { authOptions } from '@/lib/auth'
import { sitesService, syncLogsService, ordersService } from '@/lib/database'
import WooCommerceRestApi from 'woocommerce-rest-api'

export async function POST(request: NextRequest) {
  try {
    const session = await getServerSession(authOptions)
    if (!session?.user?.id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { siteId } = await request.json()
    
    if (!siteId) {
      return NextResponse.json(
        { error: 'Site ID is required' },
        { status: 400 }
      )
    }

    // Get site details
    const site = await sitesService.getById(siteId)
    if (!site) {
      return NextResponse.json(
        { error: 'Site not found' },
        { status: 404 }
      )
    }

    // Create sync log entry
    const syncLog = await syncLogsService.create({
      site_id: siteId,
      status: 'running',
      started_at: new Date().toISOString()
    })

    try {
      // Initialize WooCommerce API
      const WooCommerce = new WooCommerceRestApi({
        url: site.woo_url,
        consumerKey: site.consumer_key,
        consumerSecret: site.consumer_secret,
        version: 'wc/v3'
      })

      // Fetch orders from WooCommerce
      const response = await WooCommerce.get('orders', {
        per_page: 100,
        status: 'any'
      })

      const orders = response.data
      let syncedCount = 0

      // Process and save orders
      for (const order of orders) {
        try {
          await ordersService.upsert({
            order_id: order.id,
            site_id: siteId,
            status: order.status,
            total: parseFloat(order.total),
            currency: order.currency,
            date_created: order.date_created,
            date_modified: order.date_modified,
            customer_id: order.customer_id || null,
            customer_email: order.billing?.email || '',
            customer_name: `${order.billing?.first_name || ''} ${order.billing?.last_name || ''}`.trim(),
            billing_country: order.billing?.country || '',
            billing_state: order.billing?.state || '',
            billing_city: order.billing?.city || '',
            billing_address_1: order.billing?.address_1 || '',
            billing_address_2: order.billing?.address_2 || '',
            billing_postcode: order.billing?.postcode || '',
            billing_phone: order.billing?.phone || '',
            shipping_country: order.shipping?.country || '',
            shipping_state: order.shipping?.state || '',
            shipping_city: order.shipping?.city || '',
            shipping_address_1: order.shipping?.address_1 || '',
            shipping_address_2: order.shipping?.address_2 || '',
            shipping_postcode: order.shipping?.postcode || '',
            payment_method: order.payment_method || '',
            payment_method_title: order.payment_method_title || '',
            transaction_id: order.transaction_id || '',
            line_items: order.line_items || [],
            shipping_lines: order.shipping_lines || [],
            tax_lines: order.tax_lines || [],
            fee_lines: order.fee_lines || [],
            coupon_lines: order.coupon_lines || [],
            meta_data: order.meta_data || []
          })
          syncedCount++
        } catch (orderError) {
          console.error(`Failed to sync order ${order.id}:`, orderError)
        }
      }

      // Update sync log with success
      await syncLogsService.update(syncLog.id, {
        status: 'completed',
        completed_at: new Date().toISOString(),
        records_synced: syncedCount,
        message: `Successfully synced ${syncedCount} orders`
      })

      return NextResponse.json({
        success: true,
        syncedCount,
        totalOrders: orders.length
      })

    } catch (syncError) {
      // Update sync log with error
      await syncLogsService.update(syncLog.id, {
        status: 'failed',
        completed_at: new Date().toISOString(),
        error_message: syncError instanceof Error ? syncError.message : 'Unknown error'
      })

      throw syncError
    }

  } catch (error) {
    console.error('Sync failed:', error)
    return NextResponse.json(
      { error: 'Sync failed', details: error instanceof Error ? error.message : 'Unknown error' },
      { status: 500 }
    )
  }
}

export async function GET(request: NextRequest) {
  try {
    const session = await getServerSession(authOptions)
    if (!session?.user?.id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { searchParams } = new URL(request.url)
    const siteId = searchParams.get('siteId')

    const logs = await syncLogsService.getAll(siteId || undefined)
    return NextResponse.json(logs)
  } catch (error) {
    console.error('Failed to fetch sync logs:', error)
    return NextResponse.json(
      { error: 'Failed to fetch sync logs' },
      { status: 500 }
    )
  }
}