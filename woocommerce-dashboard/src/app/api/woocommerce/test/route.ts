import { NextRequest, NextResponse } from 'next/server'
import { getServerSession } from 'next-auth'
import { authOptions } from '@/lib/auth'
import { sitesService } from '@/lib/database'
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

    try {
      // Initialize WooCommerce API
      const WooCommerce = new WooCommerceRestApi({
        url: site.woo_url,
        consumerKey: site.consumer_key,
        consumerSecret: site.consumer_secret,
        version: 'wc/v3'
      })

      // Test connection by fetching system status
      const response = await WooCommerce.get('system_status')
      
      return NextResponse.json({
        success: true,
        message: 'WooCommerce connection successful',
        storeInfo: {
          name: response.data.settings?.general?.woocommerce_store_name || 'Unknown',
          version: response.data.environment?.version || 'Unknown',
          currency: response.data.settings?.general?.woocommerce_currency || 'Unknown'
        }
      })

    } catch (wooError) {
      console.error('WooCommerce API error:', wooError)
      return NextResponse.json(
        { 
          success: false,
          error: 'Failed to connect to WooCommerce API',
          details: wooError instanceof Error ? wooError.message : 'Unknown error'
        },
        { status: 400 }
      )
    }

  } catch (error) {
    console.error('Test connection failed:', error)
    return NextResponse.json(
      { error: 'Test connection failed' },
      { status: 500 }
    )
  }
}