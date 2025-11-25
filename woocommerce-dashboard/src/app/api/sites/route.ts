import { NextRequest, NextResponse } from 'next/server'
import { getSession } from '@/lib/simple-auth'
import { sitesService } from '@/lib/database'

export async function GET(request: NextRequest) {
  try {
    console.log('Sites API: 开始处理请求')
    const session = await getSession()
    console.log('Sites API: 获取到的session:', session)
    
    if (!session?.user?.id) {
      console.log('Sites API: 未授权，session无效')
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    console.log('Sites API: 查询用户ID:', session.user.id)
    const sites = await sitesService.getAll(session.user.id)
    console.log('Sites API: 查询到的站点数据:', sites)
    
    return NextResponse.json(sites)
  } catch (error) {
    console.error('Failed to fetch sites:', error)
    return NextResponse.json(
      { error: 'Failed to fetch sites' },
      { status: 500 }
    )
  }
}

export async function POST(request: NextRequest) {
  try {
    const session = await getSession()
    if (!session?.user?.id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const body = await request.json()
    const { name, url, woo_url, consumer_key, consumer_secret, status } = body

    if (!name || !url || !woo_url || !consumer_key || !consumer_secret) {
      return NextResponse.json(
        { error: 'Missing required fields' },
        { status: 400 }
      )
    }

    const site = await sitesService.create({
      name,
      url,
      woo_url,
      consumer_key,
      consumer_secret,
      status: status || 'active',
      user_id: session.user.id
    })

    return NextResponse.json(site, { status: 201 })
  } catch (error) {
    console.error('Failed to create site:', error)
    return NextResponse.json(
      { error: 'Failed to create site' },
      { status: 500 }
    )
  }
}

export async function PUT(request: NextRequest) {
  try {
    const session = await getServerSession(authOptions)
    if (!session?.user?.id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const body = await request.json()
    const { id, name, url, woo_url, consumer_key, consumer_secret, status } = body

    if (!id) {
      return NextResponse.json(
        { error: 'Site ID is required' },
        { status: 400 }
      )
    }

    const site = await sitesService.update(id, {
      name,
      url,
      woo_url,
      consumer_key,
      consumer_secret,
      status
    })

    return NextResponse.json(site)
  } catch (error) {
    console.error('Failed to update site:', error)
    return NextResponse.json(
      { error: 'Failed to update site' },
      { status: 500 }
    )
  }
}

export async function DELETE(request: NextRequest) {
  try {
    const session = await getServerSession(authOptions)
    if (!session?.user?.id) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { searchParams } = new URL(request.url)
    const id = searchParams.get('id')

    if (!id) {
      return NextResponse.json(
        { error: 'Site ID is required' },
        { status: 400 }
      )
    }

    await sitesService.delete(id)
    return NextResponse.json({ success: true })
  } catch (error) {
    console.error('Failed to delete site:', error)
    return NextResponse.json(
      { error: 'Failed to delete site' },
      { status: 500 }
    )
  }
}