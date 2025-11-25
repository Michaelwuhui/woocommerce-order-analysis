import { NextRequest, NextResponse } from 'next/server'
import { cookies } from 'next/headers'
import { createClient } from '@supabase/supabase-js'

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL!
const supabaseServiceKey = process.env.SUPABASE_SERVICE_ROLE_KEY!

const supabase = createClient(supabaseUrl, supabaseServiceKey)

export async function POST(request: NextRequest) {
  try {
    const { email, password } = await request.json()

    if (!email || !password) {
      return NextResponse.json({ error: '邮箱和密码不能为空' }, { status: 400 })
    }

    // 使用Supabase认证
    const { data: authData, error: authError } = await supabase.auth.signInWithPassword({
      email,
      password,
    })

    if (authError || !authData.user) {
      return NextResponse.json({ error: '邮箱或密码错误' }, { status: 401 })
    }

    // 获取用户信息
    const { data: userData, error: userError } = await supabase
      .from('users')
      .select('*')
      .eq('email', email)
      .single()

    if (userError || !userData) {
      return NextResponse.json({ error: '用户信息获取失败' }, { status: 500 })
    }

    // 创建简化session
    const sessionData = {
      user: {
        id: userData.id,
        email: userData.email,
        name: userData.name || userData.email,
        role: userData.role || 'viewer'
      }
    }

    // 设置cookie
    const cookieStore = cookies()
    cookieStore.set('simple-session', JSON.stringify(sessionData), {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      sameSite: 'lax',
      maxAge: 60 * 60 * 24 * 7, // 7天
      path: '/'
    })

    return NextResponse.json({ 
      success: true, 
      user: sessionData.user 
    })

  } catch (error) {
    console.error('Simple login error:', error)
    return NextResponse.json({ error: '服务器错误' }, { status: 500 })
  }
}