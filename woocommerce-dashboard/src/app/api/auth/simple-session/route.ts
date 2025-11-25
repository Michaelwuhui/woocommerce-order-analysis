import { NextResponse } from 'next/server'
import { getSession } from '@/lib/simple-auth'

export async function GET() {
  try {
    const session = await getSession()
    
    if (!session) {
      return NextResponse.json({ error: 'No session' }, { status: 401 })
    }
    
    return NextResponse.json(session)
  } catch (error) {
    console.error('Session check error:', error)
    return NextResponse.json({ error: 'Session check failed' }, { status: 500 })
  }
}
