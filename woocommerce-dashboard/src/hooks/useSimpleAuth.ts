'use client'

import { useState, useEffect } from 'react'
import { SimpleSession } from '@/lib/simple-auth'
import { getSession as getNextAuthSession } from 'next-auth/react'

export function useSimpleAuth() {
  const [session, setSession] = useState<SimpleSession | null>(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    checkSession()
  }, [])

  const checkSession = async () => {
    try {
      const response = await fetch('/api/auth/simple-session')
      if (response.ok) {
        const sessionData = await response.json()
        setSession(sessionData)
      } else {
        const nextSession = await getNextAuthSession()
        if (nextSession?.user) {
          const mapped: SimpleSession = {
            user: {
              id: (nextSession.user as any).id || 'unknown',
              email: nextSession.user.email || 'unknown',
              name: nextSession.user.name || nextSession.user.email || 'User',
              role: ((nextSession.user as any).role || 'viewer') as 'admin' | 'manager' | 'viewer',
            },
          }
          setSession(mapped)
        } else {
          setSession(null)
        }
      }
    } catch (error) {
      console.error('Failed to check session:', error)
      setSession(null)
    } finally {
      setLoading(false)
    }
  }

  const login = async (email: string, password: string) => {
    try {
      const response = await fetch('/api/auth/simple-login', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ email, password }),
      })

      if (response.ok) {
        const sessionData = await response.json()
        setSession(sessionData)
        return { success: true }
      } else {
        const error = await response.json()
        return { success: false, error: error.error }
      }
    } catch (error) {
      return { success: false, error: 'Login failed' }
    }
  }

  const logout = async () => {
    try {
      await fetch('/api/auth/logout', { method: 'POST' })
    } catch (error) {
      console.error('Logout error:', error)
    } finally {
      setSession(null)
    }
  }

  return {
    session,
    loading,
    login,
    logout,
    checkSession,
  }
}
