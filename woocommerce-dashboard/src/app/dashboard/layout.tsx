'use client'

import { useSimpleAuth } from '@/hooks/useSimpleAuth'
import { useRouter } from 'next/navigation'
import { useEffect } from 'react'
import { DashboardLayout } from '@/components/layout/dashboard-layout'

export default function DashboardLayoutWrapper({
  children,
}: {
  children: React.ReactNode
}) {
  const { session, loading } = useSimpleAuth()
  const router = useRouter()

  useEffect(() => {
    if (!loading && !session) {
      router.push('/simple-login')
    }
  }, [session, loading, router])

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="animate-spin rounded-full h-32 w-32 border-b-2 border-primary-600"></div>
      </div>
    )
  }

  if (!session) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="animate-spin rounded-full h-32 w-32 border-b-2 border-primary-600"></div>
      </div>
    )
  }

  return <DashboardLayout>{children}</DashboardLayout>
}