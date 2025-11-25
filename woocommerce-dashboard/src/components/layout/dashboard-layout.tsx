'use client'

import { useState } from 'react'
import { useSimpleAuth } from '@/hooks/useSimpleAuth'
import Link from 'next/link'
import { usePathname } from 'next/navigation'
import {
  Bars3Icon,
  XMarkIcon,
  HomeIcon,
  GlobeAltIcon,
  ChartBarIcon,
  DocumentTextIcon,
  ArrowPathIcon,
  UsersIcon,
  Cog6ToothIcon,
  ArrowRightOnRectangleIcon,
  UserCircleIcon,
} from '@heroicons/react/24/outline'
import { cn } from '@/lib/utils'

const navigation = [
  { name: '仪表板', href: '/dashboard', icon: HomeIcon },
  { name: '站点管理', href: '/sites', icon: GlobeAltIcon },
  { name: '订单分析', href: '/orders', icon: ChartBarIcon },
  { name: '报告中心', href: '/reports', icon: DocumentTextIcon },
  { name: '数据同步', href: '/sync', icon: ArrowPathIcon },
  { name: '用户管理', href: '/users', icon: UsersIcon },
  { name: '系统设置', href: '/settings', icon: Cog6ToothIcon },
]

interface DashboardLayoutProps {
  children: React.ReactNode
}

export function DashboardLayout({ children }: DashboardLayoutProps) {
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const { session, logout } = useSimpleAuth()
  const pathname = usePathname()

  const handleSignOut = () => {
    logout()
    window.location.href = '/simple-login'
  }

  return (
    <div className="h-full">
      {/* Mobile sidebar */}
      <div className={cn(
        'relative z-50 lg:hidden',
        sidebarOpen ? 'block' : 'hidden'
      )}>
        <div className="fixed inset-0 bg-gray-900/80" onClick={() => setSidebarOpen(false)} />
        <div className="fixed inset-0 flex">
          <div className="relative mr-16 flex w-full max-w-xs flex-1">
            <div className="absolute left-full top-0 flex w-16 justify-center pt-5">
              <button
                type="button"
                className="-m-2.5 p-2.5"
                onClick={() => setSidebarOpen(false)}
              >
                <XMarkIcon className="h-6 w-6 text-white" />
              </button>
            </div>
            <div className="flex grow flex-col gap-y-5 overflow-y-auto bg-white px-6 pb-2">
              <div className="flex h-16 shrink-0 items-center">
                <div className="h-8 w-8 bg-primary-600 rounded-lg flex items-center justify-center">
                  <span className="text-white font-bold text-sm">WC</span>
                </div>
                <span className="ml-2 text-lg font-semibold text-gray-900">
                  数据分析平台
                </span>
              </div>
              <nav className="flex flex-1 flex-col">
                <ul role="list" className="flex flex-1 flex-col gap-y-7">
                  <li>
                    <ul role="list" className="-mx-2 space-y-1">
                      {navigation.map((item) => (
                        <li key={item.name}>
                          <Link
                            href={item.href}
                            className={cn(
                              pathname === item.href
                                ? 'bg-gray-50 text-primary-600'
                                : 'text-gray-700 hover:text-primary-600 hover:bg-gray-50',
                              'group flex gap-x-3 rounded-md p-2 text-sm leading-6 font-semibold'
                            )}
                          >
                            <item.icon
                              className={cn(
                                pathname === item.href
                                  ? 'text-primary-600'
                                  : 'text-gray-400 group-hover:text-primary-600',
                                'h-6 w-6 shrink-0'
                              )}
                            />
                            {item.name}
                          </Link>
                        </li>
                      ))}
                    </ul>
                  </li>
                </ul>
              </nav>
            </div>
          </div>
        </div>
      </div>

      {/* Static sidebar for desktop */}
      <div className="hidden lg:fixed lg:inset-y-0 lg:z-50 lg:flex lg:w-72 lg:flex-col">
        <div className="flex grow flex-col gap-y-5 overflow-y-auto border-r border-gray-200 bg-white px-6">
          <div className="flex h-16 shrink-0 items-center">
            <div className="h-8 w-8 bg-primary-600 rounded-lg flex items-center justify-center">
              <span className="text-white font-bold text-sm">WC</span>
            </div>
            <span className="ml-2 text-lg font-semibold text-gray-900">
              数据分析平台
            </span>
          </div>
          <nav className="flex flex-1 flex-col">
            <ul role="list" className="flex flex-1 flex-col gap-y-7">
              <li>
                <ul role="list" className="-mx-2 space-y-1">
                  {navigation.map((item) => (
                    <li key={item.name}>
                      <Link
                        href={item.href}
                        className={cn(
                          pathname === item.href
                            ? 'bg-gray-50 text-primary-600'
                            : 'text-gray-700 hover:text-primary-600 hover:bg-gray-50',
                          'group flex gap-x-3 rounded-md p-2 text-sm leading-6 font-semibold'
                        )}
                      >
                        <item.icon
                          className={cn(
                            pathname === item.href
                              ? 'text-primary-600'
                              : 'text-gray-400 group-hover:text-primary-600',
                            'h-6 w-6 shrink-0'
                          )}
                        />
                        {item.name}
                      </Link>
                    </li>
                  ))}
                </ul>
              </li>
              <li className="-mx-6 mt-auto">
                <div className="flex items-center gap-x-4 px-6 py-3 text-sm font-semibold leading-6 text-gray-900">
                  <UserCircleIcon className="h-8 w-8 text-gray-400" />
                  <span className="sr-only">Your profile</span>
                  <span aria-hidden="true">{session?.user?.name}</span>
                  <button
                    onClick={handleSignOut}
                    className="ml-auto text-gray-400 hover:text-gray-600"
                    title="退出登录"
                  >
                    <ArrowRightOnRectangleIcon className="h-5 w-5" />
                  </button>
                </div>
              </li>
            </ul>
          </nav>
        </div>
      </div>

      <div className="lg:pl-72">
        {/* Top bar */}
        <div className="sticky top-0 z-40 flex h-16 shrink-0 items-center gap-x-4 border-b border-gray-200 bg-white px-4 shadow-sm sm:gap-x-6 sm:px-6 lg:px-8">
          <button
            type="button"
            className="-m-2.5 p-2.5 text-gray-700 lg:hidden"
            onClick={() => setSidebarOpen(true)}
          >
            <Bars3Icon className="h-6 w-6" />
          </button>

          <div className="h-6 w-px bg-gray-200 lg:hidden" />

          <div className="flex flex-1 gap-x-4 self-stretch lg:gap-x-6">
            <div className="relative flex flex-1">
              {/* 可以在这里添加搜索框 */}
            </div>
            <div className="flex items-center gap-x-4 lg:gap-x-6">
              {/* 用户信息 */}
              <div className="hidden lg:block lg:h-6 lg:w-px lg:bg-gray-200" />
              <div className="flex items-center gap-x-2">
                <UserCircleIcon className="h-8 w-8 text-gray-400" />
                <div className="hidden lg:block">
                  <p className="text-sm font-medium text-gray-900">
                    {session?.user?.name}
                  </p>
                  <p className="text-xs text-gray-500">
                    {session?.user?.role === 'admin' ? '管理员' : 
                     session?.user?.role === 'manager' ? '管理者' : '查看者'}
                  </p>
                </div>
                <button
                  onClick={handleSignOut}
                  className="ml-2 text-gray-400 hover:text-gray-600 lg:hidden"
                  title="退出登录"
                >
                  <ArrowRightOnRectangleIcon className="h-5 w-5" />
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Main content */}
        <main className="py-10">
          <div className="px-4 sm:px-6 lg:px-8">
            {children}
          </div>
        </main>
      </div>
    </div>
  )
}