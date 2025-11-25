'use client'

import { usePathname } from 'next/navigation'
import Link from 'next/link'
import { 
  HomeIcon,
  GlobeAltIcon,
  ShoppingCartIcon,
  DocumentChartBarIcon,
  ArrowPathIcon,
  UsersIcon,
  XMarkIcon
} from '@heroicons/react/24/outline'
import { clsx } from 'clsx'

const navigation = [
  { name: '仪表板', href: '/dashboard', icon: HomeIcon },
  { name: '站点管理', href: '/sites', icon: GlobeAltIcon },
  { name: '订单分析', href: '/orders', icon: ShoppingCartIcon },
  { name: '报告中心', href: '/reports', icon: DocumentChartBarIcon },
  { name: '数据同步', href: '/sync', icon: ArrowPathIcon },
  { name: '用户管理', href: '/users', icon: UsersIcon },
]

interface SidebarProps {
  isOpen: boolean
  onClose: () => void
}

export default function Sidebar({ isOpen, onClose }: SidebarProps) {
  const pathname = usePathname()

  return (
    <>
      {/* Mobile sidebar overlay */}
      {isOpen && (
        <div className="fixed inset-0 flex z-40 lg:hidden">
          <div className="fixed inset-0 bg-gray-600 bg-opacity-75" onClick={onClose} />
          <div className="relative flex-1 flex flex-col max-w-xs w-full bg-white">
            <div className="absolute top-0 right-0 -mr-12 pt-2">
              <button
                type="button"
                className="ml-1 flex items-center justify-center h-10 w-10 rounded-full focus:outline-none focus:ring-2 focus:ring-inset focus:ring-white"
                onClick={onClose}
              >
                <span className="sr-only">关闭侧边栏</span>
                <XMarkIcon className="h-6 w-6 text-white" aria-hidden="true" />
              </button>
            </div>
            <SidebarContent pathname={pathname} />
          </div>
        </div>
      )}

      {/* Desktop sidebar */}
      <div className="hidden lg:flex lg:w-64 lg:flex-col lg:fixed lg:inset-y-0">
        <div className="flex-1 flex flex-col min-h-0 bg-white border-r border-gray-200">
          <SidebarContent pathname={pathname} />
        </div>
      </div>
    </>
  )
}

function SidebarContent({ pathname }: { pathname: string }) {
  return (
    <div className="flex-1 flex flex-col pt-5 pb-4 overflow-y-auto">
      <div className="flex items-center flex-shrink-0 px-4">
        <div className="flex items-center">
          <div className="flex-shrink-0">
            <div className="h-8 w-8 bg-blue-600 rounded-lg flex items-center justify-center">
              <span className="text-white font-bold text-sm">WC</span>
            </div>
          </div>
          <div className="ml-3">
            <p className="text-sm font-medium text-gray-900">数据分析平台</p>
          </div>
        </div>
      </div>
      <nav className="mt-8 flex-1 px-2 space-y-1">
        {navigation.map((item) => {
          const isActive = pathname === item.href
          return (
            <Link
              key={item.name}
              href={item.href}
              className={clsx(
                isActive
                  ? 'bg-blue-50 border-blue-500 text-blue-700'
                  : 'border-transparent text-gray-600 hover:bg-gray-50 hover:text-gray-900',
                'group flex items-center px-3 py-2 text-sm font-medium border-l-4 transition-colors duration-200'
              )}
            >
              <item.icon
                className={clsx(
                  isActive ? 'text-blue-500' : 'text-gray-400 group-hover:text-gray-500',
                  'mr-3 flex-shrink-0 h-6 w-6'
                )}
                aria-hidden="true"
              />
              {item.name}
            </Link>
          )
        })}
      </nav>
    </div>
  )
}