'use client'

import { useSearchParams } from 'next/navigation'
import Link from 'next/link'
import { ExclamationTriangleIcon } from '@heroicons/react/24/outline'

const errorMessages = {
  Configuration: '服务器配置错误，请联系管理员',
  AccessDenied: '访问被拒绝，您没有权限访问此资源',
  Verification: '验证失败，请重试',
  Default: '登录过程中发生未知错误',
}

export default function AuthErrorPage() {
  const searchParams = useSearchParams()
  const error = searchParams.get('error') as keyof typeof errorMessages

  const errorMessage = errorMessages[error] || errorMessages.Default

  return (
    <div className="min-h-screen flex items-center justify-center bg-gray-50 py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-md w-full space-y-8">
        <div className="text-center">
          <div className="mx-auto h-12 w-12 flex items-center justify-center rounded-full bg-red-100">
            <ExclamationTriangleIcon className="h-8 w-8 text-red-600" />
          </div>
          <h2 className="mt-6 text-center text-3xl font-extrabold text-gray-900">
            登录失败
          </h2>
          <p className="mt-2 text-center text-sm text-gray-600">
            {errorMessage}
          </p>
        </div>

        <div className="mt-8 space-y-4">
          <Link
            href="/auth/signin"
            className="w-full flex justify-center py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-primary-600 hover:bg-primary-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-primary-500"
          >
            重新登录
          </Link>
          
          <div className="text-center">
            <p className="text-sm text-gray-600">
              如果问题持续存在，请{' '}
              <a href="mailto:admin@example.com" className="font-medium text-primary-600 hover:text-primary-500">
                联系管理员
              </a>
            </p>
          </div>
        </div>
      </div>
    </div>
  )
}