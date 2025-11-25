import type { Metadata } from 'next'
import { Inter } from 'next/font/google'
import './globals.css'
import { Providers } from '@/components/providers'

const inter = Inter({ subsets: ['latin'] })

export const metadata: Metadata = {
  title: 'WooCommerce多站点数据分析平台',
  description: '统一管理和分析多个WooCommerce站点的订单数据、销售趋势和业务指标',
  keywords: ['WooCommerce', '数据分析', '电商', '订单管理', '销售统计'],
  authors: [{ name: 'WooCommerce Analytics Team' }],
  viewport: 'width=device-width, initial-scale=1',
  robots: 'noindex, nofollow', // 防止搜索引擎索引内部管理系统
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="zh-CN" className="h-full">
      <body className={`${inter.className} h-full antialiased`}>
        <Providers>
          {children}
        </Providers>
      </body>
    </html>
  )
}