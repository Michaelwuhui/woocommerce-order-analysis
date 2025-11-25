const { createClient } = require('@supabase/supabase-js')
require('dotenv').config({ path: '.env.local' })

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL
const supabaseServiceKey = process.env.SUPABASE_SERVICE_ROLE_KEY

const supabase = createClient(supabaseUrl, supabaseServiceKey)

async function testSitesService() {
  try {
    console.log('测试sitesService.getAll方法...')
    
    const userId = '9b5b3552-9f5d-445e-89be-7d2a5eb1b21d'
    console.log('查询用户ID:', userId)
    
    const { data, error } = await supabase
      .from('sites')
      .select('*')
      .eq('user_id', userId)
      .order('created_at', { ascending: false })

    if (error) {
      console.error('查询错误:', error)
      return
    }

    console.log('查询结果:')
    console.log('数据数量:', data.length)
    console.log('数据内容:', JSON.stringify(data, null, 2))
    
    // 同时查询所有站点进行对比
    const { data: allSites, error: allError } = await supabase
      .from('sites')
      .select('*')

    if (allError) {
      console.error('查询所有站点错误:', allError)
      return
    }

    console.log('\n所有站点数据:')
    console.log('总数量:', allSites.length)
    allSites.forEach((site, index) => {
      console.log(`${index + 1}. ID: ${site.id}, 用户ID: ${site.user_id}, 名称: ${site.name}`)
      console.log(`   用户ID匹配: ${site.user_id === userId}`)
    })
    
  } catch (error) {
    console.error('测试错误:', error)
  }
}

testSitesService()