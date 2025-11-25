const { createClient } = require('@supabase/supabase-js')
require('dotenv').config({ path: '.env.local' })

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL
const supabaseServiceKey = process.env.SUPABASE_SERVICE_ROLE_KEY

const supabase = createClient(supabaseUrl, supabaseServiceKey)

async function checkSitesData() {
  try {
    console.log('检查数据库中的站点数据...')
    
    // 检查sites表中的所有数据
    const { data: sites, error: sitesError } = await supabase
      .from('sites')
      .select('*')
    
    if (sitesError) {
      console.error('查询sites表错误:', sitesError)
      return
    }
    
    console.log('Sites表中的数据:')
    console.log('总数:', sites.length)
    sites.forEach((site, index) => {
      console.log(`${index + 1}. ID: ${site.id}, 名称: ${site.name}, URL: ${site.url}, 用户ID: ${site.user_id}`)
    })
    
    // 检查用户表
    const { data: users, error: usersError } = await supabase
      .from('users')
      .select('*')
    
    if (usersError) {
      console.error('查询users表错误:', usersError)
      return
    }
    
    console.log('\nUsers表中的数据:')
    console.log('总数:', users.length)
    users.forEach((user, index) => {
      console.log(`${index + 1}. ID: ${user.id}, 邮箱: ${user.email}, 角色: ${user.role}`)
    })
    
    // 检查特定用户的站点
    const testUserEmail = 'kxmyt.wf@hotmail.com'
    const testUser = users.find(u => u.email === testUserEmail)
    
    if (testUser) {
      console.log(`\n测试用户 ${testUserEmail} 的站点:`)
      const userSites = sites.filter(s => s.user_id === testUser.id)
      console.log('用户站点数量:', userSites.length)
      userSites.forEach((site, index) => {
        console.log(`${index + 1}. ${site.name} - ${site.url}`)
      })
    } else {
      console.log(`\n未找到测试用户: ${testUserEmail}`)
    }
    
  } catch (error) {
    console.error('检查数据时发生错误:', error)
  }
}

checkSitesData()