const { default: fetch } = require('node-fetch')

async function testAPIDebug() {
  try {
    console.log('1. 先登录获取session...')
    
    const loginResponse = await fetch('http://localhost:3000/api/auth/simple-login', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        email: 'kxmyt.wf@hotmail.com',
        password: 'test123456'
      })
    })
    
    const loginData = await loginResponse.json()
    console.log('登录响应:', loginData)
    
    if (!loginResponse.ok) {
      console.log('登录失败')
      return
    }
    
    const cookies = loginResponse.headers.get('set-cookie')
    console.log('获取到的Cookie:', cookies)
    
    console.log('\n2. 使用cookie调用sites API...')
    
    const sitesResponse = await fetch('http://localhost:3000/api/sites', {
      headers: {
        'Cookie': cookies || ''
      }
    })
    
    console.log('Sites API响应状态:', sitesResponse.status)
    console.log('Sites API响应头:', Object.fromEntries(sitesResponse.headers.entries()))
    
    const sitesText = await sitesResponse.text()
    console.log('Sites API原始响应:', sitesText)
    
    try {
      const sitesData = JSON.parse(sitesText)
      console.log('Sites API解析后数据:', sitesData)
    } catch (e) {
      console.log('无法解析JSON响应')
    }
    
  } catch (error) {
    console.error('测试错误:', error)
  }
}

testAPIDebug()