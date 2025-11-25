const { default: fetch } = require('node-fetch')

async function testSimpleLogin() {
  try {
    console.log('测试简化登录API...')
    
    const response = await fetch('http://localhost:3000/api/auth/simple-login', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        email: 'kxmyt.wf@hotmail.com',
        password: 'test123456'
      })
    })
    
    console.log('响应状态:', response.status)
    
    const data = await response.text()
    console.log('响应内容:', data)
    
    if (response.ok) {
      console.log('登录成功!')
      
      // 获取cookie
      const cookies = response.headers.get('set-cookie')
      console.log('设置的Cookie:', cookies)
      
      // 测试获取站点数据
      console.log('\n测试获取站点数据...')
      const sitesResponse = await fetch('http://localhost:3000/api/sites', {
        headers: {
          'Cookie': cookies || ''
        }
      })
      
      console.log('站点API响应状态:', sitesResponse.status)
      const sitesData = await sitesResponse.text()
      console.log('站点数据:', sitesData)
      
    } else {
      console.log('登录失败')
    }
    
  } catch (error) {
    console.error('测试错误:', error)
  }
}

testSimpleLogin()