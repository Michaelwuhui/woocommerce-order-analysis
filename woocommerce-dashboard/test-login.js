const { default: fetch } = require('node-fetch');
require('dotenv').config({ path: '.env.local' });

async function testLogin() {
  console.log('🔍 测试登录功能...');
  
  try {
    // 首先获取CSRF token
    console.log('\n1. 获取CSRF token...');
    const csrfResponse = await fetch('http://localhost:3000/api/auth/csrf');
    const csrfData = await csrfResponse.json();
    console.log('CSRF token:', csrfData.csrfToken);

    // 尝试登录
    console.log('\n2. 尝试登录...');
    const loginResponse = await fetch('http://localhost:3000/api/auth/callback/credentials', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        email: 'kxmyt.wf@hotmail.com',
        password: 'test123456',
        csrfToken: csrfData.csrfToken,
        callbackUrl: 'http://localhost:3000/dashboard',
        json: 'true'
      })
    });

    console.log('登录响应状态:', loginResponse.status);
    console.log('登录响应头:', Object.fromEntries(loginResponse.headers.entries()));
    
    const loginResult = await loginResponse.text();
    console.log('登录响应内容:', loginResult);

  } catch (error) {
    console.error('❌ 测试过程中出错:', error.message);
  }
}

testLogin();