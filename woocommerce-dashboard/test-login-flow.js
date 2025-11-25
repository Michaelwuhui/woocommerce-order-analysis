// 使用内置的fetch (Node.js 18+)
const fetch = globalThis.fetch;

async function testLoginFlow() {
  console.log('测试简化登录流程...\n');

  try {
    // 1. 测试登录API
    console.log('1. 测试登录API...');
    const loginResponse = await fetch('http://localhost:3000/api/auth/simple-login', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        email: 'kxmyt.wf@hotmail.com',
        password: 'test123456'
      }),
    });

    console.log('登录响应状态:', loginResponse.status);
    
    if (!loginResponse.ok) {
      const errorData = await loginResponse.json();
      console.log('登录失败:', errorData);
      return;
    }

    const loginData = await loginResponse.json();
    console.log('登录成功:', loginData);

    // 获取cookie
    const setCookieHeader = loginResponse.headers.get('set-cookie');
    console.log('设置的Cookie:', setCookieHeader);

    if (!setCookieHeader) {
      console.log('错误: 没有设置cookie');
      return;
    }

    // 2. 使用cookie测试session API
    console.log('\n2. 使用cookie测试session API...');
    const sessionResponse = await fetch('http://localhost:3000/api/auth/session', {
      method: 'GET',
      headers: {
        'Cookie': setCookieHeader
      }
    });

    console.log('Session API响应状态:', sessionResponse.status);
    
    if (sessionResponse.ok) {
      const sessionData = await sessionResponse.json();
      console.log('Session数据:', sessionData);
    } else {
      const errorData = await sessionResponse.json();
      console.log('Session API错误:', errorData);
    }

    // 3. 使用cookie测试sites API
    console.log('\n3. 使用cookie测试sites API...');
    const sitesResponse = await fetch('http://localhost:3000/api/sites', {
      method: 'GET',
      headers: {
        'Cookie': setCookieHeader
      }
    });

    console.log('Sites API响应状态:', sitesResponse.status);
    
    if (sitesResponse.ok) {
      const sitesData = await sitesResponse.json();
      console.log('Sites数据:', sitesData);
    } else {
      const errorData = await sitesResponse.json();
      console.log('Sites API错误:', errorData);
    }

  } catch (error) {
    console.error('测试过程中发生错误:', error);
  }
}

testLoginFlow();