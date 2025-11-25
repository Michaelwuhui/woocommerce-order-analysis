const { createClient } = require('@supabase/supabase-js');
require('dotenv').config({ path: '.env.local' });

async function testAPI() {
  console.log('🔍 测试API和认证...');
  
  try {
    // 测试NextAuth session API
    console.log('\n1. 测试NextAuth session API...');
    const sessionResponse = await fetch('http://localhost:3000/api/auth/session');
    console.log('Session API状态:', sessionResponse.status);
    
    if (sessionResponse.ok) {
      const sessionData = await sessionResponse.json();
      console.log('Session数据:', sessionData);
    } else {
      console.log('Session API错误:', sessionResponse.statusText);
    }

    // 测试sites API (无认证)
    console.log('\n2. 测试sites API (无认证)...');
    const sitesResponse = await fetch('http://localhost:3000/api/sites');
    console.log('Sites API状态:', sitesResponse.status);
    
    if (sitesResponse.ok) {
      const sitesData = await sitesResponse.json();
      console.log('Sites数据:', sitesData);
    } else {
      console.log('Sites API错误:', sitesResponse.statusText);
      const errorText = await sitesResponse.text();
      console.log('错误详情:', errorText);
    }

    // 检查Supabase连接
    console.log('\n3. 测试Supabase连接...');
    const supabase = createClient(
      process.env.SUPABASE_URL,
      process.env.SUPABASE_ANON_KEY
    );

    const { data, error } = await supabase.from('sites').select('count');
    if (error) {
      console.log('Supabase错误:', error.message);
    } else {
      console.log('Supabase连接正常');
    }

    // 检查环境变量
    console.log('\n4. 检查关键环境变量...');
    console.log('NEXTAUTH_URL:', process.env.NEXTAUTH_URL);
    console.log('NEXTAUTH_SECRET:', process.env.NEXTAUTH_SECRET ? '已设置' : '未设置');
    console.log('SUPABASE_URL:', process.env.SUPABASE_URL ? '已设置' : '未设置');
    console.log('SUPABASE_SERVICE_ROLE_KEY:', process.env.SUPABASE_SERVICE_ROLE_KEY ? '已设置' : '未设置');

  } catch (error) {
    console.error('❌ 测试过程中出错:', error.message);
  }
}

testAPI();