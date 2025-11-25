const { createClient } = require('@supabase/supabase-js');
require('dotenv').config({ path: '.env.local' });

async function resetPassword() {
  console.log('🔍 重置用户密码...');
  
  try {
    const supabase = createClient(
      process.env.SUPABASE_URL,
      process.env.SUPABASE_SERVICE_ROLE_KEY
    );

    // 获取用户ID
    const { data: users, error: listError } = await supabase.auth.admin.listUsers();
    if (listError) {
      console.log('❌ 获取用户列表失败:', listError.message);
      return;
    }

    const user = users.users.find(u => u.email === 'kxmyt.wf@hotmail.com');
    if (!user) {
      console.log('❌ 找不到用户');
      return;
    }

    console.log('✅ 找到用户:', user.id);

    // 更新用户密码
    const { data, error } = await supabase.auth.admin.updateUserById(user.id, {
      password: 'test123456'
    });

    if (error) {
      console.log('❌ 重置密码失败:', error.message);
    } else {
      console.log('✅ 密码重置成功');
    }

    // 测试登录
    console.log('\n🔍 测试登录...');
    const { data: loginData, error: loginError } = await supabase.auth.signInWithPassword({
      email: 'kxmyt.wf@hotmail.com',
      password: 'test123456'
    });

    if (loginError) {
      console.log('❌ 登录测试失败:', loginError.message);
    } else {
      console.log('✅ 登录测试成功:', {
        id: loginData.user.id,
        email: loginData.user.email
      });
    }

  } catch (error) {
    console.error('❌ 过程中出错:', error.message);
  }
}

resetPassword();