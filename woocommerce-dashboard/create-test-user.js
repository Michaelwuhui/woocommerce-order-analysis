const { createClient } = require('@supabase/supabase-js');
require('dotenv').config({ path: '.env.local' });

async function createTestUser() {
  console.log('🔍 创建测试用户...');
  
  try {
    const supabase = createClient(
      process.env.SUPABASE_URL,
      process.env.SUPABASE_SERVICE_ROLE_KEY
    );

    // 创建测试用户
    const { data, error } = await supabase.auth.admin.createUser({
      email: 'kxmyt.wf@hotmail.com',
      password: 'test123456',
      email_confirm: true,
      user_metadata: {
        name: 'Administrator',
        role: 'admin'
      }
    });

    if (error) {
      console.log('❌ 创建用户失败:', error.message);
      
      // 如果用户已存在，尝试获取用户信息
      if (error.message.includes('already registered')) {
        console.log('✅ 用户已存在，尝试获取用户信息...');
        const { data: users, error: listError } = await supabase.auth.admin.listUsers();
        if (!listError) {
          const existingUser = users.users.find(u => u.email === 'kxmyt.wf@hotmail.com');
          if (existingUser) {
            console.log('✅ 找到现有用户:', {
              id: existingUser.id,
              email: existingUser.email,
              created_at: existingUser.created_at
            });
          }
        }
      }
    } else {
      console.log('✅ 用户创建成功:', {
        id: data.user.id,
        email: data.user.email,
        created_at: data.user.created_at
      });
    }

    // 检查public.users表中的用户
    console.log('\n🔍 检查public.users表...');
    const { data: publicUsers, error: publicError } = await supabase
      .from('users')
      .select('id, email, name, role, created_at')
      .eq('email', 'kxmyt.wf@hotmail.com');

    if (publicError) {
      console.log('❌ 查询public.users失败:', publicError.message);
    } else {
      console.log('✅ public.users中的用户:', publicUsers);
    }

  } catch (error) {
    console.error('❌ 过程中出错:', error.message);
  }
}

createTestUser();