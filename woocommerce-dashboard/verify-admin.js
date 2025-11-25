const { createClient } = require('@supabase/supabase-js');
require('dotenv').config({ path: '.env.local' });

async function verifyAdmin() {
  console.log('🔍 验证管理员账号...');
  
  const supabase = createClient(
    process.env.SUPABASE_URL,
    process.env.SUPABASE_SERVICE_ROLE_KEY,
    {
      auth: {
        autoRefreshToken: false,
        persistSession: false
      }
    }
  );

  try {
    // 获取所有用户
    const { data: users, error } = await supabase.auth.admin.listUsers();
    
    if (error) {
      throw error;
    }

    console.log('📊 用户列表:');
    users.users.forEach((user, index) => {
      console.log(`${index + 1}. 邮箱: ${user.email}`);
      console.log(`   用户ID: ${user.id}`);
      console.log(`   角色: ${user.user_metadata?.role || '未设置'}`);
      console.log(`   姓名: ${user.user_metadata?.name || '未设置'}`);
      console.log(`   邮箱已确认: ${user.email_confirmed_at ? '是' : '否'}`);
      console.log(`   创建时间: ${user.created_at}`);
      console.log('---');
    });

    // 查找管理员用户
    const adminUser = users.users.find(u => u.email === 'kxmyt.wf@hotmail.com');
    
    if (adminUser) {
      console.log('✅ 管理员账号验证成功!');
      console.log('📧 管理员邮箱:', adminUser.email);
      console.log('👤 用户ID:', adminUser.id);
      console.log('🔑 角色:', adminUser.user_metadata?.role);
      console.log('📅 创建时间:', adminUser.created_at);
      console.log('✉️  邮箱已确认:', adminUser.email_confirmed_at ? '是' : '否');
    } else {
      console.log('❌ 未找到管理员账号');
    }

  } catch (error) {
    console.error('❌ 验证过程中出错:', error.message);
  }
}

verifyAdmin();