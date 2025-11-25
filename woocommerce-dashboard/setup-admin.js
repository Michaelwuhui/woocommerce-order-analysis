const { createClient } = require('@supabase/supabase-js');
require('dotenv').config({ path: '.env.local' });

// 管理员账号信息
const ADMIN_EMAIL = 'kxmyt.wf@hotmail.com';
const ADMIN_PASSWORD = 'kxmyt090415';

async function setupAdmin() {
  console.log('🚀 开始设置管理员账号...');
  
  // 检查环境变量
  if (!process.env.SUPABASE_URL || !process.env.SUPABASE_SERVICE_ROLE_KEY) {
    console.error('❌ 错误: 缺少Supabase配置信息');
    console.error('请确保.env.local文件中包含SUPABASE_URL和SUPABASE_SERVICE_ROLE_KEY');
    process.exit(1);
  }

  // 创建Supabase客户端（使用service role key以便创建用户）
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
    console.log('📧 创建管理员用户:', ADMIN_EMAIL);
    
    // 创建用户
    const { data: user, error: createError } = await supabase.auth.admin.createUser({
      email: ADMIN_EMAIL,
      password: ADMIN_PASSWORD,
      email_confirm: true, // 自动确认邮箱
      user_metadata: {
        role: 'admin',
        name: '系统管理员',
        created_by: 'setup-script'
      }
    });

    if (createError) {
      if (createError.message.includes('already registered')) {
        console.log('⚠️  用户已存在，尝试更新用户信息...');
        
        // 获取现有用户
        const { data: existingUsers, error: listError } = await supabase.auth.admin.listUsers();
        if (listError) {
          throw listError;
        }
        
        const existingUser = existingUsers.users.find(u => u.email === ADMIN_EMAIL);
        if (existingUser) {
          // 更新用户密码和元数据
          const { data: updatedUser, error: updateError } = await supabase.auth.admin.updateUserById(
            existingUser.id,
            {
              password: ADMIN_PASSWORD,
              user_metadata: {
                role: 'admin',
                name: '系统管理员',
                updated_by: 'setup-script',
                updated_at: new Date().toISOString()
              }
            }
          );
          
          if (updateError) {
            throw updateError;
          }
          
          console.log('✅ 用户信息已更新');
          console.log('👤 用户ID:', existingUser.id);
          console.log('📧 邮箱:', existingUser.email);
        }
      } else {
        throw createError;
      }
    } else {
      console.log('✅ 管理员用户创建成功!');
      console.log('👤 用户ID:', user.user.id);
      console.log('📧 邮箱:', user.user.email);
    }

    console.log('\n🎉 管理员账号设置完成!');
    console.log('📋 登录信息:');
    console.log('   邮箱:', ADMIN_EMAIL);
    console.log('   密码:', ADMIN_PASSWORD);
    console.log('   角色: 管理员');
    console.log('\n🌐 您现在可以使用这些凭据登录系统了!');

  } catch (error) {
    console.error('❌ 设置管理员账号时出错:', error.message);
    console.error('详细错误信息:', error);
    process.exit(1);
  }
}

// 运行设置脚本
if (require.main === module) {
  setupAdmin().catch(console.error);
}

module.exports = { setupAdmin };