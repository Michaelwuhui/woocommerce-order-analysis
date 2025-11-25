const { createClient } = require('@supabase/supabase-js');
require('dotenv').config({ path: '.env.local' });

async function fixSiteUser() {
  console.log('🔧 修复站点用户ID...');
  
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
    // 获取管理员用户
    const { data: authUsers, error: usersError } = await supabase.auth.admin.listUsers();
    
    if (usersError) {
      throw usersError;
    }

    const adminUser = authUsers.users.find(u => u.email === 'kxmyt.wf@hotmail.com');
    if (!adminUser) {
      console.error('❌ 未找到管理员用户');
      return;
    }

    console.log('👤 管理员用户ID:', adminUser.id);

    // 检查users表中是否有对应记录
    const { data: dbUsers, error: dbUsersError } = await supabase
      .from('users')
      .select('*');
    
    if (dbUsersError) {
      throw dbUsersError;
    }

    console.log('📊 users表中的用户:');
    dbUsers.forEach((user, index) => {
      console.log(`${index + 1}. ID: ${user.id}, Email: ${user.email}, Role: ${user.role}`);
    });

    // 检查管理员用户是否在users表中
    const adminInDb = dbUsers.find(u => u.id === adminUser.id);
    
    if (!adminInDb) {
      console.log('🔄 在users表中创建管理员用户记录...');
      const { error: insertError } = await supabase
        .from('users')
        .insert({
          id: adminUser.id,
          email: adminUser.email,
          name: adminUser.user_metadata?.name || '系统管理员',
          role: 'admin'
        });

      if (insertError) {
        console.error('❌ 创建用户记录失败:', insertError.message);
        return;
      }
      console.log('✅ 管理员用户记录创建成功');
    } else {
      console.log('✅ 管理员用户记录已存在');
    }

    // 获取所有站点
    const { data: sites, error: sitesError } = await supabase
      .from('sites')
      .select('*');
    
    if (sitesError) {
      throw sitesError;
    }

    console.log(`📊 找到 ${sites.length} 个站点`);

    // 更新所有站点的user_id为管理员用户ID
    for (const site of sites) {
      if (site.user_id !== adminUser.id) {
        console.log(`🔄 更新站点 "${site.name}" 的用户ID...`);
        console.log(`   从: ${site.user_id}`);
        console.log(`   到: ${adminUser.id}`);

        const { error: updateError } = await supabase
          .from('sites')
          .update({ user_id: adminUser.id })
          .eq('id', site.id);

        if (updateError) {
          console.error(`❌ 更新站点 ${site.name} 失败:`, updateError.message);
        } else {
          console.log(`✅ 站点 "${site.name}" 更新成功`);
        }
      } else {
        console.log(`✅ 站点 "${site.name}" 用户ID已正确`);
      }
    }

    // 验证修复结果
    console.log('\n🔍 验证修复结果...');
    const { data: updatedSites, error: verifyError } = await supabase
      .from('sites')
      .select('*')
      .eq('user_id', adminUser.id);
    
    if (verifyError) {
      throw verifyError;
    }

    console.log(`✅ 管理员用户现在有 ${updatedSites.length} 个站点:`);
    updatedSites.forEach((site, index) => {
      console.log(`${index + 1}. ${site.name} (${site.status})`);
    });

  } catch (error) {
    console.error('❌ 修复过程中出错:', error.message);
  }
}

fixSiteUser();