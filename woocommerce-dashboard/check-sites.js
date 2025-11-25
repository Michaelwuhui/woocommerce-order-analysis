const { createClient } = require('@supabase/supabase-js');
require('dotenv').config({ path: '.env.local' });

async function checkSites() {
  console.log('🔍 检查站点数据...');
  
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
    // 获取所有站点
    const { data: sites, error: sitesError } = await supabase
      .from('sites')
      .select('*')
      .order('created_at', { ascending: false });
    
    if (sitesError) {
      throw sitesError;
    }

    console.log('📊 站点数据:');
    console.log(`总共找到 ${sites.length} 个站点`);
    
    if (sites.length > 0) {
      sites.forEach((site, index) => {
        console.log(`${index + 1}. 站点名称: ${site.name}`);
        console.log(`   站点ID: ${site.id}`);
        console.log(`   用户ID: ${site.user_id}`);
        console.log(`   URL: ${site.url}`);
        console.log(`   WooCommerce URL: ${site.woo_url}`);
        console.log(`   状态: ${site.status}`);
        console.log(`   创建时间: ${site.created_at}`);
        console.log('---');
      });
    } else {
      console.log('❌ 数据库中没有找到任何站点数据');
    }

    // 获取所有用户
    const { data: users, error: usersError } = await supabase.auth.admin.listUsers();
    
    if (usersError) {
      throw usersError;
    }

    console.log('\n👥 用户数据:');
    users.users.forEach((user, index) => {
      console.log(`${index + 1}. 邮箱: ${user.email}`);
      console.log(`   用户ID: ${user.id}`);
      console.log(`   角色: ${user.user_metadata?.role || '未设置'}`);
      console.log('---');
    });

    // 检查特定用户的站点
    const adminUser = users.users.find(u => u.email === 'kxmyt.wf@hotmail.com');
    if (adminUser) {
      console.log(`\n🔍 检查管理员用户 (${adminUser.id}) 的站点:`);
      const { data: userSites, error: userSitesError } = await supabase
        .from('sites')
        .select('*')
        .eq('user_id', adminUser.id);
      
      if (userSitesError) {
        throw userSitesError;
      }
      
      console.log(`管理员用户有 ${userSites.length} 个站点`);
      userSites.forEach((site, index) => {
        console.log(`${index + 1}. ${site.name} (${site.status})`);
      });
    }

  } catch (error) {
    console.error('❌ 检查过程中出错:', error.message);
  }
}

checkSites();