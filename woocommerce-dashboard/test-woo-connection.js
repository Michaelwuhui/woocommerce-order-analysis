const WooCommerceRestApi = require('woocommerce-rest-api').default;

async function testWooCommerceConnection() {
  console.log('Testing WooCommerce connection...');
  
  try {
    const WooCommerce = new WooCommerceRestApi({
      url: 'https://www.strefajednorazowek.pl/wp-json/wc/v3',
      consumerKey: 'ck_9e94bbcfb4f7b6cb06eb7c19b0c7eaa3a70c2221',
      consumerSecret: 'cs_64b1e2dbaf1436dd822616fd306769ce3907bbfb',
      version: 'wc/v3'
    });

    console.log('Fetching system status...');
    const response = await WooCommerce.get('system_status');
    
    console.log('WooCommerce connection successful!');
    console.log('Store Info:');
    console.log('- Name:', response.data.settings?.general?.woocommerce_store_name || 'Unknown');
    console.log('- Version:', response.data.environment?.version || 'Unknown');
    console.log('- Currency:', response.data.settings?.general?.woocommerce_currency || 'Unknown');
    
    console.log('Testing orders fetch...');
    const ordersResponse = await WooCommerce.get('orders', {
      per_page: 5,
      status: 'any'
    });
    
    console.log(`Successfully fetched ${ordersResponse.data.length} orders`);
    if (ordersResponse.data.length > 0) {
      const firstOrder = ordersResponse.data[0];
      console.log('Sample order:');
      console.log('- ID:', firstOrder.id);
      console.log('- Status:', firstOrder.status);
      console.log('- Total:', firstOrder.total);
      console.log('- Date:', firstOrder.date_created);
    }
    
    return true;
    
  } catch (error) {
    console.error('WooCommerce connection failed:');
    console.error('Error:', error.message);
    if (error.response) {
      console.error('Status:', error.response.status);
      console.error('Data:', error.response.data);
    }
    return false;
  }
}

testWooCommerceConnection();