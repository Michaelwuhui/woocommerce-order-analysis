const { createClient } = require('@supabase/supabase-js');

const supabaseUrl = 'https://scoevexkuvtyhzpdardd.supabase.co';
const supabaseServiceKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNjb2V2ZXhrdXZ0eWh6cGRhcmRkIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc2MDU3NDgzOCwiZXhwIjoyMDc2MTUwODM4fQ._5Hz7NPJ89iOQSgHT2musLt_2GG0CuCrBw_d2z6wyR8';

const supabase = createClient(supabaseUrl, supabaseServiceKey);

async function testSync() {
  console.log('Testing data synchronization...');
  
  try {
    const { data: site, error: siteError } = await supabase
      .from('sites')
      .select('*')
      .eq('name', 'strefa')
      .single();

    if (siteError) {
      console.error('Error getting site:', siteError);
      return;
    }

    console.log('Site found:', site.name, site.id);

    const WooCommerceRestApi = require('@woocommerce/woocommerce-rest-api').default;
    
    const WooCommerce = new WooCommerceRestApi({
      url: site.woo_url,
      consumerKey: site.consumer_key,
      consumerSecret: site.consumer_secret,
      version: 'wc/v3'
    });

    console.log('Fetching orders from WooCommerce...');
    const response = await WooCommerce.get('orders', {
      per_page: 5,
      status: 'any'
    });

    const orders = response.data;
    console.log(`Found ${orders.length} orders to sync`);

    const { data: syncLog, error: syncLogError } = await supabase
      .from('sync_logs')
      .insert({
        site_id: site.id,
        status: 'running',
        started_at: new Date().toISOString()
      })
      .select()
      .single();

    if (syncLogError) {
      console.error('Error creating sync log:', syncLogError);
      return;
    }

    console.log('Sync log created:', syncLog.id);

    let syncedCount = 0;

    for (const order of orders) {
      try {
        const orderData = {
          order_id: order.id,
          site_id: site.id,
          status: order.status,
          total: parseFloat(order.total),
          currency: order.currency,
          date_created: order.date_created,
          date_modified: order.date_modified,
          customer_id: order.customer_id || null,
          customer_email: order.billing?.email || '',
          customer_name: `${order.billing?.first_name || ''} ${order.billing?.last_name || ''}`.trim(),
          billing_country: order.billing?.country || '',
          billing_state: order.billing?.state || '',
          billing_city: order.billing?.city || '',
          billing_address_1: order.billing?.address_1 || '',
          billing_address_2: order.billing?.address_2 || '',
          billing_postcode: order.billing?.postcode || '',
          billing_phone: order.billing?.phone || '',
          shipping_country: order.shipping?.country || '',
          shipping_state: order.shipping?.state || '',
          shipping_city: order.shipping?.city || '',
          shipping_address_1: order.shipping?.address_1 || '',
          shipping_address_2: order.shipping?.address_2 || '',
          shipping_postcode: order.shipping?.postcode || '',
          payment_method: order.payment_method || '',
          payment_method_title: order.payment_method_title || '',
          transaction_id: order.transaction_id || '',
          line_items: order.line_items || [],
          shipping_lines: order.shipping_lines || [],
          tax_lines: order.tax_lines || [],
          fee_lines: order.fee_lines || [],
          coupon_lines: order.coupon_lines || [],
          meta_data: order.meta_data || []
        };

        const { error: orderError } = await supabase
          .from('orders')
          .upsert(orderData, {
            onConflict: 'site_id,order_id'
          });

        if (orderError) {
          console.error(`Error saving order ${order.id}:`, orderError);
        } else {
          syncedCount++;
          console.log(`✓ Synced order ${order.id} (${order.status}, ${order.total} ${order.currency})`);
        }
      } catch (orderError) {
        console.error(`Error processing order ${order.id}:`, orderError);
      }
    }

    const { error: updateError } = await supabase
      .from('sync_logs')
      .update({
        status: 'completed',
        completed_at: new Date().toISOString(),
        orders_synced: syncedCount,
        message: `Successfully synced ${syncedCount} orders`
      })
      .eq('id', syncLog.id);

    if (updateError) {
      console.error('Error updating sync log:', updateError);
    }

    const { error: siteUpdateError } = await supabase
      .from('sites')
      .update({ last_sync: new Date().toISOString() })
      .eq('id', site.id);

    if (siteUpdateError) {
      console.error('Error updating site last_sync:', siteUpdateError);
    }

    console.log(`\nSync completed successfully!`);
    console.log(`- Orders synced: ${syncedCount}/${orders.length}`);
    console.log(`- Sync log ID: ${syncLog.id}`);

  } catch (error) {
    console.error('Sync failed:', error);
  }
}