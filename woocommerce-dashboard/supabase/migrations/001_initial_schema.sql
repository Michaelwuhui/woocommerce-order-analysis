-- Create users table
CREATE TABLE IF NOT EXISTS users (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  email VARCHAR(255) UNIQUE NOT NULL,
  name VARCHAR(255),
  role VARCHAR(20) DEFAULT 'user' CHECK (role IN ('admin', 'user')),
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create sites table
CREATE TABLE IF NOT EXISTS sites (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name VARCHAR(255) NOT NULL,
  url VARCHAR(500) NOT NULL,
  woo_url VARCHAR(500) NOT NULL,
  consumer_key VARCHAR(255) NOT NULL,
  consumer_secret VARCHAR(255) NOT NULL,
  status VARCHAR(20) DEFAULT 'active' CHECK (status IN ('active', 'inactive')),
  last_sync TIMESTAMP WITH TIME ZONE,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE
);

-- Create orders table
CREATE TABLE IF NOT EXISTS orders (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  site_id UUID NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
  order_id INTEGER NOT NULL,
  status VARCHAR(50) NOT NULL,
  currency VARCHAR(10) NOT NULL,
  total DECIMAL(10, 2) NOT NULL DEFAULT 0,
  date_created TIMESTAMP WITH TIME ZONE NOT NULL,
  date_modified TIMESTAMP WITH TIME ZONE,
  customer_id INTEGER,
  customer_email VARCHAR(255),
  customer_name VARCHAR(255),
  billing_country VARCHAR(10),
  billing_state VARCHAR(100),
  billing_city VARCHAR(100),
  billing_address_1 TEXT,
  billing_address_2 TEXT,
  billing_postcode VARCHAR(20),
  billing_phone VARCHAR(50),
  shipping_country VARCHAR(10),
  shipping_state VARCHAR(100),
  shipping_city VARCHAR(100),
  shipping_address_1 TEXT,
  shipping_address_2 TEXT,
  shipping_postcode VARCHAR(20),
  payment_method VARCHAR(100),
  payment_method_title VARCHAR(255),
  transaction_id VARCHAR(255),
  line_items JSONB,
  shipping_lines JSONB,
  tax_lines JSONB,
  fee_lines JSONB,
  coupon_lines JSONB,
  meta_data JSONB,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  UNIQUE(site_id, order_id)
);

-- Create sync_logs table
CREATE TABLE IF NOT EXISTS sync_logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  site_id UUID NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
  status VARCHAR(20) NOT NULL CHECK (status IN ('running', 'completed', 'failed')),
  started_at TIMESTAMP WITH TIME ZONE NOT NULL,
  completed_at TIMESTAMP WITH TIME ZONE,
  orders_synced INTEGER DEFAULT 0,
  error_message TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create indexes for better performance
CREATE INDEX IF NOT EXISTS idx_sites_user_id ON sites(user_id);
CREATE INDEX IF NOT EXISTS idx_orders_site_id ON orders(site_id);
CREATE INDEX IF NOT EXISTS idx_orders_date_created ON orders(date_created);
CREATE INDEX IF NOT EXISTS idx_orders_status ON orders(status);
CREATE INDEX IF NOT EXISTS idx_orders_customer_email ON orders(customer_email);
CREATE INDEX IF NOT EXISTS idx_sync_logs_site_id ON sync_logs(site_id);
CREATE INDEX IF NOT EXISTS idx_sync_logs_status ON sync_logs(status);

-- Create updated_at trigger function
CREATE OR REPLACE FUNCTION update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ language 'plpgsql';

-- Create triggers for updated_at
CREATE TRIGGER update_users_updated_at BEFORE UPDATE ON users FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();
CREATE TRIGGER update_sites_updated_at BEFORE UPDATE ON sites FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();
CREATE TRIGGER update_orders_updated_at BEFORE UPDATE ON orders FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();
CREATE TRIGGER update_sync_logs_updated_at BEFORE UPDATE ON sync_logs FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- Enable Row Level Security (RLS)
ALTER TABLE users ENABLE ROW LEVEL SECURITY;
ALTER TABLE sites ENABLE ROW LEVEL SECURITY;
ALTER TABLE orders ENABLE ROW LEVEL SECURITY;
ALTER TABLE sync_logs ENABLE ROW LEVEL SECURITY;

-- Create RLS policies
-- Users can only see their own data
CREATE POLICY "Users can view own profile" ON users FOR SELECT USING (auth.uid()::text = id::text);
CREATE POLICY "Users can update own profile" ON users FOR UPDATE USING (auth.uid()::text = id::text);

-- Sites policies
CREATE POLICY "Users can view own sites" ON sites FOR SELECT USING (auth.uid()::text = user_id::text);
CREATE POLICY "Users can insert own sites" ON sites FOR INSERT WITH CHECK (auth.uid()::text = user_id::text);
CREATE POLICY "Users can update own sites" ON sites FOR UPDATE USING (auth.uid()::text = user_id::text);
CREATE POLICY "Users can delete own sites" ON sites FOR DELETE USING (auth.uid()::text = user_id::text);

-- Orders policies
CREATE POLICY "Users can view orders from own sites" ON orders FOR SELECT USING (
  EXISTS (
    SELECT 1 FROM sites 
    WHERE sites.id = orders.site_id 
    AND sites.user_id::text = auth.uid()::text
  )
);
CREATE POLICY "Users can insert orders to own sites" ON orders FOR INSERT WITH CHECK (
  EXISTS (
    SELECT 1 FROM sites 
    WHERE sites.id = orders.site_id 
    AND sites.user_id::text = auth.uid()::text
  )
);
CREATE POLICY "Users can update orders from own sites" ON orders FOR UPDATE USING (
  EXISTS (
    SELECT 1 FROM sites 
    WHERE sites.id = orders.site_id 
    AND sites.user_id::text = auth.uid()::text
  )
);

-- Sync logs policies
CREATE POLICY "Users can view sync logs from own sites" ON sync_logs FOR SELECT USING (
  EXISTS (
    SELECT 1 FROM sites 
    WHERE sites.id = sync_logs.site_id 
    AND sites.user_id::text = auth.uid()::text
  )
);
CREATE POLICY "Users can insert sync logs to own sites" ON sync_logs FOR INSERT WITH CHECK (
  EXISTS (
    SELECT 1 FROM sites 
    WHERE sites.id = sync_logs.site_id 
    AND sites.user_id::text = auth.uid()::text
  )
);
CREATE POLICY "Users can update sync logs from own sites" ON sync_logs FOR UPDATE USING (
  EXISTS (
    SELECT 1 FROM sites 
    WHERE sites.id = sync_logs.site_id 
    AND sites.user_id::text = auth.uid()::text
  )
);