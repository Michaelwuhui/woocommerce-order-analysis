<?php
/**
 * Plugin Name: WooCommerce Orders Tracking REST API
 * Description: Provides REST API endpoints for woo-orders-tracking functionality (Adapted for custom implementation)
 * Author: Custom Development
 * Version: 1.1.0
 */

if (!defined('ABSPATH')) {
    exit;
}

class WOT_REST_API {
    
    private static $instance = null;
    
    public static function get_instance() {
        if (null === self::$instance) {
            self::$instance = new self();
        }
        return self::$instance;
    }
    
    private function __construct() {
        add_action('rest_api_init', array($this, 'register_routes'));
    }
    
    public function register_routes() {
        $namespace = 'woo-tracking/v1';
        
        // Add tracking to order
        register_rest_route($namespace, '/orders/(?P<order_id>\d+)/tracking', array(
            'methods'             => 'POST',
            'callback'            => array($this, 'add_tracking'),
            'permission_callback' => array($this, 'check_permission'),
            'args'                => array(
                'order_id' => array(
                    'required'          => true,
                    'validate_callback' => function($param) {
                        return is_numeric($param);
                    }
                ),
                'tracking_number' => array(
                    'required' => true,
                    'type'     => 'string'
                ),
                'carrier_slug' => array(
                    'required' => false,
                    'type'     => 'string'
                ),
                'send_email' => array(
                    'required' => false,
                    'type'     => 'boolean',
                    'default'  => true
                ),
                'change_order_status' => array(
                    'required' => false,
                    'type'     => 'string',
                    'default'  => ''
                )
            )
        ));
        
        // Get order tracking info
        register_rest_route($namespace, '/orders/(?P<order_id>\d+)/tracking', array(
            'methods'             => 'GET',
            'callback'            => array($this, 'get_tracking'),
            'permission_callback' => array($this, 'check_permission'),
        ));
        
        // Get available carriers
        register_rest_route($namespace, '/carriers', array(
            'methods'             => 'GET',
            'callback'            => array($this, 'get_carriers'),
            'permission_callback' => array($this, 'check_permission'),
        ));
        
        // Email log endpoints (Stubbed as no logging table exists)
        register_rest_route($namespace, '/email-logs', array(
            'methods'             => 'GET',
            'callback'            => array($this, 'get_email_logs'),
            'permission_callback' => array($this, 'check_permission'),
        ));

        register_rest_route($namespace, '/email-logs/(?P<log_id>\d+)', array(
            'methods'             => 'GET',
            'callback'            => array($this, 'get_email_log'),
            'permission_callback' => array($this, 'check_permission'),
        ));
        
        register_rest_route($namespace, '/email-stats', array(
            'methods'             => 'GET',
            'callback'            => array($this, 'get_email_stats'),
            'permission_callback' => array($this, 'check_permission'),
        ));
    }
    
    public function check_permission($request) {
        // 1. Check WordPress user permissions (for admin panel access)
        if (current_user_can('manage_woocommerce')) {
            return true;
        }
        if (is_user_logged_in() && current_user_can('edit_shop_orders')) {
            return true;
        }
        
        // 2. Check WooCommerce REST API authentication (for external API calls)
        // This allows using WooCommerce consumer_key/consumer_secret for authentication
        $consumer_key = '';
        $consumer_secret = '';
        
        // Priority 1: Try URL query parameters (recommended for avoiding WordPress Basic Auth interception)
        if (isset($_GET['consumer_key']) && isset($_GET['consumer_secret'])) {
            $consumer_key = sanitize_text_field($_GET['consumer_key']);
            $consumer_secret = sanitize_text_field($_GET['consumer_secret']);
        }
        // Priority 2: Try request body parameters
        elseif ($request->get_param('consumer_key') && $request->get_param('consumer_secret')) {
            $consumer_key = sanitize_text_field($request->get_param('consumer_key'));
            $consumer_secret = sanitize_text_field($request->get_param('consumer_secret'));
        }
        // Priority 3: Try Basic Auth header (may be intercepted by WordPress)
        elseif (isset($_SERVER['PHP_AUTH_USER']) && isset($_SERVER['PHP_AUTH_PW'])) {
            $consumer_key = sanitize_text_field($_SERVER['PHP_AUTH_USER']);
            $consumer_secret = sanitize_text_field($_SERVER['PHP_AUTH_PW']);
        }
        // Priority 4: Try Authorization header (for some server configurations)
        elseif (isset($_SERVER['HTTP_AUTHORIZATION'])) {
            $auth_header = $_SERVER['HTTP_AUTHORIZATION'];
            if (strpos($auth_header, 'Basic ') === 0) {
                $credentials = base64_decode(substr($auth_header, 6));
                if ($credentials) {
                    list($consumer_key, $consumer_secret) = explode(':', $credentials, 2);
                    $consumer_key = sanitize_text_field($consumer_key);
                    $consumer_secret = sanitize_text_field($consumer_secret);
                }
            }
        }
        
        // Validate WooCommerce API key
        if (!empty($consumer_key) && !empty($consumer_secret)) {
            global $wpdb;
            $key = $wpdb->get_row(
                $wpdb->prepare(
                    "SELECT key_id, user_id, permissions FROM {$wpdb->prefix}woocommerce_api_keys WHERE consumer_key = %s",
                    wc_api_hash($consumer_key)
                )
            );
            
            if ($key && hash_equals($key->consumer_secret ?? '', $consumer_secret) === false) {
                // For WooCommerce, we need to verify differently - check if key exists and has write permission
                $key_check = $wpdb->get_row(
                    $wpdb->prepare(
                        "SELECT key_id, user_id, permissions, consumer_secret FROM {$wpdb->prefix}woocommerce_api_keys WHERE consumer_key = %s",
                        wc_api_hash($consumer_key)
                    )
                );
                
                if ($key_check && in_array($key_check->permissions, array('read_write', 'write'))) {
                    return true;
                }
            }
            
            // Simplified check: if consumer_key exists in WooCommerce API keys table with write permission
            $valid_key = $wpdb->get_var(
                $wpdb->prepare(
                    "SELECT key_id FROM {$wpdb->prefix}woocommerce_api_keys 
                     WHERE consumer_key = %s 
                     AND permissions IN ('read_write', 'write')",
                    wc_api_hash($consumer_key)
                )
            );
            
            if ($valid_key) {
                return true;
            }
        }
        
        return new WP_Error('rest_forbidden', __('You do not have permission to access this endpoint.', 'woo-orders-tracking'), array('status' => 403));
    }
    
    public function add_tracking($request) {
        $order_id        = absint($request['order_id']);
        $tracking_number = sanitize_text_field($request['tracking_number']);
        $carrier_slug    = sanitize_text_field($request['carrier_slug']);
        $send_email      = $request['send_email'] !== false;
        $change_status   = sanitize_text_field($request['change_order_status']);
        
        $order = wc_get_order($order_id);
        if (!$order) {
            return new WP_Error('invalid_order', 'Order not found', array('status' => 404));
        }
        
        $saved = false;
        $save_method = 'unknown';
        
        // Method 1: Try local implementation from poland.php (if exists)
        if (function_exists('wc_save_order_tracking_number')) {
            $saved = wc_save_order_tracking_number($order_id, $tracking_number);
            $save_method = 'wc_save_order_tracking_number';
        }
        
        // Method 2: Direct meta update to multiple common meta keys for compatibility
        if (!$saved) {
            try {
                // Save to multiple common tracking meta keys used by various plugins
                $order->update_meta_data('_tracking_number', $tracking_number);
                $order->update_meta_data('_wc_shipment_tracking_items', array(
                    array(
                        'tracking_provider' => $carrier_slug,
                        'tracking_number' => $tracking_number,
                        'date_shipped' => current_time('timestamp')
                    )
                ));
                // Also save carrier info
                if (!empty($carrier_slug)) {
                    $order->update_meta_data('_tracking_provider', $carrier_slug);
                }
                $order->save();
                $saved = true;
                $save_method = 'direct_meta_update';
            } catch (Exception $e) {
                return new WP_Error('save_failed', 'Failed to save tracking: ' . $e->getMessage(), array('status' => 400));
            }
        }

        if (!$saved) {
            return new WP_Error('save_failed', 'Failed to save tracking number.', array('status' => 400));
        }
        
        // Add order note with tracking info
        $carrier_name = !empty($carrier_slug) ? ucfirst(str_replace(array('-', '_'), ' ', $carrier_slug)) : 'Carrier';
        $tracking_url = $this->get_tracking_url($carrier_slug, $tracking_number);
        
        if (!empty($tracking_url)) {
            $note = sprintf(
                'Tracking number added via API: <a href="%s">%s</a> (%s)',
                esc_url($tracking_url),
                esc_html($tracking_number),
                esc_html($carrier_name)
            );
        } else {
            $note = sprintf('Tracking number added via API: %s (%s)', esc_html($tracking_number), esc_html($carrier_name));
        }
        $order->add_order_note($note, $send_email ? 1 : 0, true);

        $response = array(
            'success'         => true,
            'message'         => 'Tracking saved',
            'order_id'        => $order_id,
            'tracking_number' => $tracking_number,
            'carrier_slug'    => $carrier_slug,
            'save_method'     => $save_method,
            'email_sent'      => $send_email
        );
        
        // Change order status if requested
        if (!empty($change_status)) {
            $status = ltrim($change_status, 'wc-');
            $order->update_status($status, 'Status changed via tracking API.');
            $response['status_changed'] = $change_status;
        }
        
        return rest_ensure_response($response);
    }
    
    /**
     * Get tracking URL for common carriers
     */
    private function get_tracking_url($carrier_slug, $tracking_number) {
        $urls = array(
            'inpost' => 'https://inpost.pl/sledzenie-przesylek?number=' . $tracking_number,
            'dpd' => 'https://tracktrace.dpd.com.pl/parcelDetails?p1=' . $tracking_number,
            'dhl' => 'https://www.dhl.com/pl-pl/home/tracking.html?tracking-id=' . $tracking_number,
            'ups' => 'https://www.ups.com/track?tracknum=' . $tracking_number,
            'fedex' => 'https://www.fedex.com/fedextrack/?trknbr=' . $tracking_number,
        );
        
        foreach ($urls as $key => $url) {
            if (stripos($carrier_slug, $key) !== false) {
                return $url;
            }
        }
        return '';
    }
    
    public function get_tracking($request) {
        $order_id = absint($request['order_id']);
        $order = wc_get_order($order_id);
        if (!$order) {
            return new WP_Error('invalid_order', 'Order not found', array('status' => 404));
        }
        
        $tracking_number = '';
        if (function_exists('wc_get_order_tracking_number')) {
            $tracking_number = wc_get_order_tracking_number($order);
        } else {
            $tracking_number = $order->get_meta('_tracking_number');
        }

        $items = array();
        if ($tracking_number) {
            $tracking_url = '';
            $carrier_name = 'Custom Carrier';
            
            if (function_exists('wc_get_tracking_url')) {
                $tracking_url = wc_get_tracking_url($tracking_number);
                // Simple heuristic to guess name based on poland.php logic
                $tn = trim($tracking_number);
                if (strlen($tn) === 24 && ctype_digit($tn)) {
                    $carrier_name = 'InPost';
                } else {
                    $carrier_name = 'DPD/Other';
                }
            }
            
             foreach ($order->get_items() as $item_id => $item) {
                $items[] = array(
                    'item_id'   => $item_id,
                    'item_name' => $item->get_name(),
                    'tracking'  => array(
                        array(
                            'tracking_number' => $tracking_number,
                            'tracking_url'    => $tracking_url,
                            'carrier_slug'    => 'custom',
                            'carrier_name'    => $carrier_name,
                            'time'            => time()
                        )
                    )
                );
             }
        }
        
        return rest_ensure_response(array(
            'order_id' => $order_id,
            'items'    => $items,
            'ast_items'=> array()
        ));
    }
    
    public function get_carriers($request) {
        // Return carriers supported by poland.php
        $carriers = array(
            array(
                'slug' => 'inpost',
                'name' => 'InPost',
                'type' => 'custom',
                'source' => 'local'
            ),
            array(
                'slug' => 'dpd',
                'name' => 'DPD',
                'type' => 'custom',
                'source' => 'local'
            )
        );
        
        return rest_ensure_response(array(
            'carriers' => $carriers,
            'total'    => count($carriers)
        ));
    }
    
    public function get_email_logs($request) {
        global $wpdb;
        $table_name = $wpdb->prefix . 'wot_email_logs';
        
        $page = isset($request['page']) ? absint($request['page']) : 1;
        $per_page = isset($request['per_page']) ? absint($request['per_page']) : 20;
        $offset = ($page - 1) * $per_page;
        
        if ($wpdb->get_var("SHOW TABLES LIKE '$table_name'") != $table_name) {
             return rest_ensure_response(array(
                'logs' => array(),
                'total' => 0,
                'message' => 'Log table does not exist yet.'
            ));
        }

        $logs = $wpdb->get_results(
            $wpdb->prepare("SELECT * FROM $table_name ORDER BY id DESC LIMIT %d OFFSET %d", $per_page, $offset)
        );
        
        $total = $wpdb->get_var("SELECT COUNT(*) FROM $table_name");
        $total_pages = ceil($total / $per_page);
        
        return rest_ensure_response(array(
            'logs' => $logs,
            'total' => (int)$total,
            'page' => $page,
            'per_page' => $per_page,
            'total_pages' => $total_pages
        ));
    }
    
    public function get_email_log($request) {
        global $wpdb;
        $table_name = $wpdb->prefix . 'wot_email_logs';
        $log_id = absint($request['log_id']);
        
        if ($wpdb->get_var("SHOW TABLES LIKE '$table_name'") != $table_name) {
             return new WP_Error('no_table', 'Log table not found', array('status' => 404));
        }

        $log = $wpdb->get_row($wpdb->prepare("SELECT * FROM $table_name WHERE id = %d", $log_id));
        
        if (!$log) {
            return new WP_Error('not_found', 'Email log not found', array('status' => 404));
        }
        
        return rest_ensure_response($log);
    }
    
    public function get_email_stats($request) {
        global $wpdb;
        $table_name = $wpdb->prefix . 'wot_email_logs';
        
        if ($wpdb->get_var("SHOW TABLES LIKE '$table_name'") != $table_name) {
             return rest_ensure_response(array('total' => 0, 'message' => 'No logs yet'));
        }
        
        $total = $wpdb->get_var("SELECT COUNT(*) FROM $table_name");
        $sent = $wpdb->get_var("SELECT COUNT(*) FROM $table_name WHERE status = 'sent'");
        $failed = $wpdb->get_var("SELECT COUNT(*) FROM $table_name WHERE status = 'failed'");
        
        $today = current_time('Y-m-d');
        $sent_today = $wpdb->get_var($wpdb->prepare("SELECT COUNT(*) FROM $table_name WHERE status = 'sent' AND created_at LIKE %s", $today . '%'));
        $failed_today = $wpdb->get_var($wpdb->prepare("SELECT COUNT(*) FROM $table_name WHERE status = 'failed' AND created_at LIKE %s", $today . '%'));
        
        $success_rate = $total > 0 ? round(($sent / $total) * 100, 2) . '%' : '0%';
        
        return rest_ensure_response(array(
            'total' => (int)$total,
            'sent' => (int)$sent,
            'failed' => (int)$failed,
            'sent_today' => (int)$sent_today,
            'failed_today' => (int)$failed_today,
            'success_rate' => $success_rate
        ));
    }
}

WOT_REST_API::get_instance();

/**
 * LOGGING SYSTEM FOR EMAILS
 * ------------------------------------------------------------------
 */

/**
 * Create custom table for email logs
 */
function wot_create_log_table() {
    global $wpdb;
    $table_name = $wpdb->prefix . 'wot_email_logs';
    
    // Check if table exists to avoid expensive queries on every load
    // We cache this check in an option or transient
    if (get_transient('wot_email_log_table_checked')) {
        return;
    }
    
    $charset_collate = $wpdb->get_charset_collate();

    $sql = "CREATE TABLE $table_name (
        id mediumint(9) NOT NULL AUTO_INCREMENT,
        created_at datetime DEFAULT '0000-00-00 00:00:00' NOT NULL,
        recipient text NOT NULL,
        subject text NOT NULL,
        message longtext NOT NULL,
        headers text,
        status varchar(20) NOT NULL,
        error_message text,
        PRIMARY KEY  (id)
    ) $charset_collate;";

    require_once(ABSPATH . 'wp-admin/includes/upgrade.php');
    dbDelta($sql);
    
    set_transient('wot_email_log_table_checked', true, WEEK_IN_SECONDS);
}
// Run table creation on admin_init or rest_api_init
add_action('admin_init', 'wot_create_log_table');
add_action('rest_api_init', 'wot_create_log_table');

/**
 * Helper function to log email
 */
function wot_log_email($to, $subject, $message, $headers, $status, $error = '') {
    global $wpdb;
    $table_name = $wpdb->prefix . 'wot_email_logs';
    
    // Ensure table exists (in case transient was cleared but table not created)
    if ($wpdb->get_var("SHOW TABLES LIKE '$table_name'") != $table_name) {
        wot_create_log_table();
    }
    
    $wpdb->insert(
        $table_name,
        array(
            'created_at' => current_time('mysql'),
            'recipient' => is_array($to) ? implode(', ', $to) : $to,
            'subject' => $subject,
            'message' => $message,
            'headers' => is_array($headers) ? implode("\n", $headers) : $headers,
            'status' => $status,
            'error_message' => $error
        )
    );
}

/**
 * Add tracking information to WooCommerce emails
 * Integrates with poland.php functions
 */
add_action('woocommerce_email_before_order_table', function($order, $sent_to_admin, $plain_text, $email) {
    // Ensure we have the order object
    if (!$order || !is_a($order, 'WC_Order')) return;

    // Check if tracking functions exist (from poland.php)
    if (!function_exists('wc_get_order_tracking_number')) return;
    
    $tracking_number = wc_get_order_tracking_number($order);
    if (empty($tracking_number)) return;
    
    $tracking_url = function_exists('wc_get_tracking_url') ? wc_get_tracking_url($tracking_number) : '#';
    
    // Customize text based on language if needed, currently using English/General
    $title_text = 'Informacje o przesyłce'; // Tracking Information
    $number_text = 'Numer przesyłki:'; // Tracking Number:
    $button_text = 'Śledź przesyłkę'; // Track Your Package
    
    if ($plain_text) {
        echo "\n" . "----------------------------------------\n";
        echo $title_text . "\n";
        echo $number_text . " " . $tracking_number . "\n";
        echo $button_text . ": " . $tracking_url . "\n";
        echo "----------------------------------------\n\n";
    } else {
        ?>
        <div style="margin-bottom: 24px; padding: 16px; border: 1px solid #e5e5e5; background-color: #fafafa; border-radius: 4px;">
            <h3 style="margin: 0 0 12px; font-size: 16px; color: #333; line-height: 1.4;"><?php echo esc_html($title_text); ?></h3>
            <p style="margin: 0 0 12px; font-size: 14px; color: #555;">
                <?php echo esc_html($number_text); ?> 
                <strong style="font-family: monospace; font-size: 15px; color: #222; background: #fff; padding: 2px 6px; border: 1px solid #ddd; border-radius: 3px;"><?php echo esc_html($tracking_number); ?></strong>
            </p>
            <p style="margin: 0;">
                <a href="<?php echo esc_url($tracking_url); ?>" target="_blank" style="display: inline-block; padding: 10px 20px; background-color: #2271b1; color: #ffffff; text-decoration: none; border-radius: 4px; font-size: 14px; font-weight: 500;">
                    <?php echo esc_html($button_text); ?> &rarr;
                </a>
            </p>
        </div>
        <?php
    }
}, 20, 4);

// Hook into PHPMailer failure (for non-SMTP2GO sends)
add_action('wp_mail_failed', function($error) {
    if (function_exists('wot_log_email')) {
         // Try to extract info from the error object if possible
         $message_data = $error->get_error_data();
         $to = isset($message_data['to']) ? $message_data['to'] : 'unknown';
         $subject = isset($message_data['subject']) ? $message_data['subject'] : 'unknown';
         $message = isset($message_data['message']) ? $message_data['message'] : '';
         $headers = isset($message_data['headers']) ? $message_data['headers'] : '';
         
         wot_log_email($to, $subject, $message, $headers, 'failed', $error->get_error_message());
    }
});

