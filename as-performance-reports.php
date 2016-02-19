<?php

/**
 * The plugin bootstrap file
 *
 * This file is read by WordPress to generate the plugin information in the plugin
 * admin area. This file also includes all of the dependencies used by the plugin,
 * registers the activation and deactivation functions, and defines a function
 * that starts the plugin.
 *
 * @link              https://www.facebook.com/anuragsingh.me
 * @since             1.0.0
 * @package           As_Performance_Reports
 *
 * @wordpress-plugin
 * Plugin Name:       AS Performance Reports
 * Plugin URI:        http://www.dselva.co.in
 * Plugin Type:       Piklist
 * Description:       This is a short description of what the plugin does. It's displayed in the WordPress admin area.
 * Version:           1.0.0
 * Author:            Anurag Singh
 * Author URI:        https://www.facebook.com/anuragsingh.me
 * License:           GPL-2.0+
 * License URI:       http://www.gnu.org/licenses/gpl-2.0.txt
 * Text Domain:       as-performance-reports
 * Domain Path:       /languages
 */

// If this file is called directly, abort.
if ( ! defined( 'WPINC' ) ) {
    die;
}

/**
 * The code that runs during plugin activation.
 * This action is documented in includes/class-as-performance-reports-activator.php
 */
function activate_as_performance_reports() {
    require_once plugin_dir_path( __FILE__ ) . 'includes/class-as-performance-reports-activator.php';
    As_Performance_Reports_Activator::activate();
}

/**
 * The code that runs during plugin deactivation.
 * This action is documented in includes/class-as-performance-reports-deactivator.php
 */
function deactivate_as_performance_reports() {
    require_once plugin_dir_path( __FILE__ ) . 'includes/class-as-performance-reports-deactivator.php';
    As_Performance_Reports_Deactivator::deactivate();
}

register_activation_hook( __FILE__, 'activate_as_performance_reports' );
register_deactivation_hook( __FILE__, 'deactivate_as_performance_reports' );

/**
 * The core plugin class that is used to define internationalization,
 * admin-specific hooks, and public-facing site hooks.
 */
require plugin_dir_path( __FILE__ ) . 'includes/class-as-performance-reports.php';

/**
 * Begins execution of the plugin.
 *
 * Since everything within the plugin is registered via hooks,
 * then kicking off the plugin from this point in the file does
 * not affect the page life cycle.
 *
 * @since    1.0.0
 */
function run_as_performance_reports() {

    $plugin = new As_Performance_Reports();
    $plugin->run();

}
run_as_performance_reports();
