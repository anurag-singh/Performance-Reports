<?php

/**
 * Define the internationalization functionality
 *
 * Loads and defines the internationalization files for this plugin
 * so that it is ready for translation.
 *
 * @link       https://www.facebook.com/anuragsingh.me
 * @since      1.0.0
 *
 * @package    As_Performance_Reports
 * @subpackage As_Performance_Reports/includes
 */

/**
 * Define the internationalization functionality.
 *
 * Loads and defines the internationalization files for this plugin
 * so that it is ready for translation.
 *
 * @since      1.0.0
 * @package    As_Performance_Reports
 * @subpackage As_Performance_Reports/includes
 * @author     Anurag Singh <anuragsinghce@outlook.com>
 */
class As_Performance_Reports_i18n {


	/**
	 * Load the plugin text domain for translation.
	 *
	 * @since    1.0.0
	 */
	public function load_plugin_textdomain() {

		load_plugin_textdomain(
			'as-performance-reports',
			false,
			dirname( dirname( plugin_basename( __FILE__ ) ) ) . '/languages/'
		);

	}



}
