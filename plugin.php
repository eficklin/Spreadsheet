<?php
/**
 * The Spreadsheet plugin script
 *
 * @package Spreadsheet
 * @author WebServesUs
 * @copyright City Lore, 2009
 */

define('SPREADSHEET_VERSION', '1.0');
set_option('spreadsheet_version', SPREADSHEET_VERSION);

add_plugin_hook('admin_append_to_items_browse_primary', 'spreadsheet_export_link');
add_plugin_hook('define_routes', 'spreadsheet_routes');

function spreadsheet_export_link($items) {
	if (isset($_GET['search']) && count($items)) {
		$params = array();
		foreach ($_GET as $k => $v) {
			$params[$k] = $v;
			//set per_page query param here to force a complete (i.e. unpaginated) list
			//use the total_results count set in the Registry by the SearchItems helper
			$params['per_page'] = ZEND_REGISTRY::get('total_results');
		}
		echo "<a href='" . uri('spreadsheet/xls', $params) ."'>Export results as spreadsheet</a>";
	}
}

function spreadsheet_routes($router) {
	$router->addRoute(
		'spreadsheet_xls_route',
		new Zend_Controller_Router_Route(
			'spreadsheet/xls', 
			array('module' => 'spreadsheet', 'controller' => 'index', 'action' => 'xls')
		)
	);
}
?>