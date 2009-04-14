<?php
/**
 * The Spreadsheet plugin script
 *
 * @package Spreadsheet
 * @author WebServesUs
 * @copyright City Lore, 2009
 * TODO: check licensing options!
 */

define('SPREADSHEET_VERSION', '1.0');

define('SPREADSHEET_STATUS_INIT', 1);
define('SPREADSHEET_STATUS_COMPLETE', 2);
define('SPREADSHEET_STATUS_PURGED', 3);
define('SPREADSHEET_STATUS_ERROR', 4);

define('SPREADSHEET_FILES_DIR', PLUGIN_DIR . '/Spreadsheet/files');

add_plugin_hook('install', 'spreadsheet_install');
add_plugin_hook('config', 'spreadsheet_config');
add_plugin_hook('config_form', 'spreadsheet_config_form');
add_plugin_hook('uninstall', 'spreadsheet_uninstall');
add_plugin_hook('admin_append_to_dashboard_primary', 'spreadsheet_dashboard');
add_plugin_hook('admin_append_to_items_browse_primary', 'spreadsheet_export_link');
add_plugin_hook('define_routes', 'spreadsheet_routes');

function spreadsheet_install() {
	set_option('spreadsheet_version', SPREADSHEET_VERSION);
	//number of days to keep export files before purging
	set_option('spreadsheet_expiry', '90');
	
	$db = get_db();
	$db->exec(
		"CREATE TABLE IF NOT EXISTS `{$db->prefix}spreadsheets` (
			`id` INT UNSIGNED NOT NULL AUTO_INCREMENT,
			`status` INT,
			`user_id` INT,
			`file_name` VARCHAR(255),
			`terms` TEXT,
			`added` DATETIME,
			PRIMARY KEY  (`id`)
       ) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"
	);	
}

function spreadsheet_config() {
	set_option('spreadsheet_php_path', $_POST['spreadsheet_php_path']);
}

function spreadsheet_config_form() {
	if (!$path = get_option('spreadsheet_php_path')) {
      // Get the path to the PHP-CLI command. This does not account for
      // servers without a PHP CLI or those with a different command name for
      // PHP, such as "php5".
      $command = 'which php 2>&0';
      $lastLineOutput = exec($command, $output, $returnVar);
      $path = $returnVar == 0 ? trim($lastLineOutput) : '';
  }
?>
  <div class="field">
      <label for="spreadsheet_php_path">Path to PHP-CLI</label>
      <?php echo __v()->formText('spreadsheet_php_path', $path, null);?>
      <p class="explanation">Path to your server's PHP-CLI command. The PHP
      version must correspond to normal Omeka requirements. Some web hosts use PHP
      4.x for their default PHP-CLI, but many provide an alternative path to a
      PHP-CLI 5 binary. Check with your web host for more information.</p>
  </div>
<?php
}

function spreadsheet_uninstall() {
	$db = get_db();
	$db->exec("DROP TABLE IF EXISTS {$db->prefix}spreadsheets");
	delete_option('spreadsheet_version');
	delete_option('spreadsheet_expiry');
}

function spreadsheet_dashboard() {
	?>
	<dt class="spreadsheet"><a href="<?php echo uri('spreadsheet/index/status'); ?>">Spreadsheet Exports</a></dt>
	<dd class="spreadsheet"></dd>
	<p>Check the status or re-download an exported spreadsheet. To created an export, do a search then click on "Export results as spreadsheet".</p>
	<?php
}

function spreadsheet_export_link($items) {
	if (isset($_GET['search']) && count($items)) {
		$params = array();
		foreach ($_GET as $k => $v) {
			$params[$k] = $v;
			//set per_page query param here to force a complete (i.e. unpaginated) list
			//use the total_results count set in the Registry by the SearchItems helper
			$params['hits'] = ZEND_REGISTRY::get('total_results');
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
	$router->addRoute(
		'spreadsheet_download_route',
		new Zend_Controller_Router_Route(
			'spreadsheet/download/:id',
			array('module' => 'spreadsheet', 'controller' => 'index', 'action' => 'download')
		)
	);
}
?>