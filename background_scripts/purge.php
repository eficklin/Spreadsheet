<?php
// Require the core application and plugin files
$baseDir = str_replace('plugins/Spreadsheet/background_scripts', '', dirname(__FILE__));
require "{$baseDir}paths.php";
require "{$baseDir}application/libraries/Omeka/Core.php";

// Load only the required core phases.
$core = new Omeka_Core;
$core->phasedLoading('initializePluginBroker');

// Get the database object.
$db = get_db();

require 'Spreadsheet.php';
$expired_spreadsheets = $db->getTable('Spreadsheet')->getExpiredSpreadsheets();
foreach ($expired_spreadsheets as $s) {
	unlink($s->getFilePath());
	$s->status = SPREADSHEET_STATUS_PURGED;
	$s->save();
}

?>