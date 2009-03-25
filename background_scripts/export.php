<?php

// Require the core application and plugin files
$baseDir = str_replace('plugins/Spreadsheet/background_scripts', '', dirname(__FILE__));
require "{$baseDir}paths.php";
require "{$baseDir}application/libraries/Omeka/Core.php";

// Load only the required core phases.
$core = new Omeka_Core;
$core->phasedLoading('initializeCurrentUser');

// Set the memory limit.
$memoryLimit = get_option('spreadsheet_memory_limit');
ini_set('memory_limit', "$memoryLimit");

// Get the database object.
$db = get_db();

// Set the command line arguments.
$options = getopt('i:u:');

// Get the user object and set the current user to it
$userId = $options['u'];
$user = $db->getTable('User')->find($userId);
Omeka_Context::getInstance()->setCurrentUser($user);

//PHPExcel package needs an addition to the include path to find its lib of classes
set_include_path(get_include_path() . PATH_SEPARATOR . PLUGIN_DIR . '/Spreadsheet/PHPExcel/Classes');

//create the necessary PHPExcel classes
include 'PHPExcel.php';
include 'PHPExcel/IOFactory.php';
$xls = new PHPExcel();
/* using Excel5 (97/XP/03 office versions) format to avoid a dependecy on the php_zip extension 
 a necessity for the Excel 2007 format, which is not so common in shared hosting enviroments 
 as the creators of the PHPExcel lib assume */
$xls_writer = PHPExcel_IOFactory::createWriter($this->xls, 'Excel5');

$this->xls->setActiveSheetIndex(0);
$this->xls->getActiveSheet()->getDefaultStyle()->getFont()->setName('Arial');
$this->xls->getActiveSheet()->setTitle('Items');
$this->xls->getActiveSheet()->setCellValue('A1', "Omeka Export -- " . date('F j, Y'));
$this->xls->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);

//column headings
$set = new ElementSet();
$set->name = 'Dublin Core';
$elements = $set->getElements();

$col = 'A';
$row = 2;
foreach ($elements as $e) {
	$this->xls->getActiveSheet()->SetCellValue($col . $row, $e->name);
	$this->xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);
	$this->xls->getActiveSheet()->getColumnDimension($col)->setAutoSize(true);
	$col = chr(ord($col) + 1);
}

$this->xls->getActiveSheet()->setCellValue($col . $row, "Reference Image");
$this->xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);
$this->xls->getActiveSheet()->getColumnDimension($col)->setWidth(35);

$col = chr(ord($col) + 1);
$this->xls->getActiveSheet()->setCellValue($col . $row, "Collection");
$this->xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);

$col = chr(ord($col) + 1);
$this->xls->getActiveSheet()->setCellValue($col . $row, "Item Type");
$this->xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);

$col = chr(ord($col) + 1);
$this->xls->getActiveSheet()->setCellValue($col . $row, "Item Type Metadata");
$this->xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);

//get items one chunk at a time
//loop through items in chunk adding to spreadhseet

//finish up spreadsheet
//update status in db
//save to disk