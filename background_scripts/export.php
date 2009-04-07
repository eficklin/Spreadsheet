<?php
//utility function for cleanup
function cleanHTML($text) {
	$text = html_entity_decode($text);
	return strip_tags($text);
}

// Require the core application and plugin files
$baseDir = str_replace('plugins/Spreadsheet/background_scripts', '', dirname(__FILE__));
require "{$baseDir}paths.php";
require "{$baseDir}application/libraries/Omeka/Core.php";

// Load only the required core phases.
$core = new Omeka_Core;
$core->phasedLoading('initializeCurrentUser');

// Get the database object.
$db = get_db();

// Set the command line arguments.
$options = getopt('u:i:');

// Get the user object and set the current user to it
$userId = $options['u'];
$user = $db->getTable('User')->find($userId);
Omeka_Context::getInstance()->setCurrentUser($user);

// Get spreadsheet information from DB
require 'Spreadsheet.php';
$spreadsheet_id = $options['i'];
$spreadsheet = $db->getTable('Spreadsheet')->find($spreadsheet_id);

//PHPExcel package needs an addition to the include path to find its lib of classes
set_include_path(get_include_path() . PATH_SEPARATOR . PLUGIN_DIR . '/Spreadsheet/PHPExcel/Classes');

//create the necessary PHPExcel classes
include 'PHPExcel.php';
include 'PHPExcel/IOFactory.php';
$xls = new PHPExcel();
/* using Excel5 (97/XP/03 office versions) format to avoid a dependecy on the php_zip extension 
 a necessity for the Excel 2007 format, which is not so common in shared hosting enviroments 
 as the creators of the PHPExcel lib assume */
$xls_writer = PHPExcel_IOFactory::createWriter($xls, 'Excel5');

$xls->setActiveSheetIndex(0);
$xls->getActiveSheet()->getDefaultStyle()->getFont()->setName('Arial');
$xls->getActiveSheet()->setTitle('Items');
$xls->getActiveSheet()->setCellValue('A1', "Omeka Export -- " . date('F j, Y'));
$xls->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);

//column headings
$set = new ElementSet();
$set->name = 'Dublin Core';
$elements = $set->getElements();

$col = 'A';
$row = 2;
foreach ($elements as $e) {
	$xls->getActiveSheet()->SetCellValue($col . $row, $e->name);
	$xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);
	$xls->getActiveSheet()->getColumnDimension($col)->setAutoSize(true);
	$col = chr(ord($col) + 1);
}

$xls->getActiveSheet()->setCellValue($col . $row, "Reference Image");
$xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);
$xls->getActiveSheet()->getColumnDimension($col)->setWidth(35);

$col = chr(ord($col) + 1);
$xls->getActiveSheet()->setCellValue($col . $row, "Collection");
$xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);

$col = chr(ord($col) + 1);
$xls->getActiveSheet()->setCellValue($col . $row, "Item Type");
$xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);

$col = chr(ord($col) + 1);
$xls->getActiveSheet()->setCellValue($col . $row, "Item Type Metadata");
$xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);

//items worksheet
$row = 3;
foreach ($spreadsheet->items as $i) {
	//set row height to accomodate reference image and longer element texts
	$xls->getActiveSheet()->getRowDimension($row)->setRowHeight(105);
		
	//dublin core elements
	$col = 'A';
	foreach ($elements as $e) {
		$texts = $i->getElementTextsByElementNameAndSetName($e->name, 'Dublin Core');
		$texts_to_join = array();
		if (count($texts)) {
			foreach ($texts as $t) {
				$texts_to_join[] = $t->html ? cleanHTML($t->text) : $t->text;
			}
		}
		$xls->getActiveSheet()->SetCellValue($col . $row, implode('; ', $texts_to_join));
		$xls->getActiveSheet()->getStyle($col . $row)->getAlignment()->setWrapText(true);
		$col = chr(ord($col) + 1);
	}
	
	//insert thumbnail image if available
	$files = $i->getFiles();
	if (count($files) && $files[0]->hasThumbnail()) {
		$img = new PHPExcel_Worksheet_Drawing();
		$img->setName('Reference Image');
		$img->setDescription('Reference Image');
		$img->setPath($files[0]->getPath('thumbnail'));
		$img->setHeight(100);
		$img->setOffsetX(5);
		$img->setOffsetY(5);
		$img->setCoordinates($col . $row);
		$img->setWorksheet($xls->getActiveSheet());
	} else {
		$xls->getActiveSheet()->setCellValue($col . $row, "[no image available]");
	}
	
	//collection
	$col = chr(ord($col) + 1);
	$xls->getActiveSheet()->setCellValue($col . $row, $i->getCollection()->name);
	
	//item type
	$col = chr(ord($col) + 1);
	$xls->getActiveSheet()->setCellValue($col . $row, $i->getItemType()->name);
	
	//collect Item Type Metadata and join into single string
	$col = chr(ord($col) + 1);
	$metatexts = "";
	$metadata = $i->getItemTypeElements();
	if (count($metadata)) {
		foreach ($metadata as $e) {
			$texts = $i->getElementTextsByElementNameAndSetName($e->name, 'Item Type Metadata');
			$texts_to_join = array();
			if (count($texts)) {
				$metatexts .= $e->name . ": ";
				foreach ($texts as $t) {
					$texts_to_join[] = $t->html ? cleanHTML($t->text) : $t->text;
				}
				$metatexts .= implode(", ", $texts_to_join) . "; ";
			}
		}
	}
	$xls->getActiveSheet()->setCellValue($col . $row, $metatexts);
	$row++;
}

//finish up spreadsheet

//save to disk
$xls_writer->save(SPREADSHEET_FILES_DIR . "/{$spreadsheet->file_name}");

//update status in db
$spreadsheet->status = SPREADSHEET_STATUS_COMPLETE;
$spreadsheet->save();
?>