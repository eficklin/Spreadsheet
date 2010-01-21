<?php
require 'spreadsheet_functions.php';

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
$acl = Omeka_Context::getInstance()->getAcl();

// Get spreadsheet information from DB
require 'Spreadsheet.php';
$spreadsheet_id = $options['i'];
$spreadsheet = $db->getTable('Spreadsheet')->find($spreadsheet_id);
$terms = unserialize($spreadsheet->terms);

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

//column headings
$set = new ElementSet();
$set->name = 'Dublin Core';
$elements = $set->getElements();

$xls->getActiveSheet()->setCellValue('A1', 'Omeka ID');
$xls->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
$xls->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);

$xls->getActiveSheet()->setCellValue('B1', 'Title');
$xls->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
$xls->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);

$xls->getActiveSheet()->setCellValue('C1', 'Subject');
$xls->getActiveSheet()->getStyle('C1')->getFont()->setBold(true);
$xls->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);

$xls->getActiveSheet()->setCellValue('D1', 'Reference Image');
$xls->getActiveSheet()->getStyle('D1')->getFont()->setBold(true);
$xls->getActiveSheet()->getColumnDimension('D')->setWidth(35);

$col = 'E';
$row = 1;
foreach ($elements as $e) {
	if ($e->name == 'Title' || $e->name == 'Subject')
		continue;
		
	$xls->getActiveSheet()->SetCellValue($col . $row, $e->name);
	$xls->getActiveSheet()->getStyle($col . $row)->getFont()->setBold(true);
	$xls->getActiveSheet()->getColumnDimension($col)->setAutoSize(true);
	$col = chr(ord($col) + 1);
}

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
$row = 2;

$results = spreadsheet_search($spreadsheet, $chunk);

foreach ($results as $i) {
	//set row height to accomodate reference image and longer element texts
	$xls->getActiveSheet()->getRowDimension($row)->setRowHeight(105);
	
	//omkea id
	$xls->getActiveSheet()->setCellValue('A' . $row, $i->id);
	$xls->getActiveSheet()->getStyle('A' . $row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
	
	//dublin core elements
	//title and subject first (cols b, c)
	$texts = $i->getElementTextsByElementNameAndSetName('Title', 'Dublin Core');
	$texts_to_join = array();
	if (count($texts)) {
		foreach ($texts as $t) {
			$texts_to_join[] = $t->html ? cleanHTML($t->text) : $t->text;
		}
	}
	$xls->getActiveSheet()->SetCellValueExplicit('B' . $row, xlsLineBreaks(implode('; ', $texts_to_join)), PHPExcel_Cell_DataType::TYPE_STRING);
	$xls->getActiveSheet()->getStyle('B' . $row)->getAlignment()->setWrapText(true);
	$xls->getActiveSheet()->getStyle('B' . $row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
	
	$texts = $i->getElementTextsByElementNameAndSetName('Subject', 'Dublin Core');
	$texts_to_join = array();
	if (count($texts)) {
		foreach ($texts as $t) {
			$texts_to_join[] = $t->html ? cleanHTML($t->text) : $t->text;
		}
	}
	$xls->getActiveSheet()->SetCellValueExplicit('C' . $row, xlsLineBreaks(implode('; ', $texts_to_join)), PHPExcel_Cell_DataType::TYPE_STRING);
	$xls->getActiveSheet()->getStyle('C' . $row)->getAlignment()->setWrapText(true);
	$xls->getActiveSheet()->getStyle('C' . $row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
	
	//image (col d): thumbnail image if available
	$files = $i->getFiles();
	if (count($files) && $files[0]->hasThumbnail()) {
		$img = new PHPExcel_Worksheet_Drawing();
		$img->setName('Reference Image');
		$img->setDescription('Reference Image');
		$img->setPath($files[0]->getPath('thumbnail'));
		$img->setHeight(100);
		$img->setOffsetX(5);
		$img->setOffsetY(5);
		$img->setCoordinates('D' . $row);
		$img->setWorksheet($xls->getActiveSheet());
	} else {
		$xls->getActiveSheet()->setCellValue('D' . $row, "[no image available]");
		$xls->getActiveSheet()->getStyle('D' . $row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
	}
	release_object($files);
	
	//remaining dc elements (cols e and beyond)
	$col = 'E';
	foreach ($elements as $e) {
		if ($e->name == "Title" || $e->name == "Subject")
			continue;
			
		$texts = $i->getElementTextsByElementNameAndSetName($e->name, 'Dublin Core');
		$texts_to_join = array();
		if (count($texts)) {
			foreach ($texts as $t) {
				$texts_to_join[] = $t->html ? cleanHTML($t->text) : $t->text;
			}
		}
		$xls->getActiveSheet()->SetCellValueExplicit($col . $row, xlsLineBreaks(implode('; ', $texts_to_join)), PHPExcel_Cell_DataType::TYPE_STRING);
		$xls->getActiveSheet()->getStyle($col . $row)->getAlignment()->setWrapText(true);
		$xls->getActiveSheet()->getStyle($col . $row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
		$col = chr(ord($col) + 1);
	}

	//collection
	$col = chr(ord($col) + 1);
	$xls->getActiveSheet()->setCellValue($col . $row, $i->getCollection()->name);
	$xls->getActiveSheet()->getStyle($col . $row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

	//item type
	$col = chr(ord($col) + 1);
	$xls->getActiveSheet()->setCellValue($col . $row, $i->getItemType()->name);
	$xls->getActiveSheet()->getStyle($col . $row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

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
		release_object($metadata);
	}
	$xls->getActiveSheet()->setCellValue($col . $row, $metatexts);
	$xls->getActiveSheet()->SetCellValueExplicit($col . $row, implode('; ', $texts_to_join), PHPExcel_Cell_DataType::TYPE_STRING);
	$xls->getActiveSheet()->getStyle($col . $row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
	$row++;
	release_object($i);
}

//add second worksheet to record search terms that produced this result set
$xls->createSheet();
$xls->setActiveSheetIndex(1);
$xls->getActiveSheet()->setTitle('Search Terms');
$xls->getActiveSheet()->setCellValue('A1', 'Search Terms');

$next_cell = 2;
foreach ($terms as $k => $v) {
	if ($k == 'advanced') {
		foreach($v as $advanced_term) {
			$xls->getActiveSheet()->setCellValue('A' . $next_cell,
				$advanced_term['element_id'] . ' ' . $advanced_term['type'] . ' ' . $advanced_term['terms']
			);
			$next_cell++; 
		}
	} else {
		$xls->getActiveSheet()->setCellValue('A' . $next_cell, "{$k}: {$v}");
		$next_cell++;
	}
}
//finishing touch: set active sheet back to 0 so it's active when the use opens the xls file
$xls->setActiveSheetIndex(0);

//save to disk
$xls_writer->save(SPREADSHEET_FILES_DIR . "/{$spreadsheet->file_name}");

//update status in db
$spreadsheet->status = SPREADSHEET_STATUS_COMPLETE;
$spreadsheet->save();
?>