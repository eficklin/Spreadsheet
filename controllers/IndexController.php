<?php
/**
 * The Spreadsheet index controller class.
 *
 * @package Spreadsheet
 * @author WebServesUs
 * @copyright City Lore, 2009
 */

require 'Spreadsheet.php';

class Spreadsheet_IndexController extends Omeka_Controller_Action {
	
	public function init() {}
	
	public function indexAction() {}
	
	/**
	* Creates an Excel spreadsheet given the search terms supplied by the 
	* search box or advanced search form (captured from the URL query params) 
	* and sends spreadsheet to browser with headers set to force download
	*
	* @return void
	*/
	public function xlsAction() {	
		//no view needed, will force download of spreadsheet
		$this->_helper->viewRenderer->setNoRender();
		
		//perform search via helper
		$results = $this->_helper->searchItems();
		
		$spreadsheet = new Spreadsheet(current_user()->id, 'OmekaExport-' . time() . '.xls', $results);
		$spreadhseet->save();
		
		$args = escapeshellarg("-u {$spreadsheet->user_id}") . escapeshellarg("-i {$spreadsheet->id}");
		$php_path = get_option('spreadsheet_php_path');
		$script_path = PLUGIN_DIR . "/Spreadsheet/background_scripts/export.php";
		exec('nice ' . $php_path . ' ' . $script_path . ' ' . $args);
		$this->goto('status');
		/*
		//add second worksheet for search terms
		$this->xls->createSheet();
		$this->xls->setActiveSheetIndex(1);
		$this->xls->getActiveSheet()->setTitle('Search Terms');
		$this->xls->getActiveSheet()->setCellValue('A1', 'Search Terms');
		*/
		/* get search terms from query params and echo in spreadsheet as record of the 
		search that produced this result set */	
		/*
		$next_cell = 2;
		foreach ($_GET as $k => $v) {
			if ($k == 'advanced') {
				foreach($v as $advanced_term) {
					$this->xls->getActiveSheet()->setCellValue('A' . $next_cell,
						$advanced_term['element_id'] . ' ' . $advanced_term['type'] . ' ' . $advanced_term['terms']
					);
					$next_cell++; 
				}
			} else {
				$this->xls->getActiveSheet()->setCellValue('A' . $next_cell, "{$k}: {$v}");
				$next_cell++;
			}
		}
		//finishing touch: set active sheet back to 0 so it's active when the use opens the xls file
		$this->xls->setActiveSheetIndex(0);
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachement;filename="OmekaExport.xls"');
		$this->xls_writer->save('php://output');
		*/
	}
	
	public function statusAction() {
		
	}
}