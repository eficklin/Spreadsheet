<?php
/**
 * The Spreadsheet index controller class.
 *
 * @package Spreadsheet
 * @author WebServesUs
 * @copyright City Lore, 2009
 */

class Spreadsheet_IndexController extends Omeka_Controller_Action {
	
	public $xls;
	public $xls_writer;
	
	public function init() {
		//PHPExcel package needs an addition to the include path to find its lib of classes
		set_include_path(get_include_path() . PATH_SEPARATOR . PLUGIN_DIR . '/Spreadsheet/PHPExcel/Classes');
		
		//create the necessary PHPExcel classes
		include 'PHPExcel.php';
		include 'PHPExcel/IOFactory.php';
		$this->xls = new PHPExcel();
		/* using Excel5 (97/XP/03 office versions) format to avoid a dependecy on the php_zip extension 
		 a necessity for the Excel 2007 format, which is not so common in shared hosting enviroments 
		 as the creators of the PHPExcel lib assume */
		$this->xls_writer = PHPExcel_IOFactory::createWriter($this->xls, 'Excel5');
	}
	
	
	public function indexAction() {
		
	}
	
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
		
		//items worksheet
		$row = 3;
		foreach ($results['items'] as $i) {
			//set row height to accomodate reference image and longer element texts
			$this->xls->getActiveSheet()->getRowDimension($row)->setRowHeight(105);
				
			//dublin core elements
			$col = 'A';
			foreach ($elements as $e) {
				$texts = $i->getElementTextsByElementNameAndSetName($e->name, 'Dublin Core');
				$texts_to_join = array();
				if (count($texts)) {
					foreach ($texts as $t) {
						$texts_to_join[] = $t->html ? $this->_cleanHTML($t->text) : $t->text;
					}
				}
				$this->xls->getActiveSheet()->SetCellValue($col . $row, implode('; ', $texts_to_join));
				$this->xls->getActiveSheet()->getStyle($col . $row)->getAlignment()->setWrapText(true);
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
				$img->setWorksheet($this->xls->getActiveSheet());
			} else {
				$this->xls->getActiveSheet()->setCellValue($col . $row, "[no image available]");
			}
			
			//collection
			$col = chr(ord($col) + 1);
			$this->xls->getActiveSheet()->setCellValue($col . $row, $i->getCollection()->name);
			
			//item type
			$col = chr(ord($col) + 1);
			$this->xls->getActiveSheet()->setCellValue($col . $row, $i->getItemType()->name);
			
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
							$texts_to_join[] = $t->html ? $this->_cleanHTML($t->text) : $t->text;
						}
						$metatexts .= implode(", ", $texts_to_join) . "; ";
					}
				}
			}
			$this->xls->getActiveSheet()->setCellValue($col . $row, $metatexts);
			$row++;
		}
		
		//add second worksheet for search terms
		$this->xls->createSheet();
		$this->xls->setActiveSheetIndex(1);
		$this->xls->getActiveSheet()->setTitle('Search Terms');
		$this->xls->getActiveSheet()->setCellValue('A1', 'Search Terms');
		
		/* get search terms from query params and echo in spreadsheet as record of the 
		search that produced this result set */	
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
	}
	
	protected function _cleanHTML($text) {
		$text = html_entity_decode($text);
		return strip_tags($text);
	}
}