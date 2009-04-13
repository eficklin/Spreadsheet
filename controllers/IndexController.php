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
	
	public function init() {
		$this->_modelClass = 'Spreadsheet';
	}
	
	public function indexAction() {}
	
	/**
	* Creates an Excel spreadsheet given the search terms stored in Spreadsheet model class
	* spreadsheet generation happens as background process, user redirected to status page
	* where they can refresh page to check on status and download when finished
	*/
	public function xlsAction() {	
		$spreadsheet = new Spreadsheet;
		$spreadsheet->terms = serialize($_GET);
		$spreadsheet->save();
		
		$args = escapeshellarg("-u {$spreadsheet->user_id}") . ' ' . escapeshellarg("-i {$spreadsheet->id}");
		$php_path = get_option('spreadsheet_php_path');
		$script_path = PLUGIN_DIR . "/Spreadsheet/background_scripts/export.php";
		exec('nice ' . $php_path . ' ' . $script_path . ' ' . $args . ' > /dev/null 2>&1 &');
		$this->redirect->goto('status');
	}
	
	public function statusAction() {
		//get user exports
		$user_id = current_user()->id;
		$exports = $this->getTable('Spreadsheet')->findSpreadsheetsByUserId($user_id);
		$this->view->exports = $exports;
	}
	
	public function downloadAction() {
		$s = $this->findById();
		if ($s->getFilePath()) { 
			//no view needed, will force download of spreadsheet
			$this->_helper->viewRenderer->setNoRender();
			header('Content-Type: application/vnd.ms-excel');
			header("Content-Disposition: attachement;filename='{$s->file_name}'");
			header('Content-Length: ' . filesize($s->getFilePath()));
			readfile($s->getFilePath());
			return;
		} else {
			//file does not exist, what now?
			$this->flashError('That file does not exist. There may be an error or it has been purged from the system. Try creating a new export or contact the site administrator for assistance.');
			$this->redirect->goto('status');
		}
	}
}