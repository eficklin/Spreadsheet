<?php
/**
 * The Spreadsheet record class.
 *
 * @package Spreadsheet
 * @author WebServesUs
 * @copyright City Lore 2009
 */
require 'SpreadsheetTable.php';

class Spreadsheet extends Omeka_Record {
	public $user_id;
	public $file_name;
	public $status;
	public $items;
	public $terms;
	public $added;
	
	public function construct() {
		$this->user_id = current_user()->id;
		$this->file_name = "OmekaExport" . time() . ".xls";
		$this->status = SPREADSHEET_STATUS_INIT;
		$this->added = date('Y-m-d H:m:s');
	}
	
	/**
	 * returns User who initiated export
	 * @return User
	 */
	public function getUser() {
		return $this->getTable('User')->find($this->user_id);
	}
	
	/**
	 * returns array of spreadsheets created by user
	 * @return Array
	 */
	public function getUserSpreadsheets() {
		return $this->findSpreadsheetsByUserId($this->user_id);
	}
	
	/**
	 * returns path to file or null if file does not exist
	 * @return String|null
	 */
	public function getFilePath() {
		$path = SPREADSHEET_FILES_DIR . "/" . $this->file_name;
		if (file_exists($path)) {
			return $path;
		}
	}
}
?>