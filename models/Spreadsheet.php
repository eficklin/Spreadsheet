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
	
	public function construct($user_id, $file_name, $items, $status = SPREADSHEET_STATUS_INIT) {
		$this->user_id = $user_id;
		$this->file_name = $file_name;
		$this->items = $items;
		$this->status = $status;
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
		return $this->findSpreadSheetsByUserId($this->user_id);
	}
}
?>