<?php
class SpreadsheetTable extends Omeka_Db_Table {
	
	public function findSpreadsheetsByUserId($id) {
		$select = $this->getSelect()->where('user_id = ? and status != ?')->order('added DESC');
		return $this->fetchObjects($select, array($id, SPREADSHEET_STATUS_PURGED));
	}
	
	public function getExpiredSpreadsheets($expiry = 90) {
		$purge_date = date('Y-m-d', time() - ($expiry * 86400));
		$select = $this->getSelect()->where('added < ?');
		return $this->fetchObjects($select, array($purge_date));
	}
}

?>