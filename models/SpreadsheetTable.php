<?php
class SpreadsheetTable extends Omeka_Db_Table {
	
	public function findSpreadsheetsByUserId($id) {
		$select = $this->getSelect()->where('user_id = ?');
		return $this->fetchObjects($select, array($id));
	}
}

?>