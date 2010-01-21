<?php
//utility function for cleanup
function cleanHTML($text) {
	$text = html_entity_decode($text);
	return strip_tags($text);
}

//utility function to format linebreaks correctly for Excel General and Text cells
function xlsLineBreaks($str) {
	return str_replace(chr(13), '', $str);
}

/**
 * our own search function, don't like the duplication, however
 * @param $spreadsheet Spreadsheet model class storing export details
 * @param $chunk integer result sets need to be broken down to avoid memory limit failure 
 * @return Array of Items
 */

function spreadsheet_search($spreadsheet, $chunk = 1) {   
  global $db;
  global $acl;
  global $user;
	//TODO: zend must have some facility for dealing with un/serialized arrays as model data, right?
	$terms = unserialize($spreadsheet->terms);
	$itemTable = $db->getTable('Item');
  $perms  = array();
  $filter = array();
  $order  = array();
    
	//Show only public items
	if ($terms['public']) {
		$perms['public'] = true;
	}
	// User-specific item browsing
	if ($userToView = $terms['user']) {

		// Must be logged in to view items specific to certain users
		if (!$acl->isAllowed($user->role, 'browse', 'Users')) {
			$spreadsheet->status = SPREADSHEET_STATUS_ERROR;
			$spreadsheet->save();
			return;
 		}

		if (is_numeric($userToView)) {
			$filter['user'] = $userToView;
		}
	}
	
	if ($terms['featured']) {
		$filter['featured'] = true;
	}
	
	if ($collection = $terms['collection']) {
		$filter['collection'] = $collection;
	}
	
	if ($type = $terms['type']) {
		$filter['type'] = $type;
	}
	
	if (($tag = $terms['tag']) || ($tag = $terms['tags'])) {
		$filter['tags'] = $tag;
	}
	
	if (($excludeTags = $terms['excludeTags'])) {
		$filter['excludeTags'] = $excludeTags;
	}
        
	$recent = $terms['recent'];
	if ($recent !== 'false') {
		$order['recent'] = true;
	}

	if ($search = $terms['search']) {
		$filter['search'] = $search;
		//Don't order by recent-ness if we're doing a search
		unset($order['recent']);
	}
	
	//The advanced or 'itunes' search
	if ($advanced = $terms['advanced']) {
		//We need to filter out the empty entries if any were provided
		foreach ($advanced as $k => $entry) {                    
			if (empty($entry['element_id']) || empty($entry['type'])) {
				unset($advanced[$k]);
			}
		}
		$filter['advanced_search'] = $advanced;
	}
        	
	if ($range = $terms['range']) {
		$filter['range'] = $range;
	}
	
	$params = array_merge($perms, $filter, $order);
	$items = $itemTable->findBy($params);
	return $items;
}

?>