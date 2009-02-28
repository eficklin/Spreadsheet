<?php
/*
*  Module written/ported by Xavier Noguer <xnoguer@rezebra.com>
*
*  The majority of this is _NOT_ my code.  I simply ported it from the
*  PERL Spreadsheet::WriteExcel module.
*
*  The author of the Spreadsheet::WriteExcel module is John McNamara
*  <jmcnamara@cpan.org>
*
*  I _DO_ maintain this code, and John McNamara has nothing to do with the
*  porting of this code to PHP.  Any questions directly related to this
*  class library should be directed to me.
*
*  License Information:
*
*	PHPExcel_Writer_Excel5_Writer:  A library for generating Excel Spreadsheets
*	Copyright (c) 2002-2003 Xavier Noguer xnoguer@rezebra.com
*
*	This library is free software; you can redistribute it and/or
*	modify it under the terms of the GNU Lesser General Public
*	License as published by the Free Software Foundation; either
*	version 2.1 of the License, or (at your option) any later version.
*
*	This library is distributed in the hope that it will be useful,
*	but WITHOUT ANY WARRANTY; without even the implied warranty of
*	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
*	Lesser General Public License for more details.
*
*	You should have received a copy of the GNU Lesser General Public
*	License along with this library; if not, write to the Free Software
*	Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

require_once 'PHPExcel/Writer/Excel5/Parser.php';
require_once 'PHPExcel/Writer/Excel5/BIFFwriter.php';
require_once 'PHPExcel/Shared/String.php';

/**
* Class for generating Excel Spreadsheets
*
* @author   Xavier Noguer <xnoguer@rezebra.com>
* @category PHPExcel
* @package  PHPExcel_Writer_Excel5
*/

class PHPExcel_Writer_Excel5_Worksheet extends PHPExcel_Writer_Excel5_BIFFwriter
{
	/**
	* Name of the Worksheet
	* @var string
	*/
	var $name;

	/**
	* Index for the Worksheet
	* @var integer
	*/
	var $index;

	/**
	* Reference to the (default) Format object for URLs
	* @var object Format
	*/
	var $_url_format;

	/**
	* Reference to the parser used for parsing formulas
	* @var object Format
	*/
	var $_parser;

	/**
	* Filehandle to the temporary file for storing data
	* @var resource
	*/
	var $_filehandle;

	/**
	* Boolean indicating if we are using a temporary file for storing data
	* @var bool
	*/
	var $_using_tmpfile;

	/**
	* Maximum number of rows for an Excel spreadsheet (BIFF5)
	* @var integer
	*/
	var $_xls_rowmax;

	/**
	* Maximum number of columns for an Excel spreadsheet (BIFF5)
	* @var integer
	*/
	var $_xls_colmax;

	/**
	* Maximum number of characters for a string (LABEL record in BIFF5)
	* @var integer
	*/
	var $_xls_strmax;

	/**
	* First row for the DIMENSIONS record
	* @var integer
	* @see _storeDimensions()
	*/
	var $_dim_rowmin;

	/**
	* Last row for the DIMENSIONS record
	* @var integer
	* @see _storeDimensions()
	*/
	var $_dim_rowmax;

	/**
	* First column for the DIMENSIONS record
	* @var integer
	* @see _storeDimensions()
	*/
	var $_dim_colmin;

	/**
	* Last column for the DIMENSIONS record
	* @var integer
	* @see _storeDimensions()
	*/
	var $_dim_colmax;

	/**
	* Array containing format information for columns
	* @var array
	*/
	var $_colinfo;

	/**
	* Array containing the selected area for the worksheet
	* @var array
	*/
	var $_selection;

	/**
	* The active pane for the worksheet
	* @var integer
	*/
	var $_active_pane;

	/**
	* Bit specifying if the worksheet is selected
	* @var integer
	*/
	var $selected;

	/**
	* Whether to use outline.
	* @var integer
	*/
	var $_outline_on;

	/**
	* Auto outline styles.
	* @var bool
	*/
	var $_outline_style;

	/**
	* Whether to have outline summary below.
	* @var bool
	*/
	var $_outline_below;

	/**
	* Whether to have outline summary at the right.
	* @var bool
	*/
	var $_outline_right;

	/**
	* Outline row level.
	* @var integer
	*/
	var $_outline_row_level;

	/**
	* Reference to the total number of strings in the workbook
	* @var integer
	*/
	var $_str_total;

	/**
	* Reference to the number of unique strings in the workbook
	* @var integer
	*/
	var $_str_unique;

	/**
	* Reference to the array containing all the unique strings in the workbook
	* @var array
	*/
	var $_str_table;

    /**
    * The temporary dir for storing files
    * @var string
    */
    var $_tmp_dir;

	/**
	* List of temporary files created
	* @var array
	*/
	var $_tempFilesCreated = array();

	/**
	 * Index of first used row (at least 0)
	 * @var int
	 */
	private $_firstRowIndex;

	/**
	 * Index of last used row. (no used rows means -1)
	 * @var int
	 */
	private $_lastRowIndex;

	/**
	 * Index of first used column (at least 0)
	 * @var int
	 */
	private $_firstColumnIndex;

	/**
	 * Index of last used column (no used columns means -1)
	 * @var int
	 */
	private $_lastColumnIndex;

	/**
	 * Sheet object
	 * @var PHPExcel_Worksheet
	 */
	private $_phpSheet;

	/**
	* Constructor
	*
	* @param string  $name		 The name of the new worksheet
	* @param integer $index		The index of the new worksheet
	* @param mixed   &$activesheet The current activesheet of the workbook we belong to
	* @param mixed   &$firstsheet  The first worksheet in the workbook we belong to
	* @param mixed   &$url_format  The default format for hyperlinks
	* @param mixed   &$parser	  The formula parser created for the Workbook
	* @param string   $tempDir	  The temporary directory to be used
	* @param PHPExcel_Worksheet $phpSheet
	* @access private
	*/
	function PHPExcel_Writer_Excel5_Worksheet($BIFF_version, $name,
												$index, &$activesheet,
												&$firstsheet, &$str_total,
												&$str_unique, &$str_table,
												&$url_format, &$parser, $tempDir = '', $phpSheet)
	{
		// It needs to call its parent's constructor explicitly
		$this->PHPExcel_Writer_Excel5_BIFFwriter();
		$this->_BIFF_version	= $BIFF_version;
		$rowmax					= 65536; // 16384 in Excel 5
		$colmax					= 256;

		$this->name				= $name;
		$this->index			= $index;
		$this->activesheet		= &$activesheet;
		$this->firstsheet		= &$firstsheet;
		$this->_str_total		= &$str_total;
		$this->_str_unique		= &$str_unique;
		$this->_str_table		= &$str_table;
		$this->_url_format		= &$url_format;
		$this->_parser			= &$parser;
		
		$this->_phpSheet = $phpSheet;

		//$this->ext_sheets		= array();
		$this->_filehandle		= '';
		$this->_using_tmpfile	= true;
		//$this->fileclosed		= 0;
		//$this->offset			= 0;
		$this->_xls_rowmax		= $rowmax;
		$this->_xls_colmax		= $colmax;
		$this->_xls_strmax		= 255;
		$this->_dim_rowmin		= $rowmax + 1;
		$this->_dim_rowmax		= 0;
		$this->_dim_colmin		= $colmax + 1;
		$this->_dim_colmax		= 0;
		$this->_colinfo			= array();
		$this->_selection		= array(0,0,0,0);
		$this->_active_pane		= 3;
		$this->selected			= 0;

		$this->_print_headers		= 0;

		$this->col_sizes		= array();
		$this->_row_sizes		= array();

		$this->_outline_row_level	= 0;
		$this->_outline_style		= 0;
		$this->_outline_below		= 1;
		$this->_outline_right		= 1;
		$this->_outline_on			= 1;

		$this->_dv				= array();

		$this->_tmp_dir			= $tempDir;

		$this->_initialize();
	}

	/**
	 * Cleanup
	 */
	public function cleanup() {
		@fclose($this->_filehandle);

		foreach ($this->_tempFilesCreated as $file) {
			@unlink($file);
		}
	}

	/**
	* Open a tmp file to store the majority of the Worksheet data. If this fails,
	* for example due to write permissions, store the data in memory. This can be
	* slow for large files.
	*
	* @access private
	*/
	function _initialize()
	{
		// Open tmp file for storing Worksheet data
		$fileName = tempnam($this->_tmp_dir, 'XLSHEET');
		$fh = fopen($fileName, 'w+');
		if ($fh) {
			// Store filehandle
			$this->_filehandle = $fh;
			$this->_tempFilesCreated[] = $fileName;
		} else {
			// If tmpfile() fails store data in memory
			$this->_using_tmpfile = false;
		}
	}

    /**
    * Sets the temp dir used for storing files
    *
    * @access public
    * @param string $dir The dir to be used as temp dir
    * @return true if given dir is valid, false otherwise
    */
    function setTempDir($dir)
    {
        if (is_dir($dir)) {
            $this->_tmp_dir = $dir;
            return true;
        }
        return false;
    }

	/**
	* Add data to the beginning of the workbook (note the reverse order)
	* and to the end of the workbook.
	*
	* @access public
	* @see PHPExcel_Writer_Excel5_Workbook::storeWorkbook()
	* @param array $sheetnames The array of sheetnames from the Workbook this
	*						  worksheet belongs to
	*/
	function close($sheetnames)
	{
		$num_sheets = count($sheetnames);

		/***********************************************
		* Prepend in reverse order!!
		*/

		// Prepend the sheet dimensions
		$this->_storeDimensions();

		// Prepend the sheet password
		$this->_storePassword();

		// Prepend the sheet protection
		$this->_storeProtect();

		// Prepend the page setup
		$this->_storeSetup();

		/* FIXME: margins are actually appended */
		// Prepend the bottom margin
		$this->_storeMarginBottom();

		// Prepend the top margin
		$this->_storeMarginTop();

		// Prepend the right margin
		$this->_storeMarginRight();

		// Prepend the left margin
		$this->_storeMarginLeft();

		// Prepend the page vertical centering
		$this->_storeVcenter();

		// Prepend the page horizontal centering
		$this->_storeHcenter();

		// Prepend the page footer
		$this->_storeFooter();

		// Prepend the page header
		$this->_storeHeader();

		// Prepend the vertical and horizontal page breaks
		$this->_storeBreaks();

		// Prepend WSBOOL
		$this->_storeWsbool();

		// Prepend DEFAULTROWHEIGHT
		if ($this->_BIFF_version == 0x0600) {
			$this->_storeDefaultRowHeight();
		}

		// Prepend GRIDSET
		$this->_storeGridset();

		 //  Prepend GUTS
		$this->_storeGuts();

		// Prepend PRINTGRIDLINES
		$this->_storePrintGridlines();

		// Prepend PRINTHEADERS
		$this->_storePrintHeaders();

		// Prepend EXTERNSHEET references
		if ($this->_BIFF_version == 0x0500) {
			for ($i = $num_sheets; $i > 0; --$i) {
				$sheetname = $sheetnames[$i-1];
				$this->_storeExternsheet($sheetname);
			}
		}

		// Prepend the EXTERNCOUNT of external references.
		if ($this->_BIFF_version == 0x0500) {
			$this->_storeExterncount($num_sheets);
		}

		// Prepend the COLINFO records if they exist
		if (!empty($this->_colinfo)) {
			$colcount = count($this->_colinfo);
			for ($i = 0; $i < $colcount; ++$i) {
				$this->_storeColinfo($this->_colinfo[$i]);
			}
		}

		// Prepend the DEFCOLWIDTH record
		$this->_storeDefcol();

		// Prepend the BOF record
		$this->_storeBof(0x0010);

		/*
		* End of prepend. Read upwards from here.
		***********************************************/

		// Append
		$this->_storeWindow2();
		$this->_storeZoom();
		if ($this->_phpSheet->getFreezePane()) {
			$this->_storePanes();
		}
		$this->_storeSelection($this->_selection);
		$this->_storeMergedCells();
		/* TODO: add data validity */
		/*if ($this->_BIFF_version == 0x0600) {
			$this->_storeDataValidity();
		}*/

		if ($this->_BIFF_version == 0x0600) {
			$this->_storeRangeProtection();
		}

		$this->_storeEof();
	}

	/**
	 * Write a cell range address in BIFF8
	 * always fixed range
	 * See section 2.5.14 in OpenOffice.org's Documentation of the Microsoft Excel File Format
	 *
	 * @param string $range E.g. 'A1' or 'A1:B6'
	 * @return string Binary data
	 */
	private function _writeBIFF8CellRangeAddressFixed($range = 'A1')
	{
		$explodes = explode(':', $range);

		// extract first cell, e.g. 'A1'
		$firstCell = $explodes[0];

		// extract last cell, e.g. 'B6'
		if (count($explodes) == 1) {
			$lastCell = $firstCell;
		} else {
			$lastCell = $explodes[1];
		}

		$firstCellCoordinates = PHPExcel_Cell::coordinateFromString($firstCell); // e.g. array(0, 1)
		$lastCellCoordinates  = PHPExcel_Cell::coordinateFromString($lastCell);  // e.g. array(1, 6)

		$data = pack('vvvv',
			$firstCellCoordinates[1] - 1,
			$lastCellCoordinates[1] - 1,
			PHPExcel_Cell::columnIndexFromString($firstCellCoordinates[0]) - 1,
			PHPExcel_Cell::columnIndexFromString($lastCellCoordinates[0]) - 1
		);

		return $data;
	}

	/**
	* Retrieve the worksheet name.
	* This is usefull when creating worksheets without a name.
	*
	* @access public
	* @return string The worksheet's name
	*/
	function getName()
	{
		return $this->name;
	}

	/**
	* Retrieves data from memory in one chunk, or from disk in $buffer
	* sized chunks.
	*
	* @return string The data
	*/
	function getData()
	{
		$buffer = 4096;

		// Return data stored in memory
		if (isset($this->_data)) {
			$tmp   = $this->_data;
			unset($this->_data);
			$fh	= $this->_filehandle;
			if ($this->_using_tmpfile) {
				fseek($fh, 0);
			}
			return $tmp;
		}
		// Return data stored on disk
		if ($this->_using_tmpfile) {
			if ($tmp = fread($this->_filehandle, $buffer)) {
				return $tmp;
			}
		}

		// No data to return
		return '';
	}

	/**
	* Set this worksheet as a selected worksheet,
	* i.e. the worksheet has its tab highlighted.
	*
	* @access public
	*/
	function select()
	{
		$this->selected = 1;
	}

	/**
	* Set this worksheet as the active worksheet,
	* i.e. the worksheet that is displayed when the workbook is opened.
	* Also set it as selected.
	*
	* @access public
	*/
	function activate()
	{
		$this->selected = 1;
		$this->activesheet = $this->index;
	}

	/**
	* Set this worksheet as the first visible sheet.
	* This is necessary when there are a large number of worksheets and the
	* activated worksheet is not visible on the screen.
	*
	* @access public
	*/
	function setFirstSheet()
	{
		$this->firstsheet = $this->index;
	}

	/**
	* Set the width of a single column or a range of columns.
	*
	* @access public
	* @param integer $firstcol first column on the range
	* @param integer $lastcol  last column on the range
	* @param integer $width	width to set
	* @param mixed   $format   The optional XF format to apply to the columns
	* @param integer $hidden   The optional hidden atribute
	* @param integer $level	The optional outline level
	*/
	function setColumn($firstcol, $lastcol, $width, $format = null, $hidden = 0, $level = 0)
	{
		$this->_colinfo[] = array($firstcol, $lastcol, $width, &$format, $hidden, $level);

		// Set width to zero if column is hidden
		$width = ($hidden) ? 0 : $width;

		for ($col = $firstcol; $col <= $lastcol; ++$col) {
			$this->col_sizes[$col] = $width;
		}
	}

	/**
	* Set which cell or cells are selected in a worksheet
	*
	* @access public
	* @param integer $first_row	first row in the selected quadrant
	* @param integer $first_column first column in the selected quadrant
	* @param integer $last_row	 last row in the selected quadrant
	* @param integer $last_column  last column in the selected quadrant
	*/
	function setSelection($first_row,$first_column,$last_row,$last_column)
	{
		$this->_selection = array($first_row,$first_column,$last_row,$last_column);
	}

	/**
	* Set the option to print the row and column headers on the printed page.
	*
	* @access public
	* @param integer $print Whether to print the headers or not. Defaults to 1 (print).
	*/
	function printRowColHeaders($print = 1)
	{
		$this->_print_headers = $print;
	}

	/**
	* Map to the appropriate write method acording to the token recieved.
	*
	* @access public
	* @param integer $row	The row of the cell we are writing to
	* @param integer $col	The column of the cell we are writing to
	* @param mixed   $token  What we are writing
	* @param mixed   $format The optional format to apply to the cell
	*/
	function write($row, $col, $token, $format = null, $numberFormat = null)
	{
		if (($numberFormat != 'General') && (PHPExcel_Shared_Date::isDateTimeFormatCode($numberFormat))) {
			if (is_string($token)) {
				//	Error string
				return $this->writeString($row, $col, $token, $format);
			} elseif (!is_float($token)) {
				//	PHP serialized date/time or date/time object
				return $this->writeNumber($row, $col, PHPExcel_Shared_Date::PHPToExcel($token), $format);
			} else {
				//	Excel serialized date/time
				return $this->writeNumber($row, $col, $token, $format);
			}
		} elseif (preg_match("/^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/", $token)) {
			// Match number
			return $this->writeNumber($row, $col, $token, $format);
		} elseif ($token == '') {
			// Match blank
			return $this->writeBlank($row, $col, $format);
		} else {
			// Default: match string
			return $this->writeString($row, $col, $token, $format);
		}
	}

	/**
	* Returns an index to the XF record in the workbook
	*
	* @access private
	* @param mixed &$format The optional XF format
	* @return integer The XF record index
	*/
	function _XF(&$format)
	{
		if ($format) {
			return($format->getXfIndex());
		} else {
			return(0x0F);
		}
	}


	/******************************************************************************
	*******************************************************************************
	*
	* Internal methods
	*/


	/**
	* Store Worksheet data in memory using the parent's class append() or to a
	* temporary file, the default.
	*
	* @access private
	* @param string $data The binary data to append
	*/
	function _append($data)
	{
		if ($this->_using_tmpfile) {
			// Add CONTINUE records if necessary
			if (strlen($data) > $this->_limit) {
				$data = $this->_addContinue($data);
			}
			fwrite($this->_filehandle, $data);
			$this->_datasize += strlen($data);
		} else {
			parent::_append($data);
		}
	}

	/**
	* This method sets the properties for outlining and grouping. The defaults
	* correspond to Excel's defaults.
	*
	* @param bool $visible
	* @param bool $symbols_below
	* @param bool $symbols_right
	* @param bool $auto_style
	*/
	function setOutline($visible = true, $symbols_below = true, $symbols_right = true, $auto_style = false)
	{
		$this->_outline_on	= $visible;
		$this->_outline_below = $symbols_below;
		$this->_outline_right = $symbols_right;
		$this->_outline_style = $auto_style;

		// Ensure this is a boolean vale for Window2
		if ($this->_outline_on) {
			$this->_outline_on = 1;
		}
	 }

	/******************************************************************************
	*******************************************************************************
	*
	* BIFF RECORDS
	*/


	/**
	* Write a double to the specified row and column (zero indexed).
	* An integer can be written as a double. Excel will display an
	* integer. $format is optional.
	*
	* Returns  0 : normal termination
	*		 -2 : row or column out of range
	*
	* @access public
	* @param integer $row	Zero indexed row
	* @param integer $col	Zero indexed column
	* @param float   $num	The number to write
	* @param mixed   $format The optional XF format
	* @return integer
	*/
	function writeNumber($row, $col, $num, $format = null)
	{
		$record	= 0x0203;				 // Record identifier
		$length	= 0x000E;				 // Number of bytes to follow

		$xf		= $this->_XF($format);	// The cell format

		// Check that row and col are valid and store max and min values
		if ($row >= $this->_xls_rowmax) {
			return(-2);
		}
		if ($col >= $this->_xls_colmax) {
			return(-2);
		}
		if ($row <  $this->_dim_rowmin)  {
			$this->_dim_rowmin = $row;
		}
		if ($row >  $this->_dim_rowmax)  {
			$this->_dim_rowmax = $row;
		}
		if ($col <  $this->_dim_colmin)  {
			$this->_dim_colmin = $col;
		}
		if ($col >  $this->_dim_colmax)  {
			$this->_dim_colmax = $col;
		}

		$header	= pack("vv",  $record, $length);
		$data	  = pack("vvv", $row, $col, $xf);
		$xl_double = pack("d",   $num);
		if ($this->_byte_order) { // if it's Big Endian
			$xl_double = strrev($xl_double);
		}

		$this->_append($header.$data.$xl_double);
		return(0);
	}

	/**
	* Write a string to the specified row and column (zero indexed).
	* NOTE: there is an Excel 5 defined limit of 255 characters.
	* $format is optional.
	* Returns  0 : normal termination
	*		 -2 : row or column out of range
	*		 -3 : long string truncated to 255 chars
	*
	* @access public
	* @param integer $row	Zero indexed row
	* @param integer $col	Zero indexed column
	* @param string  $str	The string to write
	* @param mixed   $format The XF format for the cell
	* @return integer
	*/
	function writeString($row, $col, $str, $format = null)
	{
		if ($this->_BIFF_version == 0x0600) {
			return $this->writeStringBIFF8($row, $col, $str, $format);
		}
		$strlen	= strlen($str);
		$record	= 0x0204;				   // Record identifier
		$length	= 0x0008 + $strlen;		 // Bytes to follow
		$xf		= $this->_XF($format);	  // The cell format

		$str_error = 0;

		// Check that row and col are valid and store max and min values
		if ($row >= $this->_xls_rowmax) {
			return(-2);
		}
		if ($col >= $this->_xls_colmax) {
			return(-2);
		}
		if ($row <  $this->_dim_rowmin) {
			$this->_dim_rowmin = $row;
		}
		if ($row >  $this->_dim_rowmax) {
			$this->_dim_rowmax = $row;
		}
		if ($col <  $this->_dim_colmin) {
			$this->_dim_colmin = $col;
		}
		if ($col >  $this->_dim_colmax) {
			$this->_dim_colmax = $col;
		}

		if ($strlen > $this->_xls_strmax) { // LABEL must be < 255 chars
			$str	   = substr($str, 0, $this->_xls_strmax);
			$length	= 0x0008 + $this->_xls_strmax;
			$strlen	= $this->_xls_strmax;
			$str_error = -3;
		}

		$header	= pack("vv",   $record, $length);
		$data	  = pack("vvvv", $row, $col, $xf, $strlen);
		$this->_append($header . $data . $str);
		return($str_error);
	}

	/**
	* Write a string to the specified row and column (zero indexed).
	* This is the BIFF8 version (no 255 chars limit).
	* $format is optional.
	* Returns  0 : normal termination
	*		 -2 : row or column out of range
	*		 -3 : long string truncated to 255 chars
	*
	* @access public
	* @param integer $row	Zero indexed row
	* @param integer $col	Zero indexed column
	* @param string  $str	The string to write
	* @param mixed   $format The XF format for the cell
	* @return integer
	*/
	function writeStringBIFF8($row, $col, $str, $format = null)
	{
		$str = iconv('UTF-8', 'UTF-16LE', $str);
		$strlen = function_exists('mb_strlen') ? mb_strlen($str, 'UTF-16LE') : (strlen($str) / 2);
		$encoding  = 0x1;
		
		$record	= 0x00FD;				   // Record identifier
		$length	= 0x000A;				   // Bytes to follow
		$xf		= $this->_XF($format);	  // The cell format

		$str_error = 0;

		// Check that row and col are valid and store max and min values
		if ($this->_checkRowCol($row, $col) == false) {
			return -2;
		}

		$str = pack('vC', $strlen, $encoding).$str;

		/* check if string is already present */
		if (!isset($this->_str_table[$str])) {
			$this->_str_table[$str] = $this->_str_unique++;
		}
		$this->_str_total++;

		$header	= pack('vv',   $record, $length);
		$data	  = pack('vvvV', $row, $col, $xf, $this->_str_table[$str]);
		$this->_append($header.$data);
		return $str_error;
	}

	/**
	* Check row and col before writing to a cell, and update the sheet's
	* dimensions accordingly
	*
	* @access private
	* @param integer $row	Zero indexed row
	* @param integer $col	Zero indexed column
	* @return boolean true for success, false if row and/or col are grester
	*				 then maximums allowed.
	*/
	function _checkRowCol($row, $col)
	{
		if ($row >= $this->_xls_rowmax) {
			return false;
		}
		if ($col >= $this->_xls_colmax) {
			return false;
		}
		if ($row <  $this->_dim_rowmin) {
			$this->_dim_rowmin = $row;
		}
		if ($row >  $this->_dim_rowmax) {
			$this->_dim_rowmax = $row;
		}
		if ($col <  $this->_dim_colmin) {
			$this->_dim_colmin = $col;
		}
		if ($col >  $this->_dim_colmax) {
			$this->_dim_colmax = $col;
		}
		return true;
	}

	/**
	* Writes a note associated with the cell given by the row and column.
	* NOTE records don't have a length limit.
	*
	* @access public
	* @param integer $row	Zero indexed row
	* @param integer $col	Zero indexed column
	* @param string  $note   The note to write
	*/
	function writeNote($row, $col, $note)
	{
		$note_length	= strlen($note);
		$record		 = 0x001C;				// Record identifier
		$max_length	 = 2048;				  // Maximun length for a NOTE record
		//$length	  = 0x0006 + $note_length;	// Bytes to follow

		// Check that row and col are valid and store max and min values
		if ($row >= $this->_xls_rowmax) {
			return(-2);
		}
		if ($col >= $this->_xls_colmax) {
			return(-2);
		}
		if ($row <  $this->_dim_rowmin) {
			$this->_dim_rowmin = $row;
		}
		if ($row >  $this->_dim_rowmax) {
			$this->_dim_rowmax = $row;
		}
		if ($col <  $this->_dim_colmin) {
			$this->_dim_colmin = $col;
		}
		if ($col >  $this->_dim_colmax) {
			$this->_dim_colmax = $col;
		}

		// Length for this record is no more than 2048 + 6
		$length	= 0x0006 + min($note_length, 2048);
		$header	= pack("vv",   $record, $length);
		$data	  = pack("vvv", $row, $col, $note_length);
		$this->_append($header . $data . substr($note, 0, 2048));

		for ($i = $max_length; $i < $note_length; $i += $max_length) {
			$chunk  = substr($note, $i, $max_length);
			$length = 0x0006 + strlen($chunk);
			$header = pack("vv",   $record, $length);
			$data   = pack("vvv", -1, 0, strlen($chunk));
			$this->_append($header.$data.$chunk);
		}
		return(0);
	}

	/**
	* Write a blank cell to the specified row and column (zero indexed).
	* A blank cell is used to specify formatting without adding a string
	* or a number.
	*
	* A blank cell without a format serves no purpose. Therefore, we don't write
	* a BLANK record unless a format is specified.
	*
	* Returns  0 : normal termination (including no format)
	*		 -1 : insufficient number of arguments
	*		 -2 : row or column out of range
	*
	* @access public
	* @param integer $row	Zero indexed row
	* @param integer $col	Zero indexed column
	* @param mixed   $format The XF format
	*/
	function writeBlank($row, $col, $format)
	{
		// Don't write a blank cell unless it has a format
		if (!$format) {
			return(0);
		}

		$record	= 0x0201;				 // Record identifier
		$length	= 0x0006;				 // Number of bytes to follow
		$xf		= $this->_XF($format);	// The cell format

		// Check that row and col are valid and store max and min values
		if ($row >= $this->_xls_rowmax) {
			return(-2);
		}
		if ($col >= $this->_xls_colmax) {
			return(-2);
		}
		if ($row <  $this->_dim_rowmin) {
			$this->_dim_rowmin = $row;
		}
		if ($row >  $this->_dim_rowmax) {
			$this->_dim_rowmax = $row;
		}
		if ($col <  $this->_dim_colmin) {
			$this->_dim_colmin = $col;
		}
		if ($col >  $this->_dim_colmax) {
			$this->_dim_colmax = $col;
		}

		$header	= pack("vv",  $record, $length);
		$data	  = pack("vvv", $row, $col, $xf);
		$this->_append($header . $data);
		return 0;
	}

	/**
	 * Write a boolean or an error type to the specified row and column (zero indexed)
	 */
	public function writeBoolErr($row, $col, $value, $isError, $format)
	{
		$record = 0x0205;
		$length = 8;
		$xf = $this->_XF($format);

		// Check that row and col are valid and store max and min values
		if ($row >= $this->_xls_rowmax) {
			return(-2);
		}
		if ($col >= $this->_xls_colmax) {
			return(-2);
		}
		if ($row <  $this->_dim_rowmin)  {
			$this->_dim_rowmin = $row;
		}
		if ($row >  $this->_dim_rowmax)  {
			$this->_dim_rowmax = $row;
		}
		if ($col <  $this->_dim_colmin)  {
			$this->_dim_colmin = $col;
		}
		if ($col >  $this->_dim_colmax)  {
			$this->_dim_colmax = $col;
		}

		$header	= pack("vv",  $record, $length);
		$data	  = pack("vvvCC", $row, $col, $xf, $value, $isError);
		$this->_append($header . $data);
		return 0;
	}

	/**
	* Write a formula to the specified row and column (zero indexed).
	* The textual representation of the formula is passed to the parser in
	* Parser.php which returns a packed binary string.
	*
	* Returns  0 : normal termination
	*		 -1 : formula errors (bad formula)
	*		 -2 : row or column out of range
	*
	* @access public
	* @param integer $row	 Zero indexed row
	* @param integer $col	 Zero indexed column
	* @param string  $formula The formula text string
	* @param mixed   $format  The optional XF format
	* @return integer
	*/
	function writeFormula($row, $col, $formula, $format = null)
	{
		$record	= 0x0006;	 // Record identifier

		// Excel normally stores the last calculated value of the formula in $num.
		// Clearly we are not in a position to calculate this a priori. Instead
		// we set $num to zero and set the option flags in $grbit to ensure
		// automatic calculation of the formula when the file is opened.
		//
		$xf		= $this->_XF($format); // The cell format
		$num	   = 0x00;				// Current value of formula
		$grbit	 = 0x03;				// Option flags
		$unknown   = 0x0000;			  // Must be zero


		// Check that row and col are valid and store max and min values
		if ($this->_checkRowCol($row, $col) == false) {
			return -2;
		}

		// Strip the '=' or '@' sign at the beginning of the formula string
		if (preg_match("/^=/", $formula)) {
			$formula = preg_replace("/(^=)/", "", $formula);
		} elseif (preg_match("/^@/", $formula)) {
			$formula = preg_replace("/(^@)/", "", $formula);
		} else {
			// Error handling
			$this->writeString($row, $col, 'Unrecognised character for formula');
			return -1;
		}

		// Parse the formula using the parser in Parser.php
		$error = $this->_parser->parse($formula);

		$formula = $this->_parser->toReversePolish();

		$formlen	= strlen($formula);	// Length of the binary string
		$length	 = 0x16 + $formlen;	 // Length of the record data

		$header	= pack("vv",	  $record, $length);
		$data	  = pack("vvvdvVv", $row, $col, $xf, $num,
									 $grbit, $unknown, $formlen);

		$this->_append($header . $data . $formula);
		return 0;
	}

	/**
	* Write a hyperlink.
	* This is comprised of two elements: the visible label and
	* the invisible link. The visible label is the same as the link unless an
	* alternative string is specified. The label is written using the
	* writeString() method. Therefore the 255 characters string limit applies.
	* $string and $format are optional.
	*
	* The hyperlink can be to a http, ftp, mail, internal sheet (not yet), or external
	* directory url.
	*
	* Returns  0 : normal termination
	*		 -2 : row or column out of range
	*		 -3 : long string truncated to 255 chars
	*
	* @access public
	* @param integer $row	Row
	* @param integer $col	Column
	* @param string  $url	URL string
	* @return integer
	*/
	function writeUrl($row, $col, $url)
	{
		// Add start row and col to arg list
		return($this->_writeUrlRange($row, $col, $row, $col, $url));
	}

	/**
	* This is the more general form of writeUrl(). It allows a hyperlink to be
	* written to a range of cells. This function also decides the type of hyperlink
	* to be written. These are either, Web (http, ftp, mailto), Internal
	* (Sheet1!A1) or external ('c:\temp\foo.xls#Sheet1!A1').
	*
	* @access private
	* @see writeUrl()
	* @param integer $row1   Start row
	* @param integer $col1   Start column
	* @param integer $row2   End row
	* @param integer $col2   End column
	* @param string  $url	URL string
	* @return integer
	*/

	function _writeUrlRange($row1, $col1, $row2, $col2, $url)
	{

		// Check for internal/external sheet links or default to web link
		if (preg_match('[^internal:]', $url)) {
			return($this->_writeUrlInternal($row1, $col1, $row2, $col2, $url));
		}
		if (preg_match('[^external:]', $url)) {
			return($this->_writeUrlExternal($row1, $col1, $row2, $col2, $url));
		}
		return($this->_writeUrlWeb($row1, $col1, $row2, $col2, $url));
	}


	/**
	* Used to write http, ftp and mailto hyperlinks.
	* The link type ($options) is 0x03 is the same as absolute dir ref without
	* sheet. However it is differentiated by the $unknown2 data stream.
	*
	* @access private
	* @see writeUrl()
	* @param integer $row1   Start row
	* @param integer $col1   Start column
	* @param integer $row2   End row
	* @param integer $col2   End column
	* @param string  $url	URL string
	* @return integer
	*/
	function _writeUrlWeb($row1, $col1, $row2, $col2, $url)
	{
		$record	  = 0x01B8;					   // Record identifier
		$length	  = 0x00000;					  // Bytes to follow

		// Pack the undocumented parts of the hyperlink stream
		$unknown1	= pack("H*", "D0C9EA79F9BACE118C8200AA004BA90B02000000");
		$unknown2	= pack("H*", "E0C9EA79F9BACE118C8200AA004BA90B");

		// Pack the option flags
		$options	 = pack("V", 0x03);

		// Convert URL to a null terminated wchar string
		$url		 = join("\0", preg_split("''", $url, -1, PREG_SPLIT_NO_EMPTY));
		$url		 = $url . "\0\0\0";

		// Pack the length of the URL
		$url_len	 = pack("V", strlen($url));

		// Calculate the data length
		$length	  = 0x34 + strlen($url);

		// Pack the header data
		$header	  = pack("vv",   $record, $length);
		$data		= pack("vvvv", $row1, $row2, $col1, $col2);

		// Write the packed data
		$this->_append($header . $data .
					   $unknown1 . $options .
					   $unknown2 . $url_len . $url);
		return 0;
	}

	/**
	* Used to write internal reference hyperlinks such as "Sheet1!A1".
	*
	* @access private
	* @see writeUrl()
	* @param integer $row1   Start row
	* @param integer $col1   Start column
	* @param integer $row2   End row
	* @param integer $col2   End column
	* @param string  $url	URL string
	* @return integer
	*/
	function _writeUrlInternal($row1, $col1, $row2, $col2, $url)
	{
		$record	  = 0x01B8;					   // Record identifier
		$length	  = 0x00000;					  // Bytes to follow

		// Strip URL type
		$url = preg_replace('/^internal:/', '', $url);

		// Pack the undocumented parts of the hyperlink stream
		$unknown1	= pack("H*", "D0C9EA79F9BACE118C8200AA004BA90B02000000");

		// Pack the option flags
		$options	 = pack("V", 0x08);

		// Convert the URL type and to a null terminated wchar string
		$url		 = join("\0", preg_split("''", $url, -1, PREG_SPLIT_NO_EMPTY));
		$url		 = $url . "\0\0\0";

		// Pack the length of the URL as chars (not wchars)
		$url_len	 = pack("V", floor(strlen($url)/2));

		// Calculate the data length
		$length	  = 0x24 + strlen($url);

		// Pack the header data
		$header	  = pack("vv",   $record, $length);
		$data		= pack("vvvv", $row1, $row2, $col1, $col2);

		// Write the packed data
		$this->_append($header . $data .
					   $unknown1 . $options .
					   $url_len . $url);
		return 0;
	}

	/**
	* Write links to external directory names such as 'c:\foo.xls',
	* c:\foo.xls#Sheet1!A1', '../../foo.xls'. and '../../foo.xls#Sheet1!A1'.
	*
	* Note: Excel writes some relative links with the $dir_long string. We ignore
	* these cases for the sake of simpler code.
	*
	* @access private
	* @see writeUrl()
	* @param integer $row1   Start row
	* @param integer $col1   Start column
	* @param integer $row2   End row
	* @param integer $col2   End column
	* @param string  $url	URL string
	* @return integer
	*/
	function _writeUrlExternal($row1, $col1, $row2, $col2, $url)
	{
		// Network drives are different. We will handle them separately
		// MS/Novell network drives and shares start with \\
		if (preg_match('[^external:\\\\]', $url)) {
			return; //($this->_writeUrlExternal_net($row1, $col1, $row2, $col2, $url, $str, $format));
		}

		$record	  = 0x01B8;					   // Record identifier
		$length	  = 0x00000;					  // Bytes to follow

		// Strip URL type and change Unix dir separator to Dos style (if needed)
		//
		$url = preg_replace('/^external:/', '', $url);
		$url = preg_replace('/\//', "\\", $url);

		// Determine if the link is relative or absolute:
		//   relative if link contains no dir separator, "somefile.xls"
		//   relative if link starts with up-dir, "..\..\somefile.xls"
		//   otherwise, absolute

		$absolute	= 0x02; // Bit mask
		if (!preg_match("/\\\/", $url)) {
			$absolute	= 0x00;
		}
		if (preg_match("/^\.\.\\\/", $url)) {
			$absolute	= 0x00;
		}
		$link_type			   = 0x01 | $absolute;

		// Determine if the link contains a sheet reference and change some of the
		// parameters accordingly.
		// Split the dir name and sheet name (if it exists)
		/*if (preg_match("/\#/", $url)) {
			list($dir_long, $sheet) = split("\#", $url);
		} else {
			$dir_long = $url;
		}

		if (isset($sheet)) {
			$link_type |= 0x08;
			$sheet_len  = pack("V", strlen($sheet) + 0x01);
			$sheet	  = join("\0", split('', $sheet));
			$sheet	 .= "\0\0\0";
		} else {
			$sheet_len   = '';
			$sheet	   = '';
		}*/
		$dir_long = $url;
		if (preg_match("/\#/", $url)) {
			$link_type |= 0x08;
		}



		// Pack the link type
		$link_type   = pack("V", $link_type);

		// Calculate the up-level dir count e.g.. (..\..\..\ == 3)
		$up_count	= preg_match_all("/\.\.\\\/", $dir_long, $useless);
		$up_count	= pack("v", $up_count);

		// Store the short dos dir name (null terminated)
		$dir_short   = preg_replace("/\.\.\\\/", '', $dir_long) . "\0";

		// Store the long dir name as a wchar string (non-null terminated)
		//$dir_long	   = join("\0", split('', $dir_long));
		$dir_long	   = $dir_long . "\0";

		// Pack the lengths of the dir strings
		$dir_short_len = pack("V", strlen($dir_short)	  );
		$dir_long_len  = pack("V", strlen($dir_long)	   );
		$stream_len	= pack("V", 0);//strlen($dir_long) + 0x06);

		// Pack the undocumented parts of the hyperlink stream
		$unknown1 = pack("H*",'D0C9EA79F9BACE118C8200AA004BA90B02000000'	   );
		$unknown2 = pack("H*",'0303000000000000C000000000000046'			   );
		$unknown3 = pack("H*",'FFFFADDE000000000000000000000000000000000000000');
		$unknown4 = pack("v",  0x03											);

		// Pack the main data stream
		$data		= pack("vvvv", $row1, $row2, $col1, $col2) .
						  $unknown1	 .
						  $link_type	.
						  $unknown2	 .
						  $up_count	 .
						  $dir_short_len.
						  $dir_short	.
						  $unknown3	 .
						  $stream_len   ;/*.
						  $dir_long_len .
						  $unknown4	 .
						  $dir_long	 .
						  $sheet_len	.
						  $sheet		;*/

		// Pack the header data
		$length   = strlen($data);
		$header   = pack("vv", $record, $length);

		// Write the packed data
		$this->_append($header. $data);
		return 0;
	}

	/**
	* This method is used to set the height and format for a row.
	*
	* @access public
	* @param integer $row	The row to set
	* @param integer $height Height we are giving to the row.
	*						Use null to set XF without setting height
	* @param mixed   $format XF format we are giving to the row
	* @param bool	$hidden The optional hidden attribute
	* @param integer $level  The optional outline level for row, in range [0,7]
	*/
	function setRow($row, $height, $format = null, $hidden = false, $level = 0)
	{
		$record	  = 0x0208;			   // Record identifier
		$length	  = 0x0010;			   // Number of bytes to follow

		$colMic	  = 0x0000;			   // First defined column
		$colMac	  = 0x0000;			   // Last defined column
		$irwMac	  = 0x0000;			   // Used by Excel to optimise loading
		$reserved	= 0x0000;			   // Reserved
		$grbit	   = 0x0000;			   // Option flags
		$ixfe		= $this->_XF($format);  // XF index

		if ( $height < 0 ){
			$height = null;
		}

		// set _row_sizes so _sizeRow() can use it
		$this->_row_sizes[$row] = $height;

		// Use setRow($row, null, $XF) to set XF format without setting height
		if ($height != null) {
			$miyRw = $height * 20;  // row height
		} else {
			$miyRw = 0xff;		  // default row height is 256
		}

		$level = max(0, min($level, 7));  // level should be between 0 and 7
		$this->_outline_row_level = max($level, $this->_outline_row_level);


		// Set the options flags. fUnsynced is used to show that the font and row
		// heights are not compatible. This is usually the case for WriteExcel.
		// The collapsed flag 0x10 doesn't seem to be used to indicate that a row
		// is collapsed. Instead it is used to indicate that the previous row is
		// collapsed. The zero height flag, 0x20, is used to collapse a row.

		$grbit |= $level;
		if ($hidden) {
			$grbit |= 0x0020;
		}
		$grbit |= 0x0040; // fUnsynced
		if ($format) {
			$grbit |= 0x0080;
		}
		$grbit |= 0x0100;

		$header   = pack("vv",	   $record, $length);
		$data	 = pack("vvvvvvvv", $row, $colMic, $colMac, $miyRw,
									 $irwMac,$reserved, $grbit, $ixfe);
		$this->_append($header.$data);
	}

	/**
	* Writes Excel DIMENSIONS to define the area in which there is data.
	*
	* @access private
	*/
	function _storeDimensions()
	{
		$record	= 0x0200;				 // Record identifier
		$row_min   = $this->_dim_rowmin;	 // First row
		$row_max   = $this->_dim_rowmax + 1; // Last row plus 1
		$col_min   = $this->_dim_colmin;	 // First column
		$col_max   = $this->_dim_colmax + 1; // Last column plus 1
		$reserved  = 0x0000;				 // Reserved by Excel

		if ($this->_BIFF_version == 0x0500) {
			$length	= 0x000A;			   // Number of bytes to follow
			$data	  = pack("vvvvv", $row_min, $row_max,
									   $col_min, $col_max, $reserved);
		} elseif ($this->_BIFF_version == 0x0600) {
			$length	= 0x000E;
			//$data	  = pack("VVvvv", $row_min, $row_max,
			//						   $col_min, $col_max, $reserved);
			$data = pack("VVvvv", $this->_firstRowIndex, $this->_lastRowIndex + 1,
							$this->_firstColumnIndex, $this->_lastColumnIndex + 1, $reserved);
		}
		$header = pack("vv", $record, $length);
		$this->_prepend($header.$data);
	}

	/**
	* Write BIFF record Window2.
	*
	* @access private
	*/
	function _storeWindow2()
	{
		$record		 = 0x023E;	 // Record identifier
		if ($this->_BIFF_version == 0x0500) {
			$length		 = 0x000A;	 // Number of bytes to follow
		} elseif ($this->_BIFF_version == 0x0600) {
			$length		 = 0x0012;
		}

		$grbit		  = 0x00B6;	 // Option flags
		$rwTop		  = 0x0000;	 // Top row visible in window
		$colLeft		= 0x0000;	 // Leftmost column visible in window


		// The options flags that comprise $grbit
		$fDspFmla	   = 0;					 // 0 - bit
		$fDspGrid	   = $this->_phpSheet->getShowGridlines() ? 1 : 0; // 1
		$fDspRwCol	  = 1;					 // 2
		$fFrozen		= $this->_phpSheet->getFreezePane() ? 1 : 0;		// 3
		$fDspZeros	  = 1;					 // 4
		$fDefaultHdr	= 1;					 // 5
		$fArabic		= 0;					 // 6
		$fDspGuts	   = $this->_outline_on;	// 7
		$fFrozenNoSplit = 0;					 // 0 - bit
		$fSelected	  = $this->selected;	   // 1
		$fPaged		 = 1;					 // 2

		$grbit			 = $fDspFmla;
		$grbit			|= $fDspGrid	   << 1;
		$grbit			|= $fDspRwCol	  << 2;
		$grbit			|= $fFrozen		<< 3;
		$grbit			|= $fDspZeros	  << 4;
		$grbit			|= $fDefaultHdr	<< 5;
		$grbit			|= $fArabic		<< 6;
		$grbit			|= $fDspGuts	   << 7;
		$grbit			|= $fFrozenNoSplit << 8;
		$grbit			|= $fSelected	  << 9;
		$grbit			|= $fPaged		 << 10;

		$header  = pack("vv",   $record, $length);
		$data	= pack("vvv", $grbit, $rwTop, $colLeft);
		// FIXME !!!
		if ($this->_BIFF_version == 0x0500) {
			$rgbHdr		 = 0x00000000; // Row/column heading and gridline color
			$data .= pack("V", $rgbHdr);
		} elseif ($this->_BIFF_version == 0x0600) {
			$rgbHdr	   = 0x0040; // Row/column heading and gridline color index
			$zoom_factor_page_break = 0x0000;
			$zoom_factor_normal	 = 0x0000;
			$data .= pack("vvvvV", $rgbHdr, 0x0000, $zoom_factor_page_break, $zoom_factor_normal, 0x00000000);
		}
		$this->_append($header.$data);
	}

	/**
	 * Write BIFF record DEFAULTROWHEIGHT.
	 *
	 * @access private
	 */
	private function _storeDefaultRowHeight()
	{
		$defaultRowHeight = $this->_phpSheet->getDefaultRowDimension()->getRowHeight();

		if ($defaultRowHeight < 0) {
			return;
		}

		// convert to twips
		$defaultRowHeight = (int) 20 * $defaultRowHeight;

		$record   = 0x0225;	  // Record identifier
		$length   = 0x0004;	  // Number of bytes to follow

		$header   = pack("vv", $record, $length);
		$data	 = pack("vv",  1, $defaultRowHeight);
		$this->_prepend($header . $data);
	}

	/**
	* Write BIFF record DEFCOLWIDTH if COLINFO records are in use.
	*
	* @access private
	*/
	function _storeDefcol()
	{
		$defaultColWidth = (int) $this->_phpSheet->getDefaultColumnDimension()->getWidth();
		
		if ($defaultColWidth < 0) {
			return;
		}
		
		$record   = 0x0055;	  // Record identifier
		$length   = 0x0002;	  // Number of bytes to follow

		$header = pack("vv", $record, $length);
		$data = pack("v", $defaultColWidth);
		$this->_prepend($header . $data);
	}

	/**
	* Write BIFF record COLINFO to define column widths
	*
	* Note: The SDK says the record length is 0x0B but Excel writes a 0x0C
	* length record.
	*
	* @access private
	* @param array $col_array This is the only parameter received and is composed of the following:
	*				0 => First formatted column,
	*				1 => Last formatted column,
	*				2 => Col width (8.43 is Excel default),
	*				3 => The optional XF format of the column,
	*				4 => Option flags.
	*				5 => Optional outline level
	*/
	function _storeColinfo($col_array)
	{
		if (isset($col_array[0])) {
			$colFirst = $col_array[0];
		}
		if (isset($col_array[1])) {
			$colLast = $col_array[1];
		}
		if (isset($col_array[2])) {
			$coldx = $col_array[2];
		} else {
			$coldx = 8.43;
		}
		if (isset($col_array[3])) {
			$format = $col_array[3];
		} else {
			$format = 0;
		}
		if (isset($col_array[4])) {
			$grbit = $col_array[4];
		} else {
			$grbit = 0;
		}
		if (isset($col_array[5])) {
			$level = $col_array[5];
		} else {
			$level = 0;
		}
		$record   = 0x007D;		  // Record identifier
		$length   = 0x000B;		  // Number of bytes to follow

		$coldx   *= 256;			 // Convert to units of 1/256 of a char

		$ixfe	 = $this->_XF($format);
		$reserved = 0x00;			// Reserved

		$level = max(0, min($level, 7));
		$grbit |= $level << 8;

		$header   = pack("vv",	 $record, $length);
		$data	 = pack("vvvvvC", $colFirst, $colLast, $coldx,
								   $ixfe, $grbit, $reserved);
		$this->_prepend($header.$data);
	}

	/**
	* Write BIFF record SELECTION.
	*
	* @access private
	* @param array $array array containing ($rwFirst,$colFirst,$rwLast,$colLast)
	* @see setSelection()
	*/
	function _storeSelection($array)
	{
		list($rwFirst,$colFirst,$rwLast,$colLast) = $array;
		$record   = 0x001D;				  // Record identifier
		$length   = 0x000F;				  // Number of bytes to follow

		$pnn	  = $this->_active_pane;	 // Pane position
		$rwAct	= $rwFirst;				// Active row
		$colAct   = $colFirst;			   // Active column
		$irefAct  = 0;					   // Active cell ref
		$cref	 = 1;					   // Number of refs

		if (!isset($rwLast)) {
			$rwLast   = $rwFirst;	   // Last  row in reference
		}
		if (!isset($colLast)) {
			$colLast  = $colFirst;	  // Last  col in reference
		}

		// Swap last row/col for first row/col as necessary
		if ($rwFirst > $rwLast) {
			list($rwFirst, $rwLast) = array($rwLast, $rwFirst);
		}

		if ($colFirst > $colLast) {
			list($colFirst, $colLast) = array($colLast, $colFirst);
		}

		$header   = pack("vv",		 $record, $length);
		$data	 = pack("CvvvvvvCC",  $pnn, $rwAct, $colAct,
									   $irefAct, $cref,
									   $rwFirst, $rwLast,
									   $colFirst, $colLast);
		$this->_append($header . $data);
	}

	/**
	* Store the MERGEDCELLS records for all ranges of merged cells
	*
	* @access private
	*/
	function _storeMergedCells()
	{
		$mergeCells = $this->_phpSheet->getMergeCells();
		$countMergeCells = count($mergeCells);
		
		if ($countMergeCells == 0) {
			return;
		}
		
		// maximum allowed number of merged cells per record
		// there is room for 1027, but if we set higher than 259, record will be split, fix later
		$maxCountMergeCellsPerRecord = 259;
		
		// record identifier
		$record = 0x00E5;
		
		// counter for total number of merged cells treated so far by the writer
		$i = 0;
		
		// counter for number of merged cells written in record currently being written
		$j = 0;
		
		// initialize record data
		$recordData = '';
		
		// loop through the merged cells
		foreach ($mergeCells as $mergeCell) {
			++$i;
			++$j;

			// extract the row and column indexes
			list($first, $last) = PHPExcel_Cell::splitRange($mergeCell);
			list($firstColumn, $firstRow) = PHPExcel_Cell::coordinateFromString($first);
			list($lastColumn, $lastRow) = PHPExcel_Cell::coordinateFromString($last);

			$recordData .= pack('vvvv', $firstRow - 1, $lastRow - 1, PHPExcel_Cell::columnIndexFromString($firstColumn) - 1, PHPExcel_Cell::columnIndexFromString($lastColumn) - 1);

			// flush record if we have reached limit for number of merged cells, or reached final merged cell
			if ($j == $maxCountMergeCellsPerRecord or $i == $countMergeCells) {
				$recordData = pack('v', $j) . $recordData;
				$length = strlen($recordData);
				$header = pack('vv', $record, $length);
				$this->_append($header . $recordData);
				
				// initialize for next record, if any
				$recordData = '';
				$j = 0;
			}
		}
	}

	/**
	 * Write BIFF record RANGEPROTECTION
	 * 
	 * Openoffice.org's Documentaion of the Microsoft Excel File Format uses term RANGEPROTECTION for these records
	 * Microsoft Office Excel 97-2007 Binary File Format Specification uses term FEAT for these records
	 */
	private function _storeRangeProtection()
	{
		foreach ($this->_phpSheet->getProtectedCells() as $range => $password) {
			// number of ranges, e.g. 'A1:B3 C20:D25'
			$cellRanges = explode(' ', $range);
			$cref = count($cellRanges);

			$recordData = pack(
				'vvVVvCVvVv',
				0x0868,
				0x00,
				0x0000,
				0x0000,
				0x02,
				0x0,
				0x0000,
				$cref,
				0x0000,
				0x00
			);

			foreach ($cellRanges as $cellRange) {
				$recordData .= $this->_writeBIFF8CellRangeAddressFixed($cellRange);
			}

			// the rgbFeat structure
			$recordData .= pack(
				'VV',
				0x0000,
				hexdec($password)
			);

			$recordData .= PHPExcel_Shared_String::UTF8toBIFF8UnicodeLong('p' . md5($recordData));

			$length = strlen($recordData);

			$record = 0x0868;		// Record identifier
			$header = pack("vv", $record, $length);
			$this->_append($header . $recordData);
		}
	}

	/**
	* Write BIFF record EXTERNCOUNT to indicate the number of external sheet
	* references in a worksheet.
	*
	* Excel only stores references to external sheets that are used in formulas.
	* For simplicity we store references to all the sheets in the workbook
	* regardless of whether they are used or not. This reduces the overall
	* complexity and eliminates the need for a two way dialogue between the formula
	* parser the worksheet objects.
	*
	* @access private
	* @param integer $count The number of external sheet references in this worksheet
	*/
	function _storeExterncount($count)
	{
		$record = 0x0016;		  // Record identifier
		$length = 0x0002;		  // Number of bytes to follow

		$header = pack("vv", $record, $length);
		$data   = pack("v",  $count);
		$this->_prepend($header . $data);
	}

	/**
	* Writes the Excel BIFF EXTERNSHEET record. These references are used by
	* formulas. A formula references a sheet name via an index. Since we store a
	* reference to all of the external worksheets the EXTERNSHEET index is the same
	* as the worksheet index.
	*
	* @access private
	* @param string $sheetname The name of a external worksheet
	*/
	function _storeExternsheet($sheetname)
	{
		$record	= 0x0017;		 // Record identifier

		// References to the current sheet are encoded differently to references to
		// external sheets.
		//
		if ($this->name == $sheetname) {
			$sheetname = '';
			$length	= 0x02;  // The following 2 bytes
			$cch	   = 1;	 // The following byte
			$rgch	  = 0x02;  // Self reference
		} else {
			$length	= 0x02 + strlen($sheetname);
			$cch	   = strlen($sheetname);
			$rgch	  = 0x03;  // Reference to a sheet in the current workbook
		}

		$header = pack("vv",  $record, $length);
		$data   = pack("CC", $cch, $rgch);
		$this->_prepend($header . $data . $sheetname);
	}

	/**
	* Writes the Excel BIFF PANE record.
	* The panes can either be frozen or thawed (unfrozen).
	* Frozen panes are specified in terms of an integer number of rows and columns.
	* Thawed panes are specified in terms of Excel's units for rows and columns.
	*
	* @access private
	*/
	function _storePanes()
	{
		$panes = array();
		if ($freezePane = $this->_phpSheet->getFreezePane()) {
			list($column, $row) = PHPExcel_Cell::coordinateFromString($freezePane);
			$panes[0] = $row - 1;
			$panes[1] = PHPExcel_Cell::columnIndexFromString($column) - 1;
		} else {
			// thaw panes
			return;
		}
		
		$y	   = isset($panes[0]) ? $panes[0] : null;
		$x	   = isset($panes[1]) ? $panes[1] : null;
		$rwTop   = isset($panes[2]) ? $panes[2] : null;
		$colLeft = isset($panes[3]) ? $panes[3] : null;
		if (count($panes) > 4) { // if Active pane was received
			$pnnAct = $panes[4];
		} else {
			$pnnAct = null;
		}
		$record  = 0x0041;	   // Record identifier
		$length  = 0x000A;	   // Number of bytes to follow

		// Code specific to frozen or thawed panes.
		if ($this->_phpSheet->getFreezePane()) {
			// Set default values for $rwTop and $colLeft
			if (!isset($rwTop)) {
				$rwTop   = $y;
			}
			if (!isset($colLeft)) {
				$colLeft = $x;
			}
		} else {
			// Set default values for $rwTop and $colLeft
			if (!isset($rwTop)) {
				$rwTop   = 0;
			}
			if (!isset($colLeft)) {
				$colLeft = 0;
			}

			// Convert Excel's row and column units to the internal units.
			// The default row height is 12.75
			// The default column width is 8.43
			// The following slope and intersection values were interpolated.
			//
			$y = 20*$y	  + 255;
			$x = 113.879*$x + 390;
		}


		// Determine which pane should be active. There is also the undocumented
		// option to override this should it be necessary: may be removed later.
		//
		if (!isset($pnnAct)) {
			if ($x != 0 && $y != 0) {
				$pnnAct = 0; // Bottom right
			}
			if ($x != 0 && $y == 0) {
				$pnnAct = 1; // Top right
			}
			if ($x == 0 && $y != 0) {
				$pnnAct = 2; // Bottom left
			}
			if ($x == 0 && $y == 0) {
				$pnnAct = 3; // Top left
			}
		}

		$this->_active_pane = $pnnAct; // Used in _storeSelection

		$header	 = pack("vv",	$record, $length);
		$data	   = pack("vvvvv", $x, $y, $rwTop, $colLeft, $pnnAct);
		$this->_append($header . $data);
	}

	/**
	* Store the page setup SETUP BIFF record.
	*
	* @access private
	*/
	function _storeSetup()
	{
		$record	   = 0x00A1;				  // Record identifier
		$length	   = 0x0022;				  // Number of bytes to follow

		$iPaperSize   = $this->_phpSheet->getPageSetup()->getPaperSize();	// Paper size

		$iScale = $this->_phpSheet->getPageSetup()->getScale() ?
			$this->_phpSheet->getPageSetup()->getScale() : 100;   // Print scaling factor

		$iPageStart   = 0x01;				 // Starting page number
		$iFitWidth	= (int) $this->_phpSheet->getPageSetup()->getFitToWidth();	// Fit to number of pages wide
		$iFitHeight	= (int) $this->_phpSheet->getPageSetup()->getFitToHeight();	// Fit to number of pages high
		$grbit		= 0x00;				 // Option flags
		$iRes		 = 0x0258;			   // Print resolution
		$iVRes		= 0x0258;			   // Vertical print resolution
		
		$numHdr	   = $this->_phpSheet->getPageMargins()->getHeader();  // Header Margin
		
		$numFtr	   = $this->_phpSheet->getPageMargins()->getFooter();   // Footer Margin
		$iCopies	  = 0x01;				 // Number of copies

		$fLeftToRight = 0x0;					 // Print over then down

		// Page orientation
		$fLandscape = ($this->_phpSheet->getPageSetup()->getOrientation() == PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE) ?
			0x0 : 0x1;

		$fNoPls	   = 0x0;					 // Setup not read from printer
		$fNoColor	 = 0x0;					 // Print black and white
		$fDraft	   = 0x0;					 // Print draft quality
		$fNotes	   = 0x0;					 // Print notes
		$fNoOrient	= 0x0;					 // Orientation not set
		$fUsePage	 = 0x0;					 // Use custom starting page

		$grbit		   = $fLeftToRight;
		$grbit		  |= $fLandscape	<< 1;
		$grbit		  |= $fNoPls		<< 2;
		$grbit		  |= $fNoColor	  << 3;
		$grbit		  |= $fDraft		<< 4;
		$grbit		  |= $fNotes		<< 5;
		$grbit		  |= $fNoOrient	 << 6;
		$grbit		  |= $fUsePage	  << 7;

		$numHdr = pack("d", $numHdr);
		$numFtr = pack("d", $numFtr);
		if ($this->_byte_order) { // if it's Big Endian
			$numHdr = strrev($numHdr);
			$numFtr = strrev($numFtr);
		}

		$header = pack("vv", $record, $length);
		$data1  = pack("vvvvvvvv", $iPaperSize,
								   $iScale,
								   $iPageStart,
								   $iFitWidth,
								   $iFitHeight,
								   $grbit,
								   $iRes,
								   $iVRes);
		$data2  = $numHdr.$numFtr;
		$data3  = pack("v", $iCopies);
		$this->_prepend($header . $data1 . $data2 . $data3);
	}

	/**
	* Store the header caption BIFF record.
	*
	* @access private
	*/
	function _storeHeader()
	{
		$record  = 0x0014;			   // Record identifier

		/* removing for now
		// need to fix character count (multibyte!)
		if (strlen($this->_phpSheet->getHeaderFooter()->getOddHeader()) <= 255) {
			$str	  = $this->_phpSheet->getHeaderFooter()->getOddHeader();	   // header string
		} else {
			$str = '';
		}
		*/
		
		if ($this->_BIFF_version == 0x0600) {
			$recordData = PHPExcel_Shared_String::UTF8toBIFF8UnicodeLong($this->_phpSheet->getHeaderFooter()->getOddHeader());
			$length = strlen($recordData);
		} else {
			$cch	  = strlen($str);		 // Length of header string
			$length  = 1 + $cch;			 // Bytes to follow
			$data	  = pack("C",  $cch);
			$recordData = $data . $str;
		}

		$header   = pack("vv", $record, $length);

		$this->_prepend($header . $recordData);
	}

	/**
	* Store the footer caption BIFF record.
	*
	* @access private
	*/
	function _storeFooter()
	{
		$record  = 0x0015;			   // Record identifier

		/* removing for now
		// need to fix character count (multibyte!)
		if (strlen($this->_phpSheet->getHeaderFooter()->getOddFooter()) <= 255) {
			$str = $this->_phpSheet->getHeaderFooter()->getOddFooter();
		} else {
			$str = '';
		}
		*/
		
		if ($this->_BIFF_version == 0x0600) {
			$recordData = PHPExcel_Shared_String::UTF8toBIFF8UnicodeLong($this->_phpSheet->getHeaderFooter()->getOddFooter());
			$length = strlen($recordData);
		} else {
			$cch	  = strlen($str);		 // Length of footer string
			$length  = 1 + $cch;
			$data	  = pack("C",  $cch);
			$recordData = $data . $str;
		}

		$header	= pack("vv", $record, $length);

		$this->_prepend($header . $recordData);
	}

	/**
	* Store the horizontal centering HCENTER BIFF record.
	*
	* @access private
	*/
	function _storeHcenter()
	{
		$record   = 0x0083;			  // Record identifier
		$length   = 0x0002;			  // Bytes to follow

		$fHCenter = $this->_phpSheet->getPageSetup()->getHorizontalCentered() ? 1 : 0;	 // Horizontal centering

		$header	= pack("vv", $record, $length);
		$data	  = pack("v",  $fHCenter);

		$this->_prepend($header.$data);
	}

	/**
	* Store the vertical centering VCENTER BIFF record.
	*
	* @access private
	*/
	function _storeVcenter()
	{
		$record   = 0x0084;			  // Record identifier
		$length   = 0x0002;			  // Bytes to follow

		$fVCenter = $this->_phpSheet->getPageSetup()->getVerticalCentered() ? 1 : 0;	 // Horizontal centering

		$header	= pack("vv", $record, $length);
		$data	  = pack("v",  $fVCenter);
		$this->_prepend($header . $data);
	}

	/**
	* Store the LEFTMARGIN BIFF record.
	*
	* @access private
	*/
	function _storeMarginLeft()
	{
		$record  = 0x0026;				   // Record identifier
		$length  = 0x0008;				   // Bytes to follow

		$margin  = $this->_phpSheet->getPageMargins()->getLeft();	 // Margin in inches

		$header	= pack("vv",  $record, $length);
		$data	  = pack("d",   $margin);
		if ($this->_byte_order) { // if it's Big Endian
			$data = strrev($data);
		}

		$this->_prepend($header . $data);
	}

	/**
	* Store the RIGHTMARGIN BIFF record.
	*
	* @access private
	*/
	function _storeMarginRight()
	{
		$record  = 0x0027;				   // Record identifier
		$length  = 0x0008;				   // Bytes to follow

		$margin  = $this->_phpSheet->getPageMargins()->getRight();	 // Margin in inches

		$header	= pack("vv",  $record, $length);
		$data	  = pack("d",   $margin);
		if ($this->_byte_order) { // if it's Big Endian
			$data = strrev($data);
		}

		$this->_prepend($header . $data);
	}

	/**
	* Store the TOPMARGIN BIFF record.
	*
	* @access private
	*/
	function _storeMarginTop()
	{
		$record  = 0x0028;				   // Record identifier
		$length  = 0x0008;				   // Bytes to follow

		$margin  = $this->_phpSheet->getPageMargins()->getTop();	 // Margin in inches

		$header	= pack("vv",  $record, $length);
		$data	  = pack("d",   $margin);
		if ($this->_byte_order) { // if it's Big Endian
			$data = strrev($data);
		}

		$this->_prepend($header . $data);
	}

	/**
	* Store the BOTTOMMARGIN BIFF record.
	*
	* @access private
	*/
	function _storeMarginBottom()
	{
		$record  = 0x0029;				   // Record identifier
		$length  = 0x0008;				   // Bytes to follow

		$margin  = $this->_phpSheet->getPageMargins()->getBottom();	 // Margin in inches

		$header	= pack("vv",  $record, $length);
		$data	  = pack("d",   $margin);
		if ($this->_byte_order) { // if it's Big Endian
			$data = strrev($data);
		}

		$this->_prepend($header . $data);
	}

	/**
	* Write the PRINTHEADERS BIFF record.
	*
	* @access private
	*/
	function _storePrintHeaders()
	{
		$record	  = 0x002a;				   // Record identifier
		$length	  = 0x0002;				   // Bytes to follow

		$fPrintRwCol = $this->_print_headers;	 // Boolean flag

		$header	  = pack("vv", $record, $length);
		$data		= pack("v", $fPrintRwCol);
		$this->_prepend($header . $data);
	}

	/**
	* Write the PRINTGRIDLINES BIFF record. Must be used in conjunction with the
	* GRIDSET record.
	*
	* @access private
	*/
	function _storePrintGridlines()
	{
		$record	  = 0x002b;					// Record identifier
		$length	  = 0x0002;					// Bytes to follow

		$fPrintGrid  = $this->_phpSheet->getPrintGridlines() ? 1 : 0;	// Boolean flag

		$header	  = pack("vv", $record, $length);
		$data		= pack("v", $fPrintGrid);
		$this->_prepend($header . $data);
	}

	/**
	* Write the GRIDSET BIFF record. Must be used in conjunction with the
	* PRINTGRIDLINES record.
	*
	* @access private
	*/
	function _storeGridset()
	{
		$record	  = 0x0082;						// Record identifier
		$length	  = 0x0002;						// Bytes to follow

		$fGridSet	= !$this->_phpSheet->getPrintGridlines();	 // Boolean flag

		$header	  = pack("vv",  $record, $length);
		$data		= pack("v",   $fGridSet);
		$this->_prepend($header . $data);
	}

	/**
	* Write the GUTS BIFF record. This is used to configure the gutter margins
	* where Excel outline symbols are displayed. The visibility of the gutters is
	* controlled by a flag in WSBOOL.
	*
	* @see _storeWsbool()
	* @access private
	*/
	function _storeGuts()
	{
		$record	  = 0x0080;   // Record identifier
		$length	  = 0x0008;   // Bytes to follow

		$dxRwGut	 = 0x0000;   // Size of row gutter
		$dxColGut	= 0x0000;   // Size of col gutter

		$row_level   = $this->_outline_row_level;
		$col_level   = 0;

		// Calculate the maximum column outline level. The equivalent calculation
		// for the row outline level is carried out in setRow().
		$colcount = count($this->_colinfo);
		for ($i = 0; $i < $colcount; ++$i) {
			$col_level = max($this->_colinfo[$i][5], $col_level);
		}

		// Set the limits for the outline levels (0 <= x <= 7).
		$col_level = max(0, min($col_level, 7));

		// The displayed level is one greater than the max outline levels
		if ($row_level) {
			++$row_level;
		}
		if ($col_level) {
			++$col_level;
		}

		$header	  = pack("vv",   $record, $length);
		$data		= pack("vvvv", $dxRwGut, $dxColGut, $row_level, $col_level);

		$this->_prepend($header.$data);
	}


	/**
	* Write the WSBOOL BIFF record, mainly for fit-to-page. Used in conjunction
	* with the SETUP record.
	*
	* @access private
	*/
	function _storeWsbool()
	{
		$record	  = 0x0081;   // Record identifier
		$length	  = 0x0002;   // Bytes to follow
		$grbit	   = 0x0000;

		// The only option that is of interest is the flag for fit to page. So we
		// set all the options in one go.
		//
		// Set the option flags
		$grbit |= 0x0001;						   // Auto page breaks visible
		if ($this->_outline_style) {
			$grbit |= 0x0020; // Auto outline styles
		}
		if ($this->_phpSheet->getShowSummaryBelow()) {
			$grbit |= 0x0040; // Outline summary below
		}
		if ($this->_phpSheet->getShowSummaryRight()) {
			$grbit |= 0x0080; // Outline summary right
		}
		if ($this->_phpSheet->getPageSetup()->getFitToWidth() || $this->_phpSheet->getPageSetup()->getFitToHeight()) {
			$grbit |= 0x0100; // Page setup fit to page
		}
		if ($this->_outline_on) {
			$grbit |= 0x0400; // Outline symbols displayed
		}

		$header	  = pack("vv", $record, $length);
		$data		= pack("v",  $grbit);
		$this->_prepend($header . $data);
	}

	/**
	 * Write the HORIZONTALPAGEBREAKS and VERTICALPAGEBREAKS BIFF records.
	 */
	private function _storeBreaks()
	{
		// initialize
		$vbreaks = array();
		$hbreaks = array();

		foreach ($this->_phpSheet->getBreaks() as $cell => $breakType) {
			// Fetch coordinates
			$coordinates = PHPExcel_Cell::coordinateFromString($cell);

			// Decide what to do by the type of break
			switch ($breakType) {
				case PHPExcel_Worksheet::BREAK_COLUMN:
					// Add to list of vertical breaks
					$vbreaks[] = PHPExcel_Cell::columnIndexFromString($coordinates[0]) - 1;
					break;

				case PHPExcel_Worksheet::BREAK_ROW:
					// Add to list of horizontal breaks
					$hbreaks[] = $coordinates[1];
					break;

				case PHPExcel_Worksheet::BREAK_NONE:
				default:
					// Nothing to do
					break;
			}
		}
		
		// vertical page breaks
		if (count($vbreaks) > 0) {

			// 1000 vertical pagebreaks appears to be an internal Excel 5 limit.
			// It is slightly higher in Excel 97/200, approx. 1026
			$vbreaks = array_slice($vbreaks, 0, 1000);

			// Sort and filter array of page breaks
			sort($vbreaks, SORT_NUMERIC);
			if ($vbreaks[0] == 0) { // don't use first break if it's 0
				array_shift($vbreaks);
			}

			$record  = 0x001a;			   // Record identifier
			$cbrk	= count($vbreaks);	   // Number of page breaks
			if ($this->_BIFF_version == 0x0600) {
				$length  = 2 + 6 * $cbrk;	  // Bytes to follow
			} else {
				$length  = 2 + 2 * $cbrk;	  // Bytes to follow
			}

			$header  = pack("vv",  $record, $length);
			$data	= pack("v",   $cbrk);

			// Append each page break
			foreach ($vbreaks as $vbreak) {
				if ($this->_BIFF_version == 0x0600) {
					$data .= pack("vvv", $vbreak, 0x0000, 0xffff);
				} else {
					$data .= pack("v", $vbreak);
				}
			}

			$this->_prepend($header . $data);
		}
		
		//horizontal page breaks
		if (count($hbreaks) > 0) {

			// Sort and filter array of page breaks
			sort($hbreaks, SORT_NUMERIC);
			if ($hbreaks[0] == 0) { // don't use first break if it's 0
				array_shift($hbreaks);
			}

			$record  = 0x001b;			   // Record identifier
			$cbrk	= count($hbreaks);	   // Number of page breaks
			if ($this->_BIFF_version == 0x0600) {
				$length  = 2 + 6 * $cbrk;	  // Bytes to follow
			} else {
				$length  = 2 + 2 * $cbrk;	  // Bytes to follow
			}

			$header  = pack("vv", $record, $length);
			$data	= pack("v",  $cbrk);

			// Append each page break
			foreach ($hbreaks as $hbreak) {
				if ($this->_BIFF_version == 0x0600) {
					$data .= pack("vvv", $hbreak, 0x0000, 0x00ff);
				} else {
					$data .= pack("v", $hbreak);
				}
			}

			$this->_prepend($header . $data);
		}
	}

	/**
	* Set the Biff PROTECT record to indicate that the worksheet is protected.
	*
	* @access private
	*/
	function _storeProtect()
	{
		// Exit unless sheet protection has been specified
		if (!$this->_phpSheet->getProtection()->getSheet()) {
			return;
		}

		$record	  = 0x0012;			 // Record identifier
		$length	  = 0x0002;			 // Bytes to follow

		$fLock	   = 1;	// Worksheet is protected

		$header	  = pack("vv", $record, $length);
		$data		= pack("v",  $fLock);

		$this->_prepend($header.$data);
	}

	/**
	* Write the worksheet PASSWORD record.
	*
	* @access private
	*/
	function _storePassword()
	{
		// Exit unless sheet protection and password have been specified
		if (!$this->_phpSheet->getProtection()->getSheet() || !$this->_phpSheet->getProtection()->getPassword()) {
			return;
		}

		$record	  = 0x0013;			   // Record identifier
		$length	  = 0x0002;			   // Bytes to follow

		$wPassword   = hexdec($this->_phpSheet->getProtection()->getPassword());	 // Encoded password

		$header	  = pack("vv", $record, $length);
		$data		= pack("v",  $wPassword);

		$this->_prepend($header . $data);
	}


	/**
	* Insert a 24bit bitmap image in a worksheet.
	*
	* @access public
	* @param integer $row	 The row we are going to insert the bitmap into
	* @param integer $col	 The column we are going to insert the bitmap into
	* @param mixed   $bitmap  The bitmap filename or GD-image resource
	* @param integer $x	   The horizontal position (offset) of the image inside the cell.
	* @param integer $y	   The vertical position (offset) of the image inside the cell.
	* @param float   $scale_x The horizontal scale
	* @param float   $scale_y The vertical scale
	*/
	function insertBitmap($row, $col, $bitmap, $x = 0, $y = 0, $scale_x = 1, $scale_y = 1)
	{
		$bitmap_array = (is_resource($bitmap) ? $this->_processBitmapGd($bitmap) : $this->_processBitmap($bitmap));
		list($width, $height, $size, $data) = $bitmap_array; //$this->_processBitmap($bitmap);

		// Scale the frame of the image.
		$width  *= $scale_x;
		$height *= $scale_y;

		// Calculate the vertices of the image and write the OBJ record
		$this->_positionImage($col, $row, $x, $y, $width, $height);

		// Write the IMDATA record to store the bitmap data
		$record	  = 0x007f;
		$length	  = 8 + $size;
		$cf		  = 0x09;
		$env		 = 0x01;
		$lcb		 = $size;

		$header	  = pack("vvvvV", $record, $length, $cf, $env, $lcb);
		$this->_append($header.$data);
	}

	/**
	* Calculate the vertices that define the position of the image as required by
	* the OBJ record.
	*
	*		 +------------+------------+
	*		 |	 A	  |	  B	 |
	*   +-----+------------+------------+
	*   |	 |(x1,y1)	 |			|
	*   |  1  |(A1)._______|______	  |
	*   |	 |	|			  |	 |
	*   |	 |	|			  |	 |
	*   +-----+----|	BITMAP	|-----+
	*   |	 |	|			  |	 |
	*   |  2  |	|______________.	 |
	*   |	 |			|		(B2)|
	*   |	 |			|	 (x2,y2)|
	*   +---- +------------+------------+
	*
	* Example of a bitmap that covers some of the area from cell A1 to cell B2.
	*
	* Based on the width and height of the bitmap we need to calculate 8 vars:
	*	 $col_start, $row_start, $col_end, $row_end, $x1, $y1, $x2, $y2.
	* The width and height of the cells are also variable and have to be taken into
	* account.
	* The values of $col_start and $row_start are passed in from the calling
	* function. The values of $col_end and $row_end are calculated by subtracting
	* the width and height of the bitmap from the width and height of the
	* underlying cells.
	* The vertices are expressed as a percentage of the underlying cell width as
	* follows (rhs values are in pixels):
	*
	*	   x1 = X / W *1024
	*	   y1 = Y / H *256
	*	   x2 = (X-1) / W *1024
	*	   y2 = (Y-1) / H *256
	*
	*	   Where:  X is distance from the left side of the underlying cell
	*			   Y is distance from the top of the underlying cell
	*			   W is the width of the cell
	*			   H is the height of the cell
	*
	* @access private
	* @note  the SDK incorrectly states that the height should be expressed as a
	*		percentage of 1024.
	* @param integer $col_start Col containing upper left corner of object
	* @param integer $row_start Row containing top left corner of object
	* @param integer $x1		Distance to left side of object
	* @param integer $y1		Distance to top of object
	* @param integer $width	 Width of image frame
	* @param integer $height	Height of image frame
	*/
	function _positionImage($col_start, $row_start, $x1, $y1, $width, $height)
	{
		// Initialise end cell to the same as the start cell
		$col_end	= $col_start;  // Col containing lower right corner of object
		$row_end	= $row_start;  // Row containing bottom right corner of object

		// Zero the specified offset if greater than the cell dimensions
		if ($x1 >= $this->_sizeCol($col_start)) {
			$x1 = 0;
		}
		if ($y1 >= $this->_sizeRow($row_start)) {
			$y1 = 0;
		}

		$width	  = $width  + $x1 -1;
		$height	 = $height + $y1 -1;

		// Subtract the underlying cell widths to find the end cell of the image
		while ($width >= $this->_sizeCol($col_end)) {
			$width -= $this->_sizeCol($col_end);
			++$col_end;
		}

		// Subtract the underlying cell heights to find the end cell of the image
		while ($height >= $this->_sizeRow($row_end)) {
			$height -= $this->_sizeRow($row_end);
			++$row_end;
		}

		// Bitmap isn't allowed to start or finish in a hidden cell, i.e. a cell
		// with zero eight or width.
		//
		if ($this->_sizeCol($col_start) == 0) {
			return;
		}
		if ($this->_sizeCol($col_end)   == 0) {
			return;
		}
		if ($this->_sizeRow($row_start) == 0) {
			return;
		}
		if ($this->_sizeRow($row_end)   == 0) {
			return;
		}

		// Convert the pixel values to the percentage value expected by Excel
		$x1 = $x1	 / $this->_sizeCol($col_start)   * 1024;
		$y1 = $y1	 / $this->_sizeRow($row_start)   *  256;
		$x2 = $width  / $this->_sizeCol($col_end)	 * 1024; // Distance to right side of object
		$y2 = $height / $this->_sizeRow($row_end)	 *  256; // Distance to bottom of object

		$this->_storeObjPicture($col_start, $x1,
								 $row_start, $y1,
								 $col_end, $x2,
								 $row_end, $y2);
	}

	/**
	* Convert the width of a cell from user's units to pixels. By interpolation
	* the relationship is: y = 7x +5. If the width hasn't been set by the user we
	* use the default value. If the col is hidden we use a value of zero.
	*
	* @access private
	* @param integer $col The column
	* @return integer The width in pixels
	*/
	function _sizeCol($col)
	{
		// Look up the cell value to see if it has been changed
		if (isset($this->col_sizes[$col])) {
			if ($this->col_sizes[$col] == 0) {
				return(0);
			} else {
				return(floor(7 * $this->col_sizes[$col] + 5));
			}
		} else {
			return(64);
		}
	}

	/**
	* Convert the height of a cell from user's units to pixels. By interpolation
	* the relationship is: y = 4/3x. If the height hasn't been set by the user we
	* use the default value. If the row is hidden we use a value of zero. (Not
	* possible to hide row yet).
	*
	* @access private
	* @param integer $row The row
	* @return integer The width in pixels
	*/
	function _sizeRow($row)
	{
		// Look up the cell value to see if it has been changed
		if (isset($this->_row_sizes[$row])) {
			if ($this->_row_sizes[$row] == 0) {
				return(0);
			} else {
				return(floor(4/3 * $this->_row_sizes[$row]));
			}
		} else {
			return(17);
		}
	}

	/**
	* Store the OBJ record that precedes an IMDATA record. This could be generalise
	* to support other Excel objects.
	*
	* @access private
	* @param integer $colL Column containing upper left corner of object
	* @param integer $dxL  Distance from left side of cell
	* @param integer $rwT  Row containing top left corner of object
	* @param integer $dyT  Distance from top of cell
	* @param integer $colR Column containing lower right corner of object
	* @param integer $dxR  Distance from right of cell
	* @param integer $rwB  Row containing bottom right corner of object
	* @param integer $dyB  Distance from bottom of cell
	*/
	function _storeObjPicture($colL,$dxL,$rwT,$dyT,$colR,$dxR,$rwB,$dyB)
	{
		$record	  = 0x005d;   // Record identifier
		$length	  = 0x003c;   // Bytes to follow

		$cObj		= 0x0001;   // Count of objects in file (set to 1)
		$OT		  = 0x0008;   // Object type. 8 = Picture
		$id		  = 0x0001;   // Object ID
		$grbit	   = 0x0614;   // Option flags

		$cbMacro	 = 0x0000;   // Length of FMLA structure
		$Reserved1   = 0x0000;   // Reserved
		$Reserved2   = 0x0000;   // Reserved

		$icvBack	 = 0x09;	 // Background colour
		$icvFore	 = 0x09;	 // Foreground colour
		$fls		 = 0x00;	 // Fill pattern
		$fAuto	   = 0x00;	 // Automatic fill
		$icv		 = 0x08;	 // Line colour
		$lns		 = 0xff;	 // Line style
		$lnw		 = 0x01;	 // Line weight
		$fAutoB	  = 0x00;	 // Automatic border
		$frs		 = 0x0000;   // Frame style
		$cf		  = 0x0009;   // Image format, 9 = bitmap
		$Reserved3   = 0x0000;   // Reserved
		$cbPictFmla  = 0x0000;   // Length of FMLA structure
		$Reserved4   = 0x0000;   // Reserved
		$grbit2	  = 0x0001;   // Option flags
		$Reserved5   = 0x0000;   // Reserved


		$header	  = pack("vv", $record, $length);
		$data		= pack("V", $cObj);
		$data	   .= pack("v", $OT);
		$data	   .= pack("v", $id);
		$data	   .= pack("v", $grbit);
		$data	   .= pack("v", $colL);
		$data	   .= pack("v", $dxL);
		$data	   .= pack("v", $rwT);
		$data	   .= pack("v", $dyT);
		$data	   .= pack("v", $colR);
		$data	   .= pack("v", $dxR);
		$data	   .= pack("v", $rwB);
		$data	   .= pack("v", $dyB);
		$data	   .= pack("v", $cbMacro);
		$data	   .= pack("V", $Reserved1);
		$data	   .= pack("v", $Reserved2);
		$data	   .= pack("C", $icvBack);
		$data	   .= pack("C", $icvFore);
		$data	   .= pack("C", $fls);
		$data	   .= pack("C", $fAuto);
		$data	   .= pack("C", $icv);
		$data	   .= pack("C", $lns);
		$data	   .= pack("C", $lnw);
		$data	   .= pack("C", $fAutoB);
		$data	   .= pack("v", $frs);
		$data	   .= pack("V", $cf);
		$data	   .= pack("v", $Reserved3);
		$data	   .= pack("v", $cbPictFmla);
		$data	   .= pack("v", $Reserved4);
		$data	   .= pack("v", $grbit2);
		$data	   .= pack("V", $Reserved5);

		$this->_append($header . $data);
	}

	/**
	* Convert a GD-image into the internal format.
	*
	* @access private
	* @param resource $image The image to process
	* @return array Array with data and properties of the bitmap
	*/
	function _processBitmapGd($image) {
		$width = imagesx($image);
		$height = imagesy($image);

		$data = pack("Vvvvv", 0x000c, $width, $height, 0x01, 0x18);
		for ($j=$height; $j--; ) {
			for ($i=0; $i < $width; ++$i) {
				$color = imagecolorsforindex($image, imagecolorat($image, $i, $j));
				foreach (array("red", "green", "blue") as $key) {
					$color[$key] = $color[$key] + round((255 - $color[$key]) * $color["alpha"] / 127);
				}
				$data .= chr($color["blue"]) . chr($color["green"]) . chr($color["red"]);
			}
			if (3*$width % 4) {
				$data .= str_repeat("\x00", 4 - 3*$width % 4);
			}
		}

		return array($width, $height, strlen($data), $data);
	}

	/**
	* Convert a 24 bit bitmap into the modified internal format used by Windows.
	* This is described in BITMAPCOREHEADER and BITMAPCOREINFO structures in the
	* MSDN library.
	*
	* @access private
	* @param string $bitmap The bitmap to process
	* @return array Array with data and properties of the bitmap
	*/
	function _processBitmap($bitmap)
	{
		// Open file.
		$bmp_fd = @fopen($bitmap,"rb");
		if (!$bmp_fd) {
			throw new Exception("Couldn't import $bitmap");
		}

		// Slurp the file into a string.
		$data = fread($bmp_fd, filesize($bitmap));

		// Check that the file is big enough to be a bitmap.
		if (strlen($data) <= 0x36) {
			throw new Exception("$bitmap doesn't contain enough data.\n");
		}

		// The first 2 bytes are used to identify the bitmap.
		$identity = unpack("A2ident", $data);
		if ($identity['ident'] != "BM") {
			throw new Exception("$bitmap doesn't appear to be a valid bitmap image.\n");
		}

		// Remove bitmap data: ID.
		$data = substr($data, 2);

		// Read and remove the bitmap size. This is more reliable than reading
		// the data size at offset 0x22.
		//
		$size_array   = unpack("Vsa", substr($data, 0, 4));
		$size   = $size_array['sa'];
		$data   = substr($data, 4);
		$size  -= 0x36; // Subtract size of bitmap header.
		$size  += 0x0C; // Add size of BIFF header.

		// Remove bitmap data: reserved, offset, header length.
		$data = substr($data, 12);

		// Read and remove the bitmap width and height. Verify the sizes.
		$width_and_height = unpack("V2", substr($data, 0, 8));
		$width  = $width_and_height[1];
		$height = $width_and_height[2];
		$data   = substr($data, 8);
		if ($width > 0xFFFF) {
			throw new Exception("$bitmap: largest image width supported is 65k.\n");
		}
		if ($height > 0xFFFF) {
			throw new Exception("$bitmap: largest image height supported is 65k.\n");
		}

		// Read and remove the bitmap planes and bpp data. Verify them.
		$planes_and_bitcount = unpack("v2", substr($data, 0, 4));
		$data = substr($data, 4);
		if ($planes_and_bitcount[2] != 24) { // Bitcount
			throw new Exception("$bitmap isn't a 24bit true color bitmap.\n");
		}
		if ($planes_and_bitcount[1] != 1) {
			throw new Exception("$bitmap: only 1 plane supported in bitmap image.\n");
		}

		// Read and remove the bitmap compression. Verify compression.
		$compression = unpack("Vcomp", substr($data, 0, 4));
		$data = substr($data, 4);

		//$compression = 0;
		if ($compression['comp'] != 0) {
			throw new Exception("$bitmap: compression not supported in bitmap image.\n");
		}

		// Remove bitmap data: data size, hres, vres, colours, imp. colours.
		$data = substr($data, 20);

		// Add the BITMAPCOREHEADER data
		$header  = pack("Vvvvv", 0x000c, $width, $height, 0x01, 0x18);
		$data	= $header . $data;

		return (array($width, $height, $size, $data));
	}

	/**
	* Store the window zoom factor. This should be a reduced fraction but for
	* simplicity we will store all fractions with a numerator of 100.
	*
	* @access private
	*/
	function _storeZoom()
	{
		// If scale is 100 we don't need to write a record
		if ($this->_phpSheet->getSheetView()->getZoomScale() == 100) {
			return;
		}

		$record	  = 0x00A0;			   // Record identifier
		$length	  = 0x0004;			   // Bytes to follow

		$header	  = pack("vv", $record, $length);
		$data		= pack("vv", $this->_phpSheet->getSheetView()->getZoomScale(), 100);
		$this->_append($header . $data);
	}

	/**
	* Store the DVAL and DV records.
	*
	* @access private
	*/
	function _storeDataValidity()
	{
		$record	  = 0x01b2;	  // Record identifier
		$length	  = 0x0012;	  // Bytes to follow

		$grbit	   = 0x0002;	  // Prompt box at cell, no cached validity data at DV records
		$horPos	  = 0x00000000;  // Horizontal position of prompt box, if fixed position
		$verPos	  = 0x00000000;  // Vertical position of prompt box, if fixed position
		$objId	   = 0xffffffff;  // Object identifier of drop down arrow object, or -1 if not visible

		$header	  = pack('vv', $record, $length);
		$data		= pack('vVVVV', $grbit, $horPos, $verPos, $objId,
									 count($this->_dv));
		$this->_append($header.$data);

		$record = 0x01be;			  // Record identifier
		foreach ($this->_dv as $dv) {
			$length = strlen($dv);	  // Bytes to follow
			$header = pack("vv", $record, $length);
			$this->_append($header . $dv);
		}
	}

	/**
	 * Set sheet dimensions
	 *
	 * @param int $firstRowIndex
	 * @param int $lastRowIndex
	 * @param int $firstColumnIndex
	 * @param int $lastColumnIndex
	 */
	public function setDimensions($firstRowIndex = 0, $lastRowIndex = -1, $firstColumnIndex = 0, $lastColumnIndex = -1)
	{
		$this->_firstRowIndex = $firstRowIndex;
		$this->_lastRowIndex = $lastRowIndex;
		$this->_firstColumnIndex = $firstColumnIndex;
		$this->_lastColumnIndex = $lastColumnIndex;
	}

}
