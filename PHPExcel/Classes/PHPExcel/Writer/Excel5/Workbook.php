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
*    PHPExcel_Writer_Excel5_Writer:  A library for generating Excel Spreadsheets
*    Copyright (c) 2002-2003 Xavier Noguer xnoguer@rezebra.com
*
*    This library is free software; you can redistribute it and/or
*    modify it under the terms of the GNU Lesser General Public
*    License as published by the Free Software Foundation; either
*    version 2.1 of the License, or (at your option) any later version.
*
*    This library is distributed in the hope that it will be useful,
*    but WITHOUT ANY WARRANTY; without even the implied warranty of
*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
*    Lesser General Public License for more details.
*
*    You should have received a copy of the GNU Lesser General Public
*    License along with this library; if not, write to the Free Software
*    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

require_once 'PHPExcel/Writer/Excel5/Format.php';
require_once 'PHPExcel/Writer/Excel5/BIFFwriter.php';
require_once 'PHPExcel/Writer/Excel5/Worksheet.php';
require_once 'PHPExcel/Writer/Excel5/Parser.php';
require_once 'PHPExcel/Shared/Date.php';
require_once 'PHPExcel/Shared/OLE/OLE_Root.php';
require_once 'PHPExcel/Shared/OLE/OLE_File.php';
require_once 'PHPExcel/Shared/String.php';

/**
* Class for generating Excel Spreadsheets
*
* @author   Xavier Noguer <xnoguer@rezebra.com>
* @category PHPExcel
* @package  PHPExcel_Writer_Excel5
*/

class PHPExcel_Writer_Excel5_Workbook extends PHPExcel_Writer_Excel5_BIFFwriter
{
    /**
    * Filename for the Workbook
    * @var string
    */
    var $_filename;

    /**
    * Formula parser
    * @var object Parser
    */
    var $_parser;

    /**
    * The active worksheet of the workbook (0 indexed)
    * @var integer
    */
    var $_activesheet;

    /**
    * 1st displayed worksheet in the workbook (0 indexed)
    * @var integer
    */
    var $_firstsheet;

    /**
    * Number of workbook tabs selected
    * @var integer
    */
    var $_selected;

    /**
    * Index for creating adding new formats to the workbook
    * @var integer
    */
    var $_xf_index;

    /**
    * Flag for preventing close from being called twice.
    * @var integer
    * @see close()
    */
    var $_fileclosed;

    /**
    * The BIFF file size for the workbook.
    * @var integer
    * @see _calcSheetOffsets()
    */
    var $_biffsize;

    /**
    * The default sheetname for all sheets created.
    * @var string
    */
    var $_sheetname;

    /**
    * The default XF format.
    * @var object Format
    */
    var $_tmp_format;

    /**
    * Array containing references to all of this workbook's worksheets
    * @var array
    */
    var $_worksheets;

    /**
    * Array of sheetnames for creating the EXTERNSHEET records
    * @var array
    */
    var $_sheetnames;

    /**
    * Array containing references to all of this workbook's formats
    * @var array
    */
    var $_formats;

    /**
    * Array containing the colour palette
    * @var array
    */
    var $_palette;

    /**
    * The default format for URLs.
    * @var object Format
    */
    var $_url_format;

    /**
    * The codepage indicates the text encoding used for strings
    * @var integer
    */
    var $_codepage;

    /**
    * The country code used for localization
    * @var integer
    */
    var $_country_code;

    /**
    * The temporary dir for storing the OLE file
    * @var string
    */
    var $_tmp_dir;

    /**
    * number of bytes for sizeinfo of strings
    * @var integer
    */
    var $_string_sizeinfo_size;

    /**
    * Workbook
    * @var PHPExcel
    */
    private $_phpExcel;

    
	
	/**
    * Class constructor
    *
    * @param string filename for storing the workbook. "-" for writing to stdout.
	* @param PHPExcel $phpExcel The Workbook
    * @access public
    */
    function PHPExcel_Writer_Excel5_Workbook($filename, $phpExcel)
    {
        // It needs to call its parent's constructor explicitly
        $this->PHPExcel_Writer_Excel5_BIFFwriter();

        $this->_filename         = $filename;
        $this->_parser           = new PHPExcel_Writer_Excel5_Parser($this->_byte_order, $this->_BIFF_version);
        $this->_activesheet      = 0;
        $this->_firstsheet       = 0;
        $this->_selected         = 0;
        $this->_xf_index         = 16; // 15 style XF's and 1 cell XF.
        $this->_fileclosed       = 0;
        $this->_biffsize         = 0;
        $this->_sheetname        = 'Sheet';
		
        $this->_tmp_format       = new PHPExcel_Writer_Excel5_Format($this->_BIFF_version);
		$this->_tmp_format->setFontFamily($phpExcel->getSheet(0)->getDefaultStyle()->getFont()->getName());
		$this->_tmp_format->setSize($phpExcel->getSheet(0)->getDefaultStyle()->getFont()->getSize());
		
        $this->_worksheets       = array();
        $this->_sheetnames       = array();
        $this->_formats          = array();
        $this->_palette          = array();
        $this->_codepage         = 0x04E4; // FIXME: should change for BIFF8
        $this->_country_code     = -1;
        $this->_string_sizeinfo  = 3;

        // Add the default format for hyperlinks
        $this->_url_format =& $this->addFormat(array('color' => 'blue', 'underline' => 1));
        $this->_str_total       = 0;
        $this->_str_unique      = 0;
        $this->_str_table       = array();
        $this->_setPaletteXl97();
        $this->_tmp_dir         = '';
		
		$this->_phpExcel = $phpExcel;
    }

    /**
    * Calls finalization methods.
    * This method should always be the last one to be called on every workbook
    *
    * @access public
    * @return mixed true on success
    */
    function close()
    {
        if ($this->_fileclosed) { // Prevent close() from being called twice.
            return true;
        }
        $res = $this->_storeWorkbook();
		foreach ($this->_worksheets as $sheet) {
			$sheet->cleanup();
		}
        $this->_fileclosed = 1;
        return true;
    }

    /**
    * An accessor for the _worksheets[] array
    * Returns an array of the worksheet objects in a workbook
    * It actually calls to worksheets()
    *
    * @access public
    * @see worksheets()
    * @return array
    */
    function sheets()
    {
        return $this->worksheets();
    }

    /**
    * An accessor for the _worksheets[] array.
    * Returns an array of the worksheet objects in a workbook
    *
    * @access public
    * @return array
    */
    function worksheets()
    {
        return $this->_worksheets;
    }

    /**
    * Sets the BIFF version.
    * This method exists just to access experimental functionality
    * from BIFF8. It will be deprecated !
    * Only possible value is 8 (Excel 97/2000).
    * For any other value it fails silently.
    *
    * @access public
    * @param integer $version The BIFF version
    */
    function setVersion($version)
    {
        if ($version == 8) { // only accept version 8
            $version = 0x0600;
            $this->_BIFF_version = $version;
            // change BIFFwriter limit for CONTINUE records
            $this->_limit = 8228;
            $this->_tmp_format->_BIFF_version = $version;
            $this->_url_format->_BIFF_version = $version;
            $this->_parser->_BIFF_version = $version;
            $this->_codepage = 0x04B0;

            $total_worksheets = count($this->_worksheets);
            // change version for all worksheets too
            for ($i = 0; $i < $total_worksheets; ++$i) {
                $this->_worksheets[$i]->_BIFF_version = $version;
            }

            $total_formats = count($this->_formats);
            // change version for all formats too
            for ($i = 0; $i < $total_formats; ++$i) {
                $this->_formats[$i]->_BIFF_version = $version;
            }
        }
    }

    /**
    * Set the country identifier for the workbook
    *
    * @access public
    * @param integer $code Is the international calling country code for the
    *                      chosen country.
    */
    function setCountry($code)
    {
        $this->_country_code = $code;
    }

    /**
    * Add a new worksheet to the Excel workbook.
    * If no name is given the name of the worksheet will be Sheeti$i, with
    * $i in [1..].
    *
    * @access public
    * @param string $name the optional name of the worksheet
	* @param PHPExcel_Worksheet $phpSheet
    * @return mixed reference to a worksheet object on success
    */
    function &addWorksheet($name = '', $phpSheet = null)
    {
        $index     = count($this->_worksheets);
        $sheetname = $this->_sheetname;

        if ($name == '') {
            $name = $sheetname.($index+1);
        }

        // Check that sheetname is <= 31 chars (Excel limit before BIFF8).
        if ($this->_BIFF_version != 0x0600)
        {
            if (strlen($name) > 31) {
                throw new Exception("Sheetname $name must be <= 31 chars");
            }
        }

        // Check that the worksheet name doesn't already exist: a fatal Excel error.
        $total_worksheets = count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
            if ($this->_worksheets[$i]->getName() == $name) {
                throw new Exception("Worksheet '$name' already exists");
            }
        }

        $worksheet = new PHPExcel_Writer_Excel5_Worksheet($this->_BIFF_version,
                                   $name, $index,
                                   $this->_activesheet, $this->_firstsheet,
                                   $this->_str_total, $this->_str_unique,
                                   $this->_str_table, $this->_url_format,
                                   $this->_parser, $this->_tmp_dir,
								   $phpSheet);

        $this->_worksheets[$index] = &$worksheet;    // Store ref for iterator
        $this->_sheetnames[$index] = $name;          // Store EXTERNSHEET names
        $this->_parser->setExtSheet($name, $index);  // Register worksheet name with parser

		// for BIFF8
		if ($this->_BIFF_version == 0x0600) {
			$supbook_index = 0x00;
			$ref = pack('vvv', $supbook_index, $total_worksheets, $total_worksheets);
			$this->_parser->_references[] = $ref;  // Register reference with parser
		}


        return $worksheet;
    }

    /**
    * Add a new format to the Excel workbook.
    * Also, pass any properties to the Format constructor.
    *
    * @access public
    * @param array $properties array with properties for initializing the format.
    * @return &PHPExcel_Writer_Excel5_Format reference to an Excel Format
    */
    function &addFormat($properties = array())
    {
        $format = new PHPExcel_Writer_Excel5_Format($this->_BIFF_version, $this->_xf_index, $properties);
        $this->_xf_index += 1;
        $this->_formats[] = &$format;
        return $format;
    }

    /**
    * Change the RGB components of the elements in the colour palette.
    *
    * @access public
    * @param integer $index colour index
    * @param integer $red   red RGB value [0-255]
    * @param integer $green green RGB value [0-255]
    * @param integer $blue  blue RGB value [0-255]
    * @return integer The palette index for the custom color
    */
    function setCustomColor($index, $red, $green, $blue)
    {
        // Match a HTML #xxyyzz style parameter
        /*if (defined $_[1] and $_[1] =~ /^#(\w\w)(\w\w)(\w\w)/ ) {
            @_ = ($_[0], hex $1, hex $2, hex $3);
        }*/

        // Check that the colour index is the right range
        if ($index < 8 or $index > 64) {
            // TODO: assign real error codes
            throw new Exception("Color index $index outside range: 8 <= index <= 64");
        }

        // Check that the colour components are in the right range
        if (($red   < 0 or $red   > 255) ||
            ($green < 0 or $green > 255) ||
            ($blue  < 0 or $blue  > 255))
        {
            throw new Exception("Color component outside range: 0 <= color <= 255");
        }

        $index -= 8; // Adjust colour index (wingless dragonfly)

        // Set the RGB value
        $this->_palette[$index] = array($red, $green, $blue, 0);
        return($index + 8);
    }

    /**
    * Sets the colour palette to the Excel 97+ default.
    *
    * @access private
    */
    function _setPaletteXl97()
    {
        $this->_palette = array(
                           array(0x00, 0x00, 0x00, 0x00),   // 8
                           array(0xff, 0xff, 0xff, 0x00),   // 9
                           array(0xff, 0x00, 0x00, 0x00),   // 10
                           array(0x00, 0xff, 0x00, 0x00),   // 11
                           array(0x00, 0x00, 0xff, 0x00),   // 12
                           array(0xff, 0xff, 0x00, 0x00),   // 13
                           array(0xff, 0x00, 0xff, 0x00),   // 14
                           array(0x00, 0xff, 0xff, 0x00),   // 15
                           array(0x80, 0x00, 0x00, 0x00),   // 16
                           array(0x00, 0x80, 0x00, 0x00),   // 17
                           array(0x00, 0x00, 0x80, 0x00),   // 18
                           array(0x80, 0x80, 0x00, 0x00),   // 19
                           array(0x80, 0x00, 0x80, 0x00),   // 20
                           array(0x00, 0x80, 0x80, 0x00),   // 21
                           array(0xc0, 0xc0, 0xc0, 0x00),   // 22
                           array(0x80, 0x80, 0x80, 0x00),   // 23
                           array(0x99, 0x99, 0xff, 0x00),   // 24
                           array(0x99, 0x33, 0x66, 0x00),   // 25
                           array(0xff, 0xff, 0xcc, 0x00),   // 26
                           array(0xcc, 0xff, 0xff, 0x00),   // 27
                           array(0x66, 0x00, 0x66, 0x00),   // 28
                           array(0xff, 0x80, 0x80, 0x00),   // 29
                           array(0x00, 0x66, 0xcc, 0x00),   // 30
                           array(0xcc, 0xcc, 0xff, 0x00),   // 31
                           array(0x00, 0x00, 0x80, 0x00),   // 32
                           array(0xff, 0x00, 0xff, 0x00),   // 33
                           array(0xff, 0xff, 0x00, 0x00),   // 34
                           array(0x00, 0xff, 0xff, 0x00),   // 35
                           array(0x80, 0x00, 0x80, 0x00),   // 36
                           array(0x80, 0x00, 0x00, 0x00),   // 37
                           array(0x00, 0x80, 0x80, 0x00),   // 38
                           array(0x00, 0x00, 0xff, 0x00),   // 39
                           array(0x00, 0xcc, 0xff, 0x00),   // 40
                           array(0xcc, 0xff, 0xff, 0x00),   // 41
                           array(0xcc, 0xff, 0xcc, 0x00),   // 42
                           array(0xff, 0xff, 0x99, 0x00),   // 43
                           array(0x99, 0xcc, 0xff, 0x00),   // 44
                           array(0xff, 0x99, 0xcc, 0x00),   // 45
                           array(0xcc, 0x99, 0xff, 0x00),   // 46
                           array(0xff, 0xcc, 0x99, 0x00),   // 47
                           array(0x33, 0x66, 0xff, 0x00),   // 48
                           array(0x33, 0xcc, 0xcc, 0x00),   // 49
                           array(0x99, 0xcc, 0x00, 0x00),   // 50
                           array(0xff, 0xcc, 0x00, 0x00),   // 51
                           array(0xff, 0x99, 0x00, 0x00),   // 52
                           array(0xff, 0x66, 0x00, 0x00),   // 53
                           array(0x66, 0x66, 0x99, 0x00),   // 54
                           array(0x96, 0x96, 0x96, 0x00),   // 55
                           array(0x00, 0x33, 0x66, 0x00),   // 56
                           array(0x33, 0x99, 0x66, 0x00),   // 57
                           array(0x00, 0x33, 0x00, 0x00),   // 58
                           array(0x33, 0x33, 0x00, 0x00),   // 59
                           array(0x99, 0x33, 0x00, 0x00),   // 60
                           array(0x99, 0x33, 0x66, 0x00),   // 61
                           array(0x33, 0x33, 0x99, 0x00),   // 62
                           array(0x33, 0x33, 0x33, 0x00),   // 63
                         );
    }

    /**
    * Assemble worksheets into a workbook and send the BIFF data to an OLE
    * storage.
    *
    * @access private
    * @return mixed true on success
    */
    function _storeWorkbook()
    {
        if (count($this->_worksheets) == 0) {
            return true;
        }

        // Ensure that at least one worksheet has been selected.
        if ($this->_activesheet == 0) {
            $this->_worksheets[0]->selected = 1;
        }

        // Calculate the number of selected worksheet tabs and call the finalization
        // methods for each worksheet
        $total_worksheets = count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
            if ($this->_worksheets[$i]->selected) {
                $this->_selected++;
            }
            $this->_worksheets[$i]->close($this->_sheetnames);
        }

        // Add part 1 of the Workbook globals, what goes before the SHEET records
        $this->_storeBof(0x0005);
        $this->_storeCodepage();
        if ($this->_BIFF_version == 0x0600) {
            $this->_storeWindow1();
        }
        if ($this->_BIFF_version == 0x0500) {
            $this->_storeExterns();    // For print area and repeat rows
            $this->_storeNames();      // For print area and repeat rows
        }
        if ($this->_BIFF_version == 0x0500) {
            $this->_storeWindow1();
        }
        $this->_storeDatemode();
        $this->_storeAllFonts();
        $this->_storeAllNumFormats();
        $this->_storeAllXfs();
        $this->_storeAllStyles();
        $this->_storePalette();
		$this->_calculateSharedStringsSizes();

        // Prepare part 3 of the workbook global stream, what goes after the SHEET records
		$part3 = '';
		if ($this->_country_code != -1) {
            $part3 .= $this->writeCountry();
        }

        if ($this->_BIFF_version == 0x0600) {
            $part3 .= $this->writeSupbookInternal();
            /* TODO: store external SUPBOOK records and XCT and CRN records
            in case of external references for BIFF8 */
            $part3 .= $this->writeExternsheetBiff8();
			$part3 .= $this->writeAllDefinedNamesBiff8();
            $part3 .= $this->writeSharedStringsTable();
        }

        $part3 .= $this->writeEof();

        // Add part 2 of the Workbook globals, the SHEET records
        $this->_calcSheetOffsets();
        for ($i = 0; $i < $total_worksheets; ++$i) {
            $this->_storeBoundsheet($this->_worksheets[$i]->name,$this->_worksheets[$i]->offset);
        }

		// Add part 3 of the Workbook globals
		$this->_data .= $part3;

        // Store the workbook in an OLE container
        $res = $this->_storeOLEFile();
        return true;
    }

    /**
    * Sets the temp dir used for storing the OLE file
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
    * Store the workbook in an OLE container
    *
    * @access private
    * @return mixed true on success
    */
    function _storeOLEFile()
    {
        $OLE = new PHPExcel_Shared_OLE_PPS_File(PHPExcel_Shared_OLE::Asc2Ucs('Book'));
        if ($this->_tmp_dir != '') {
            $OLE->setTempDir($this->_tmp_dir);
        }
        $res = $OLE->init();
        $OLE->append($this->_data);

        $total_worksheets = count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
            while ($tmp = $this->_worksheets[$i]->getData()) {
                $OLE->append($tmp);
            }
        }

        $root = new PHPExcel_Shared_OLE_PPS_Root(time(), time(), array($OLE));
        if ($this->_tmp_dir != '') {
            $root->setTempDir($this->_tmp_dir);
        }

        $res = $root->save($this->_filename);
        return true;
    }

    /**
    * Calculate offsets for Worksheet BOF records.
    *
    * @access private
    */
    function _calcSheetOffsets()
    {
        if ($this->_BIFF_version == 0x0600) {
            $boundsheet_length = 12;  // fixed length for a BOUNDSHEET record
        } else {
            $boundsheet_length = 11;
        }

		// size of Workbook globals part 1 + 3
        $offset            = $this->_datasize;

        // add size of Workbook globals part 2, the length of the SHEET records
        $total_worksheets = count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
			if ($this->_BIFF_version == 0x0600) {
				if (function_exists('mb_strlen') and function_exists('mb_convert_encoding')) {
					// sheet name is stored in uncompressed notation
					$offset += $boundsheet_length + 2 * mb_strlen($this->_worksheets[$i]->name, 'UTF-8');
				} else {
					// sheet name is stored in compressed notation, and ASCII is assumed
					$offset += $boundsheet_length + strlen($this->_worksheets[$i]->name);
				}
			} else {
				$offset += $boundsheet_length + strlen($this->_worksheets[$i]->name);
			}
        }

        // add the sizes of each of the Sheet substreams, respectively
		for ($i = 0; $i < $total_worksheets; ++$i) {
            $this->_worksheets[$i]->offset = $offset;
            $offset += $this->_worksheets[$i]->_datasize;
        }
        $this->_biffsize = $offset;
    }

    /**
    * Store the Excel FONT records.
    *
    * @access private
    */
    function _storeAllFonts()
    {
        // tmp_format is added by the constructor. We use this to write the default XF's
        $format = $this->_tmp_format;
        $font   = $format->getFont();

        // Note: Fonts are 0-indexed. According to the SDK there is no index 4,
        // so the following fonts are 0, 1, 2, 3, 5
        //
        for ($i = 1; $i <= 5; ++$i){
            $this->_append($font);
        }

        // Iterate through the XF objects and write a FONT record if it isn't the
        // same as the default FONT and if it hasn't already been used.
        //
        $fonts = array();
        $index = 6;                  // The first user defined FONT

        $key = $format->getFontKey(); // The default font from _tmp_format
        $fonts[$key] = 0;             // Index of the default font

        $total_formats = count($this->_formats);
        for ($i = 0; $i < $total_formats; ++$i) {
            $key = $this->_formats[$i]->getFontKey();
            if (isset($fonts[$key])) {
                // FONT has already been used
                $this->_formats[$i]->font_index = $fonts[$key];
            } else {
                // Add a new FONT record
                $fonts[$key]        = $index;
                $this->_formats[$i]->font_index = $index;
                ++$index;
                $font = $this->_formats[$i]->getFont();
                $this->_append($font);
            }
        }
    }

    /**
    * Store user defined numerical formats i.e. FORMAT records
    *
    * @access private
    */
    function _storeAllNumFormats()
    {
        // Leaning num_format syndrome
        $hash_num_formats = array();
        $num_formats      = array();
        $index = 164;

        // Iterate through the XF objects and write a FORMAT record if it isn't a
        // built-in format type and if the FORMAT string hasn't already been used.
        $total_formats = count($this->_formats);
        for ($i = 0; $i < $total_formats; ++$i) {
            $num_format = $this->_formats[$i]->_num_format;

			//////////////////////////////////////////////////////////////////////////////////////////
			// Removing this block for now. No true support for built-in number formats in PHPExcel //
			//////////////////////////////////////////////////////////////////////////////////////////
            /**
            // Check if $num_format is an index to a built-in format.
            // Also check for a string of zeros, which is a valid format string
            // but would evaluate to zero.
            //
            if (!preg_match("/^0+\d/", $num_format)) {
                if (preg_match("/^\d+$/", $num_format)) { // built-in format
                    continue;
                }
            }
            **/

            if (isset($hash_num_formats[$num_format])) {
                // FORMAT has already been used
                $this->_formats[$i]->_num_format = $hash_num_formats[$num_format];
            } else{
                // Add a new FORMAT
                $hash_num_formats[$num_format]  = $index;
                $this->_formats[$i]->_num_format = $index;
                $num_formats[] = $num_format;
                ++$index;
            }
        }

        // Write the new FORMAT records starting from 0xA4
        $index = 164;
        foreach ($num_formats as $num_format) {
            $this->_storeNumFormat($num_format,$index);
            ++$index;
        }
    }

    /**
    * Write all XF records.
    *
    * @access private
    */
    function _storeAllXfs()
    {
        // _tmp_format is added by the constructor. We use this to write the default XF's
        // The default font index is 0
        //
        $format = $this->_tmp_format;
        for ($i = 0; $i <= 14; ++$i) {
            $xf = $format->getXf('style'); // Style XF
            $this->_append($xf);
        }

        $xf = $format->getXf('cell');      // Cell XF
        $this->_append($xf);

        // User defined XFs
        $total_formats = count($this->_formats);
        for ($i = 0; $i < $total_formats; ++$i) {
            $xf = $this->_formats[$i]->getXf('cell');
            $this->_append($xf);
        }
    }

    /**
    * Write all STYLE records.
    *
    * @access private
    */
    function _storeAllStyles()
    {
        $this->_storeStyle();
    }

    /**
    * Write the EXTERNCOUNT and EXTERNSHEET records. These are used as indexes for
    * the NAME records.
    *
    * @access private
    */
    function _storeExterns()
    {
        // Create EXTERNCOUNT with number of worksheets
        $this->_storeExterncount(count($this->_worksheets));

        // Create EXTERNSHEET for each worksheet
        foreach ($this->_sheetnames as $sheetname) {
            $this->_storeExternsheet($sheetname);
        }
    }

    /**
    * Write the NAME record to define the print area and the repeat rows and cols.
    *
    * @access private
    */
    function _storeNames()
    {
        // Create the print area NAME records
        $total_worksheets = count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
            // Write a Name record if the print area has been defined
            if (isset($this->_worksheets[$i]->print_rowmin)) {
                $this->_storeNameShort(
                    $this->_worksheets[$i]->index,
                    0x06, // NAME type
                    $this->_worksheets[$i]->print_rowmin,
                    $this->_worksheets[$i]->print_rowmax,
                    $this->_worksheets[$i]->print_colmin,
                    $this->_worksheets[$i]->print_colmax
                    );
            }
        }

        // Create the print title NAME records
        $total_worksheets = count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
            $rowmin = $this->_worksheets[$i]->title_rowmin;
            $rowmax = $this->_worksheets[$i]->title_rowmax;
            $colmin = $this->_worksheets[$i]->title_colmin;
            $colmax = $this->_worksheets[$i]->title_colmax;

            // Determine if row + col, row, col or nothing has been defined
            // and write the appropriate record
            //
            if (isset($rowmin) && isset($colmin)) {
                // Row and column titles have been defined.
                // Row title has been defined.
                $this->_storeNameLong(
                    $this->_worksheets[$i]->index,
                    0x07, // NAME type
                    $rowmin,
                    $rowmax,
                    $colmin,
                    $colmax
                    );
            } elseif (isset($rowmin)) {
                // Row title has been defined.
                $this->_storeNameShort(
                    $this->_worksheets[$i]->index,
                    0x07, // NAME type
                    $rowmin,
                    $rowmax,
                    0x00,
                    0xff
                    );
            } elseif (isset($colmin)) {
                // Column title has been defined.
                $this->_storeNameShort(
                    $this->_worksheets[$i]->index,
                    0x07, // NAME type
                    0x0000,
                    0x3fff,
                    $colmin,
                    $colmax
                    );
            } else {
                // Print title hasn't been defined.
            }
        }
    }


/**
 * Writes all the DEFINEDNAME records (BIFF8).
 * So far this is only used for repeating rows/columns (print titles) and print areas
 */
public function writeAllDefinedNamesBiff8()
{
	$chunk = '';

	// write the print titles (repeating rows, columns), if any
	$total_worksheets = count($this->_worksheets);
	for ($i = 0; $i < $total_worksheets; ++$i) {
		// repeatColumns / repeatRows
		if ($this->_phpExcel->getSheet($i)->getPageSetup()->isColumnsToRepeatAtLeftSet() || $this->_phpExcel->getSheet($i)->getPageSetup()->isRowsToRepeatAtTopSet()) {
			// Row and column titles have been defined
			
			// Columns to repeat
			if ($this->_phpExcel->getSheet($i)->getPageSetup()->isColumnsToRepeatAtLeftSet()) {
				$repeat = $this->_phpExcel->getSheet($i)->getPageSetup()->getColumnsToRepeatAtLeft();
				$colmin = PHPExcel_Cell::columnIndexFromString($repeat[0]) - 1;
				$colmax = PHPExcel_Cell::columnIndexFromString($repeat[1]) - 1;
			} else {
				$colmin = 0;
				$colmax = 255;
			}
			// Rows to repeat
			if ($this->_phpExcel->getSheet($i)->getPageSetup()->isRowsToRepeatAtTopSet()) {
				$repeat = $this->_phpExcel->getSheet($i)->getPageSetup()->getRowsToRepeatAtTop();
				$rowmin = $repeat[0] - 1;
				$rowmax = $repeat[1] - 1;
			} else {
				$rowmin = 0;
				$rowmax = 65535;
			}

			// construct formula data manually because parser does not recognize absolute 3d cell references
			$formulaData = pack('Cvvvvv', 0x3B, $i, $rowmin, $rowmax, $colmin, $colmax);

			// store the DEFINEDNAME record
			$chunk .= $this->writeData($this->writeDefinedNameBiff8(pack('C', 0x07), $formulaData, $i + 1, true));
		}
	}

	// write the print areas, if any
	for ($i = 0; $i < $total_worksheets; ++$i) {
		if ($this->_phpExcel->getSheet($i)->getPageSetup()->isPrintAreaSet()) {
			// Print area
			$printArea = PHPExcel_Cell::splitRange($this->_phpExcel->getSheet($i)->getPageSetup()->getPrintArea());
			$printArea[0] = PHPExcel_Cell::coordinateFromString($printArea[0]);
			$printArea[1] = PHPExcel_Cell::coordinateFromString($printArea[1]);
		
			$print_rowmin = $printArea[0][1] - 1;
			$print_rowmax = $printArea[1][1] - 1;
			$print_colmin = PHPExcel_Cell::columnIndexFromString($printArea[0][0]) - 1;
			$print_colmax = PHPExcel_Cell::columnIndexFromString($printArea[1][0]) - 1;

			// construct formula data manually because parser does not recognize absolute 3d cell references
			$formulaData = pack('Cvvvvv', 0x3B, $i, $print_rowmin, $print_rowmax, $print_colmin, $print_colmax);

			// store the DEFINEDNAME record
			$chunk .= $this->writeData($this->writeDefinedNameBiff8(pack('C', 0x06), $formulaData, $i + 1, true));
		}
	}

	return $chunk;
}

/**
 * Write a DEFINEDNAME record for BIFF8 using explicit binary formula data
 *
 * @param	string		$name			The name in UTF-8
 * @param	string		$formulaData	The binary formula data
 * @param	string		$sheetIndex		1-based sheet index the defined name applies to. 0 = global
 * @param	boolean		$isBuiltIn		Built-in name?
 * @return	string	Complete binary record data
 */
public function writeDefinedNameBiff8($name, $formulaData, $sheetIndex = 0, $isBuiltIn = false)
{
	$record = 0x0018;

	// option flags
	$options = $isBuiltIn ? 0x20 : 0x00;

	// length of the name, character count
	$nlen = function_exists('mb_strlen') ?
		mb_strlen($name, 'UTF8') : strlen($name);

	// name with stripped length field
	$name = substr(PHPExcel_Shared_String::UTF8toBIFF8UnicodeLong($name), 2);

	// size of the formula (in bytes)
	$sz = strlen($formulaData);

	// combine the parts
	$data = pack('vCCvvvCCCC', $options, 0, $nlen, $sz, 0, $sheetIndex, 0, 0, 0, 0)
		. $name . $formulaData;
	$length = strlen($data);

	$header = pack('vv', $record, $length);

	return $header . $data;
}




    /******************************************************************************
    *
    * BIFF RECORDS
    *
    */

    /**
    * Stores the CODEPAGE biff record.
    *
    * @access private
    */
    function _storeCodepage()
    {
        $record          = 0x0042;             // Record identifier
        $length          = 0x0002;             // Number of bytes to follow
        $cv              = $this->_codepage;   // The code page

        $header          = pack('vv', $record, $length);
        $data            = pack('v',  $cv);

        $this->_append($header . $data);
    }

    /**
    * Write Excel BIFF WINDOW1 record.
    *
    * @access private
    */
    function _storeWindow1()
    {
        $record    = 0x003D;                 // Record identifier
        $length    = 0x0012;                 // Number of bytes to follow

        $xWn       = 0x0000;                 // Horizontal position of window
        $yWn       = 0x0000;                 // Vertical position of window
        $dxWn      = 0x25BC;                 // Width of window
        $dyWn      = 0x1572;                 // Height of window

        $grbit     = 0x0038;                 // Option flags
        $ctabsel   = $this->_selected;       // Number of workbook tabs selected
        $wTabRatio = 0x0258;                 // Tab to scrollbar ratio

        $itabFirst = $this->_firstsheet;     // 1st displayed worksheet
        $itabCur   = $this->_activesheet;    // Active worksheet

        $header    = pack("vv",        $record, $length);
        $data      = pack("vvvvvvvvv", $xWn, $yWn, $dxWn, $dyWn,
                                       $grbit,
                                       $itabCur, $itabFirst,
                                       $ctabsel, $wTabRatio);
        $this->_append($header . $data);
    }

    /**
    * Writes Excel BIFF BOUNDSHEET record.
    * FIXME: inconsistent with BIFF documentation
    *
    * @param string  $sheetname Worksheet name
    * @param integer $offset    Location of worksheet BOF
    * @access private
    */
    function _storeBoundsheet($sheetname,$offset)
    {
        $record    = 0x0085;                    // Record identifier
        if ($this->_BIFF_version == 0x0600) {
			//$recordData = $this->_writeUnicodeDataShort($sheetname);
			$recordData = PHPExcel_Shared_String::UTF8toBIFF8UnicodeShort($sheetname);
            $length    = 0x06 + strlen($recordData); // Number of bytes to follow
        } else {
            $length = 0x07 + strlen($sheetname); // Number of bytes to follow
        }

        $grbit     = 0x0000;                    // Visibility and sheet type

        $header    = pack("vv",  $record, $length);
        if ($this->_BIFF_version == 0x0600) {
            $data      = pack("Vv", $offset, $grbit);
			$this->_append($header.$data.$recordData);
        } else {
			$cch       = strlen($sheetname);        // Length of sheet name
            $data      = pack("VvC", $offset, $grbit, $cch);
			$this->_append($header.$data.$sheetname);
        }
    }

    /**
    * Write Internal SUPBOOK record
    *
    * @access private
    */
    public function writeSupbookInternal()
    {
        $record    = 0x01AE;   // Record identifier
        $length    = 0x0004;   // Bytes to follow

        $header    = pack("vv", $record, $length);
        //$data      = pack("vv", count($this->_worksheets), 0x0104);
        $data      = pack("vv", count($this->_worksheets), 0x0401);
        //$this->_append($header . $data);
        return $this->writeData($header . $data);
    }

    /**
    * Writes the Excel BIFF EXTERNSHEET record. These references are used by
    * formulas.
    *
    * @param string $sheetname Worksheet name
    * @access private
    */
    public function writeExternsheetBiff8()
    {
        $total_references = count($this->_parser->_references);
        $record   = 0x0017;                     // Record identifier
        $length   = 2 + 6 * $total_references;  // Number of bytes to follow

        $supbook_index = 0;           // FIXME: only using internal SUPBOOK record
        $header           = pack("vv",  $record, $length);
        $data             = pack('v', $total_references);
        for ($i = 0; $i < $total_references; ++$i) {
            $data .= $this->_parser->_references[$i];
        }
        //$this->_append($header . $data);
        return $this->writeData($header . $data);
    }

    /**
    * Write Excel BIFF STYLE records.
    *
    * @access private
    */
    function _storeStyle()
    {
        $record    = 0x0293;   // Record identifier
        $length    = 0x0004;   // Bytes to follow

        $ixfe      = 0x8000;   // Index to style XF
        $BuiltIn   = 0x00;     // Built-in style
        $iLevel    = 0xff;     // Outline style level

        $header    = pack("vv",  $record, $length);
        $data      = pack("vCC", $ixfe, $BuiltIn, $iLevel);
        $this->_append($header . $data);
    }


    /**
    * Writes Excel FORMAT record for non "built-in" numerical formats.
    *
    * @param string  $format Custom format string
    * @param integer $ifmt   Format index code
    * @access private
    */
    function _storeNumFormat($format, $ifmt)
    {
        $record    = 0x041E;                      // Record identifier

        if ($this->_BIFF_version == 0x0600) {
			//$numberFormatString = $this->_writeUnicodeDataLong($format);
			$numberFormatString = PHPExcel_Shared_String::UTF8toBIFF8UnicodeLong($format);
            $length    = 2 + strlen($numberFormatString);      // Number of bytes to follow
        } elseif ($this->_BIFF_version == 0x0500) {
            $length    = 3 + strlen($format);      // Number of bytes to follow
        }


        $header    = pack("vv", $record, $length);
        if ($this->_BIFF_version == 0x0600) {
            $data      = pack("v", $ifmt) .  $numberFormatString;
            $this->_append($header . $data);
        } elseif ($this->_BIFF_version == 0x0500) {
            $cch       = strlen($format);             // Length of format string
            $data      = pack("vC", $ifmt, $cch);
            $this->_append($header . $data . $format);
        }
    }

    /**
    * Write DATEMODE record to indicate the date system in use (1904 or 1900).
    *
    * @access private
    */
    function _storeDatemode()
    {
        $record    = 0x0022;         // Record identifier
        $length    = 0x0002;         // Bytes to follow

        $f1904     = (PHPExcel_Shared_Date::getExcelCalendar() == PHPExcel_Shared_Date::CALENDAR_MAC_1904) ?
			1 : 0;   // Flag for 1904 date system

        $header    = pack("vv", $record, $length);
        $data      = pack("v", $f1904);
        $this->_append($header . $data);
    }


    /**
    * Write BIFF record EXTERNCOUNT to indicate the number of external sheet
    * references in the workbook.
    *
    * Excel only stores references to external sheets that are used in NAME.
    * The workbook NAME record is required to define the print area and the repeat
    * rows and columns.
    *
    * A similar method is used in Worksheet.php for a slightly different purpose.
    *
    * @param integer $cxals Number of external references
    * @access private
    */
    function _storeExterncount($cxals)
    {
        $record   = 0x0016;          // Record identifier
        $length   = 0x0002;          // Number of bytes to follow

        $header   = pack("vv", $record, $length);
        $data     = pack("v",  $cxals);
        $this->_append($header . $data);
    }


    /**
    * Writes the Excel BIFF EXTERNSHEET record. These references are used by
    * formulas. NAME record is required to define the print area and the repeat
    * rows and columns.
    *
    * A similar method is used in Worksheet.php for a slightly different purpose.
    *
    * @param string $sheetname Worksheet name
    * @access private
    */
    function _storeExternsheet($sheetname)
    {
        $record      = 0x0017;                     // Record identifier
        $length      = 0x02 + strlen($sheetname);  // Number of bytes to follow

        $cch         = strlen($sheetname);         // Length of sheet name
        $rgch        = 0x03;                       // Filename encoding

        $header      = pack("vv",  $record, $length);
        $data        = pack("CC", $cch, $rgch);
        $this->_append($header . $data . $sheetname);
    }


    /**
    * Store the NAME record in the short format that is used for storing the print
    * area, repeat rows only and repeat columns only.
    *
    * @param integer $index  Sheet index
    * @param integer $type   Built-in name type
    * @param integer $rowmin Start row
    * @param integer $rowmax End row
    * @param integer $colmin Start colum
    * @param integer $colmax End column
    * @access private
    */
    function _storeNameShort($index, $type, $rowmin, $rowmax, $colmin, $colmax)
    {
        $record          = 0x0018;       // Record identifier
        $length          = 0x0024;       // Number of bytes to follow

        $grbit           = 0x0020;       // Option flags
        $chKey           = 0x00;         // Keyboard shortcut
        $cch             = 0x01;         // Length of text name
        $cce             = 0x0015;       // Length of text definition
        $ixals           = $index + 1;   // Sheet index
        $itab            = $ixals;       // Equal to ixals
        $cchCustMenu     = 0x00;         // Length of cust menu text
        $cchDescription  = 0x00;         // Length of description text
        $cchHelptopic    = 0x00;         // Length of help topic text
        $cchStatustext   = 0x00;         // Length of status bar text
        $rgch            = $type;        // Built-in name type

        $unknown03       = 0x3b;
        $unknown04       = 0xffff-$index;
        $unknown05       = 0x0000;
        $unknown06       = 0x0000;
        $unknown07       = 0x1087;
        $unknown08       = 0x8005;

        $header             = pack("vv", $record, $length);
        $data               = pack("v", $grbit);
        $data              .= pack("C", $chKey);
        $data              .= pack("C", $cch);
        $data              .= pack("v", $cce);
        $data              .= pack("v", $ixals);
        $data              .= pack("v", $itab);
        $data              .= pack("C", $cchCustMenu);
        $data              .= pack("C", $cchDescription);
        $data              .= pack("C", $cchHelptopic);
        $data              .= pack("C", $cchStatustext);
        $data              .= pack("C", $rgch);
        $data              .= pack("C", $unknown03);
        $data              .= pack("v", $unknown04);
        $data              .= pack("v", $unknown05);
        $data              .= pack("v", $unknown06);
        $data              .= pack("v", $unknown07);
        $data              .= pack("v", $unknown08);
        $data              .= pack("v", $index);
        $data              .= pack("v", $index);
        $data              .= pack("v", $rowmin);
        $data              .= pack("v", $rowmax);
        $data              .= pack("C", $colmin);
        $data              .= pack("C", $colmax);
        $this->_append($header . $data);
    }


    /**
    * Store the NAME record in the long format that is used for storing the repeat
    * rows and columns when both are specified. This shares a lot of code with
    * _storeNameShort() but we use a separate method to keep the code clean.
    * Code abstraction for reuse can be carried too far, and I should know. ;-)
    *
    * @param integer $index Sheet index
    * @param integer $type  Built-in name type
    * @param integer $rowmin Start row
    * @param integer $rowmax End row
    * @param integer $colmin Start colum
    * @param integer $colmax End column
    * @access private
    */
    function _storeNameLong($index, $type, $rowmin, $rowmax, $colmin, $colmax)
    {
        $record          = 0x0018;       // Record identifier
        $length          = 0x003d;       // Number of bytes to follow
        $grbit           = 0x0020;       // Option flags
        $chKey           = 0x00;         // Keyboard shortcut
        $cch             = 0x01;         // Length of text name
        $cce             = 0x002e;       // Length of text definition
        $ixals           = $index + 1;   // Sheet index
        $itab            = $ixals;       // Equal to ixals
        $cchCustMenu     = 0x00;         // Length of cust menu text
        $cchDescription  = 0x00;         // Length of description text
        $cchHelptopic    = 0x00;         // Length of help topic text
        $cchStatustext   = 0x00;         // Length of status bar text
        $rgch            = $type;        // Built-in name type

        $unknown01       = 0x29;
        $unknown02       = 0x002b;
        $unknown03       = 0x3b;
        $unknown04       = 0xffff-$index;
        $unknown05       = 0x0000;
        $unknown06       = 0x0000;
        $unknown07       = 0x1087;
        $unknown08       = 0x8008;

        $header             = pack("vv",  $record, $length);
        $data               = pack("v", $grbit);
        $data              .= pack("C", $chKey);
        $data              .= pack("C", $cch);
        $data              .= pack("v", $cce);
        $data              .= pack("v", $ixals);
        $data              .= pack("v", $itab);
        $data              .= pack("C", $cchCustMenu);
        $data              .= pack("C", $cchDescription);
        $data              .= pack("C", $cchHelptopic);
        $data              .= pack("C", $cchStatustext);
        $data              .= pack("C", $rgch);
        $data              .= pack("C", $unknown01);
        $data              .= pack("v", $unknown02);
        // Column definition
        $data              .= pack("C", $unknown03);
        $data              .= pack("v", $unknown04);
        $data              .= pack("v", $unknown05);
        $data              .= pack("v", $unknown06);
        $data              .= pack("v", $unknown07);
        $data              .= pack("v", $unknown08);
        $data              .= pack("v", $index);
        $data              .= pack("v", $index);
        $data              .= pack("v", 0x0000);
        $data              .= pack("v", 0x3fff);
        $data              .= pack("C", $colmin);
        $data              .= pack("C", $colmax);
        // Row definition
        $data              .= pack("C", $unknown03);
        $data              .= pack("v", $unknown04);
        $data              .= pack("v", $unknown05);
        $data              .= pack("v", $unknown06);
        $data              .= pack("v", $unknown07);
        $data              .= pack("v", $unknown08);
        $data              .= pack("v", $index);
        $data              .= pack("v", $index);
        $data              .= pack("v", $rowmin);
        $data              .= pack("v", $rowmax);
        $data              .= pack("C", 0x00);
        $data              .= pack("C", 0xff);
        // End of data
        $data              .= pack("C", 0x10);
        $this->_append($header . $data);
    }

    /**
    * Stores the COUNTRY record for localization
    *
    * @return string
    */
    public function writeCountry()
    {
        $record          = 0x008C;    // Record identifier
        $length          = 4;         // Number of bytes to follow

        $header = pack('vv',  $record, $length);
        /* using the same country code always for simplicity */
        $data = pack('vv', $this->_country_code, $this->_country_code);
        //$this->_append($header . $data);
        return $this->writeData($header . $data);
    }

    /**
    * Stores the PALETTE biff record.
    *
    * @access private
    */
    function _storePalette()
    {
        $aref            = $this->_palette;

        $record          = 0x0092;                 // Record identifier
        $length          = 2 + 4 * count($aref);   // Number of bytes to follow
        $ccv             =         count($aref);   // Number of RGB values to follow
        $data = '';                                // The RGB data

        // Pack the RGB data
        foreach ($aref as $color) {
            foreach ($color as $byte) {
                $data .= pack("C",$byte);
            }
        }

        $header = pack("vvv",  $record, $length, $ccv);
        $this->_append($header . $data);
    }

    /**
    * Calculate
    * Handling of the SST continue blocks is complicated by the need to include an
    * additional continuation byte depending on whether the string is split between
    * blocks or whether it starts at the beginning of the block. (There are also
    * additional complications that will arise later when/if Rich Strings are
    * supported).
    *
    * @access private
    */
    function _calculateSharedStringsSizes()
    {
        /* Iterate through the strings to calculate the CONTINUE block sizes.
           For simplicity we use the same size for the SST and CONTINUE records:
           8228 : Maximum Excel97 block size
             -4 : Length of block header
             -8 : Length of additional SST header information
         = 8216
        */
        $continue_limit     = 8208;
        $block_length       = 0;
        $written            = 0;
        $this->_block_sizes = array();
        $continue           = 0;

        foreach (array_keys($this->_str_table) as $string) {
            $string_length = strlen($string);
            $headerinfo    = unpack("vlength/Cencoding", $string);
            $encoding      = $headerinfo["encoding"];
            $split_string  = 0;

            // Block length is the total length of the strings that will be
            // written out in a single SST or CONTINUE block.
            $block_length += $string_length;

            // We can write the string if it doesn't cross a CONTINUE boundary
            if ($block_length < $continue_limit) {
                $written      += $string_length;
                continue;
            }

            // Deal with the cases where the next string to be written will exceed
            // the CONTINUE boundary. If the string is very long it may need to be
            // written in more than one CONTINUE record.
            while ($block_length >= $continue_limit) {

                // We need to avoid the case where a string is continued in the first
                // n bytes that contain the string header information.
                $header_length   = 3; // Min string + header size -1
                $space_remaining = $continue_limit - $written - $continue;


                /* TODO: Unicode data should only be split on char (2 byte)
                boundaries. Therefore, in some cases we need to reduce the
                amount of available
                */
                $align = 0;

                // Only applies to Unicode strings
                if ($encoding == 1) {
                    // Min string + header size -1
                    $header_length = 4;

                    if ($space_remaining > $header_length) {
                        // String contains 3 byte header => split on odd boundary
                        if (!$split_string && $space_remaining % 2 != 1) {
                            --$space_remaining;
                            $align = 1;
                        }
                        // Split section without header => split on even boundary
                        else if ($split_string && $space_remaining % 2 == 1) {
                            --$space_remaining;
                            $align = 1;
                        }

                        $split_string = 1;
                    }
                }

                if ($space_remaining > $header_length) {
                    // Write as much as possible of the string in the current block
                    $written      += $space_remaining;

                    // Reduce the current block length by the amount written
                    $block_length -= $continue_limit - $continue - $align;

                    // Store the max size for this block
                    $this->_block_sizes[] = $continue_limit - $align;

                    // If the current string was split then the next CONTINUE block
                    // should have the string continue flag (grbit) set unless the
                    // split string fits exactly into the remaining space.
                    if ($block_length > 0) {
                        $continue = 1;
                    } else {
                        $continue = 0;
                    }
                } else {
                    // Store the max size for this block
                    $this->_block_sizes[] = $written + $continue;

                    // Not enough space to start the string in the current block
                    $block_length -= $continue_limit - $space_remaining - $continue;
                    $continue = 0;

                }

                // If the string (or substr) is small enough we can write it in the
                // new CONTINUE block. Else, go through the loop again to write it in
                // one or more CONTINUE blocks
                if ($block_length < $continue_limit) {
                    $written = $block_length;
                } else {
                    $written = 0;
                }
            }
        }

        // Store the max size for the last block unless it is empty
        if ($written + $continue) {
            $this->_block_sizes[] = $written + $continue;
        }


        /* Calculate the total length of the SST and associated CONTINUEs (if any).
         The SST record will have a length even if it contains no strings.
         This length is required to set the offsets in the BOUNDSHEET records since
         they must be written before the SST records
        */

        $tmp_block_sizes = array();
        $tmp_block_sizes = $this->_block_sizes;

        $length  = 12;
        if (!empty($tmp_block_sizes)) {
            $length += array_shift($tmp_block_sizes); // SST information
        }
        while (!empty($tmp_block_sizes)) {
            $length += 4 + array_shift($tmp_block_sizes); // add CONTINUE headers
        }

        return $length;
    }

    /**
    * Write all of the workbooks strings into an indexed array.
    * See the comments in _calculate_shared_string_sizes() for more information.
    *
    * The Excel documentation says that the SST record should be followed by an
    * EXTSST record. The EXTSST record is a hash table that is used to optimise
    * access to SST. However, despite the documentation it doesn't seem to be
    * required so we will ignore it.
    *
    * @access private
    */
    public function writeSharedStringsTable()
    {
        $chunk = '';

		$record  = 0x00fc;  // Record identifier
        $length  = 0x0008;  // Number of bytes to follow
        $total   = 0x0000;

        // Iterate through the strings to calculate the CONTINUE block sizes
        $continue_limit = 8208;
        $block_length   = 0;
        $written        = 0;
        $continue       = 0;

        // sizes are upside down
        $tmp_block_sizes = $this->_block_sizes;
//        $tmp_block_sizes = array_reverse($this->_block_sizes);

        // The SST record is required even if it contains no strings. Thus we will
        // always have a length
        //
        if (!empty($tmp_block_sizes)) {
            $length = 8 + array_shift($tmp_block_sizes);
        } else {
            // No strings
            $length = 8;
        }



        // Write the SST block header information
        $header      = pack("vv", $record, $length);
        $data        = pack("VV", $this->_str_total, $this->_str_unique);
        //$this->_append($header . $data);
        $chunk .= $this->writeData($header . $data);




        /* TODO: not good for performance */
        foreach (array_keys($this->_str_table) as $string) {

            $string_length = strlen($string);
            $headerinfo    = unpack("vlength/Cencoding", $string);
            $encoding      = $headerinfo["encoding"];
            $split_string  = 0;

            // Block length is the total length of the strings that will be
            // written out in a single SST or CONTINUE block.
            //
            $block_length += $string_length;


            // We can write the string if it doesn't cross a CONTINUE boundary
            if ($block_length < $continue_limit) {
                //$this->_append($string);
                $chunk .= $this->writeData($string);
                $written += $string_length;
                continue;
            }

            // Deal with the cases where the next string to be written will exceed
            // the CONTINUE boundary. If the string is very long it may need to be
            // written in more than one CONTINUE record.
            //
            while ($block_length >= $continue_limit) {

                // We need to avoid the case where a string is continued in the first
                // n bytes that contain the string header information.
                //
                $header_length   = 3; // Min string + header size -1
                $space_remaining = $continue_limit - $written - $continue;


                // Unicode data should only be split on char (2 byte) boundaries.
                // Therefore, in some cases we need to reduce the amount of available
                // space by 1 byte to ensure the correct alignment.
                $align = 0;

                // Only applies to Unicode strings
                if ($encoding == 1) {
                    // Min string + header size -1
                    $header_length = 4;

                    if ($space_remaining > $header_length) {
                        // String contains 3 byte header => split on odd boundary
                        if (!$split_string && $space_remaining % 2 != 1) {
                            --$space_remaining;
                            $align = 1;
                        }
                        // Split section without header => split on even boundary
                        else if ($split_string && $space_remaining % 2 == 1) {
                            --$space_remaining;
                            $align = 1;
                        }

                        $split_string = 1;
                    }
                }


                if ($space_remaining > $header_length) {
                    // Write as much as possible of the string in the current block
                    $tmp = substr($string, 0, $space_remaining);
                    //$this->_append($tmp);
                    $chunk .= $this->writeData($tmp);

                    // The remainder will be written in the next block(s)
                    $string = substr($string, $space_remaining);

                    // Reduce the current block length by the amount written
                    $block_length -= $continue_limit - $continue - $align;

                    // If the current string was split then the next CONTINUE block
                    // should have the string continue flag (grbit) set unless the
                    // split string fits exactly into the remaining space.
                    //
                    if ($block_length > 0) {
                        $continue = 1;
                    } else {
                        $continue = 0;
                    }
                } else {
                    // Not enough space to start the string in the current block
                    $block_length -= $continue_limit - $space_remaining - $continue;
                    $continue = 0;
                }

                // Write the CONTINUE block header
                if (!empty($this->_block_sizes)) {
                    $record  = 0x003C;
                    $length  = array_shift($tmp_block_sizes);
                    $header  = pack('vv', $record, $length);
                    if ($continue) {
                        $header .= pack('C', $encoding);
                    }
                    //$this->_append($header);
                    $chunk .= $this->writeData($header);
                }

                // If the string (or substr) is small enough we can write it in the
                // new CONTINUE block. Else, go through the loop again to write it in
                // one or more CONTINUE blocks
                //
                if ($block_length < $continue_limit) {
                    //$this->_append($string);
                    $chunk .= $this->writeData($string);
                    $written = $block_length;
                } else {
                    $written = 0;
                }
            }
        }
		return $chunk;
    }
}
