<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2009 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel_Writer
 * @copyright  Copyright (c) 2006 - 2009 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.6.5, 2009-01-05
 */


/** PHPExcel_IWriter */
require_once 'PHPExcel/Writer/IWriter.php';

/** PHPExcel_Cell */
require_once 'PHPExcel/Cell.php';

/** PHPExcel_RichText */
require_once 'PHPExcel/RichText.php';

/** PHPExcel_Shared_Drawing */
require_once 'PHPExcel/Shared/Drawing.php';

/** PHPExcel_Shared_String */
require_once 'PHPExcel/Shared/String.php';

/** PHPExcel_HashTable */
require_once 'PHPExcel/HashTable.php';


/**
 * PHPExcel_Writer_HTML
 *
 * @category   PHPExcel
 * @package    PHPExcel_Writer
 * @copyright  Copyright (c) 2006 - 2009 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Writer_HTML implements PHPExcel_Writer_IWriter {
	/**
	 * PHPExcel object
	 *
	 * @var PHPExcel
	 */
	private $_phpExcel;

	/**
	 * Sheet index to write
	 *
	 * @var int
	 */
	private $_sheetIndex;

	/**
	 * Pre-calculate formulas
	 *
	 * @var boolean
	 */
	private $_preCalculateFormulas = true;

	/**
	 * Images root
	 *
	 * @var string
	 */
	private $_imagesRoot = '.';

	/**
	 * Create a new PHPExcel_Writer_HTML
	 *
	 * @param 	PHPExcel	$phpExcel	PHPExcel object
	 */
	public function __construct(PHPExcel $phpExcel) {
		$this->_phpExcel = $phpExcel;
		$this->_sheetIndex = 0;
		$this->_imagesRoot = '.';
	}

	/**
	 * Save PHPExcel to file
	 *
	 * @param 	string 		$pFileName
	 * @throws 	Exception
	 */
	public function save($pFilename = null) {
		$saveArrayReturnType = PHPExcel_Calculation::getArrayReturnType();
		PHPExcel_Calculation::setArrayReturnType(PHPExcel_Calculation::RETURN_ARRAY_AS_VALUE);

		// Open file
		$fileHandle = fopen($pFilename, 'w');
		if ($fileHandle === false) {
			throw new Exception("Could not open file $pFilename for writing.");
		}

		// Write headers
		fwrite($fileHandle, $this->generateHTMLHeader(true));

		// Write data
		fwrite($fileHandle, $this->generateSheetData());

		// Write footer
		fwrite($fileHandle, $this->generateHTMLFooter());

		// Close file
		fclose($fileHandle);

		PHPExcel_Calculation::setArrayReturnType($saveArrayReturnType);
	}

	/**
	 * Map VAlign
	 */
	private function _mapVAlign($vAlign) {
		switch ($vAlign) {
			case PHPExcel_Style_Alignment::VERTICAL_BOTTOM: return 'bottom';
			case PHPExcel_Style_Alignment::VERTICAL_TOP: return 'top';
			case PHPExcel_Style_Alignment::VERTICAL_CENTER:
			case PHPExcel_Style_Alignment::VERTICAL_JUSTIFY: return 'middle';
			default: return ' baseline';
		}
	}

	/**
	 * Map HAlign
	 */
	private function _mapHAlign($hAlign) {
		switch ($hAlign) {
			case PHPExcel_Style_Alignment::HORIZONTAL_GENERAL:
			case PHPExcel_Style_Alignment::HORIZONTAL_LEFT: return 'left';
			case PHPExcel_Style_Alignment::HORIZONTAL_RIGHT: return 'right';
			case PHPExcel_Style_Alignment::HORIZONTAL_CENTER: return 'center';
			case PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY: return 'justify';
			default: return ' baseline';
		}
	}

	/**
	 * Map border style
	 */
	private function _mapBorderStyle($borderStyle) {
		switch ($borderStyle) {
			case PHPExcel_Style_Border::BORDER_NONE: return '0px';
			case PHPExcel_Style_Border::BORDER_DASHED: return '1px dashed';
			case PHPExcel_Style_Border::BORDER_DOTTED: return '1px dotted';
			case PHPExcel_Style_Border::BORDER_THICK: return '2px solid';
			default: return '1px solid'; // map others to thin
		}
	}

	/**
	 * Get sheet index
	 *
	 * @return int
	 */
	public function getSheetIndex() {
		return $this->_sheetIndex;
	}

	/**
	 * Set sheet index
	 *
	 * @param	int		$pValue		Sheet index
	 */
	public function setSheetIndex($pValue = 0) {
		$this->_sheetIndex = $pValue;
	}

	/**
	 * Write all sheets (resets sheetIndex to NULL)
	 */
	public function writeAllSheets() {
		$this->_sheetIndex = null;
	}

	/**
	 * Generate HTML header
	 *
	 * @param	boolean		$pIncludeStyles		Include styles?
	 * @return	string
	 * @throws Exception
	 */
	public function generateHTMLHeader($pIncludeStyles = false) {
		// PHPExcel object known?
		if (is_null($this->_phpExcel)) {
			throw new Exception('Internal PHPExcel object not set to an instance of an object.');
		}

		// Construct HTML
		$html = '';
		$html .= '<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">' . "\r\n";
		$html .= '<!-- Generated by PHPExcel - http://www.phpexcel.net -->' . "\r\n";
		$html .= '<html>' . "\r\n";
		$html .= '  <head>' . "\r\n";
		$html .= '    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">' . "\r\n";
		$html .= '    <title>' . htmlspecialchars($this->_phpExcel->getProperties()->getTitle()) . '</title>' . "\r\n";
		if ($pIncludeStyles) {
			$html .= $this->generateStyles(true);
		}
		$html .= '  </head>' . "\r\n";
		$html .= '' . "\r\n";
		$html .= '  <body>' . "\r\n";

		// Return
		return $html;
	}

	/**
	 * Generate sheet data
	 *
	 * @return	string
	 * @throws Exception
	 */
	public function generateSheetData() {
		// PHPExcel object known?
		if (is_null($this->_phpExcel)) {
			throw new Exception('Internal PHPExcel object not set to an instance of an object.');
		}

		// Fetch sheets
		$sheets = array();
		if (is_null($this->_sheetIndex)) {
			$sheets = $this->_phpExcel->getAllSheets();
		} else {
			$sheets[] = $this->_phpExcel->getSheet($this->_sheetIndex);
		}

		// Construct HTML
		$html = '';

		// Loop all sheets
		foreach ($sheets as $sheet) {
			// Calculate hash code
			$hashCode = $sheet->getHashCode();

			// Get cell collection
			$cellCollection = $sheet->getCellCollection();

			// Write table header
			$html .= $this->_generateTableHeader($hashCode);

	    	// Get worksheet dimension
	    	$dimension = explode(':', $sheet->calculateWorksheetDimension());
	    	$dimension[0] = PHPExcel_Cell::coordinateFromString($dimension[0]);
	    	$dimension[0][0] = PHPExcel_Cell::columnIndexFromString($dimension[0][0]) - 1;
	    	$dimension[1] = PHPExcel_Cell::coordinateFromString($dimension[1]);
	    	$dimension[1][0] = PHPExcel_Cell::columnIndexFromString($dimension[1][0]) - 1;

	    	// Loop trough cells
	    	$rowData = null;
	    	for ($row = $dimension[0][1]; $row <= $dimension[1][1]; ++$row) {
				// Start a new row
				$rowData = array();

				// Loop trough columns
	    		for ($column = $dimension[0][0]; $column <= $dimension[1][0]; ++$column) {
	    			// Cell exists?
	    			if ($sheet->cellExistsByColumnAndRow($column, $row)) {
	    				$rowData[$column] = $sheet->getCellByColumnAndRow($column, $row);
	    			} else {
	    				$rowData[$column] = '';
	    			}
	    		}

	    		// Write row
				$html .= $this->_generateRow($sheet, $rowData, $row - 1);
	    	}

			// Write table footer
			$html .= $this->_generateTableFooter();
		}

		// Return
		return $html;
	}

	/**
	 * Generate image tag in cell
	 *
	 * @param	PHPExcel_Worksheet 	$pSheet			PHPExcel_Worksheet
	 * @param	string				$coordinates	Cell coordinates
	 * @return	string
	 * @throws	Exception
	 */
	private function _writeImageTagInCell(PHPExcel_Worksheet $pSheet, $coordinates) {
		// Construct HTML
		$html = '';

		// Write images
		foreach ($pSheet->getDrawingCollection() as $drawing) {
			if ($drawing instanceof PHPExcel_Worksheet_BaseDrawing) {
				if ($drawing->getCoordinates() == $coordinates) {
					$filename = $drawing->getPath();

					// Strip off eventual '.'
					if (substr($filename, 0, 1) == '.') {
						$filename = substr($filename, 1);
					}

					// Prepend images root
					$filename = $this->getImagesRoot() . $filename;

					// Strip off eventual '.'
					if (substr($filename, 0, 1) == '.' && substr($filename, 0, 2) != './') {
						$filename = substr($filename, 1);
					}

					// Convert UTF8 data to PCDATA
					$filename = htmlspecialchars($filename);

					$html .= "\r\n";
					$html .= '        <img  style="position: relative; left: ' . $drawing->getOffsetX() . 'px; top: ' . $drawing->getOffsetY() . 'px; width: ' . $drawing->getWidth() . 'px; height: ' . $drawing->getHeight() . 'px;" src="' . $filename . '" border="0">' . "\r\n";
				}
			}
		}

		// Return
		return $html;
	}

	/**
	 * Generate CSS styles
	 *
	 * @param	boolean	$generateSurroundingHTML	Generate surrounding HTML tags? (<style> and </style>)
	 * @return	string
	 * @throws	Exception
	 */
	public function generateStyles($generateSurroundingHTML = true) {
		// PHPExcel object known?
		if (is_null($this->_phpExcel)) {
			throw new Exception('Internal PHPExcel object not set to an instance of an object.');
		}

		// Construct HTML
		$html = '';

		// Start styles
		if ($generateSurroundingHTML) {
			$html .= '    <style type="text/css">' . "\r\n";
			$html .= '      html {' . "\r\n";
			$html .= '        font-family: Calibri, Arial, Helvetica, sans-serif;' . "\r\n";
			$html .= '        font-size: 10pt;' . "\r\n";
			$html .= '        background-color: white;' . "\r\n";
			$html .= '      }' . "\r\n";
		}

		// Fetch sheets
		$sheets = array();
		if (is_null($this->_sheetIndex)) {
			$sheets = $this->_phpExcel->getAllSheets();
		} else {
			$sheets[] = $this->_phpExcel->getSheet($this->_sheetIndex);
		}

		// Write styles per sheet
		foreach ($sheets as $sheet) {
			// Calculate hash code
			$hashCode = $sheet->getHashCode();

			// Write styles
			$html .= '      table.sheet' . $hashCode . ', table.sheet' . $hashCode . ' td {' . "\r\n";
			if ($sheet->getShowGridlines()) {
				$html .= '        border: 1px dotted black;' . "\r\n";
			}
			$html .= '        page-break-after: always;' . "\r\n";
			$html .= '      }' . "\r\n";

			// Default column width
			$columnDimension = $sheet->getDefaultColumnDimension();

			$html .= '      table.sheet' . $hashCode . ' td {' . "\r\n";
			$html .= '        width: ' . PHPExcel_Shared_Drawing::cellDimensionToPixels($columnDimension->getWidth()) . 'px;' . "\r\n";
			if ($columnDimension->getVisible() === false) {
				$html .= '        display: none;' . "\r\n";
				$html .= '        visibility: hidden;' . "\r\n";
			}
			$html .= '      }' . "\r\n";

			// Calculate column widths
			$sheet->calculateColumnWidths();
			foreach ($sheet->getColumnDimensions() as $columnDimension) {
				$column = PHPExcel_Cell::columnIndexFromString($columnDimension->getColumnIndex()) - 1;

				$html .= '      table.sheet' . $hashCode . ' td.column' . $column  . ' {' . "\r\n";
				$html .= '        width: ' . PHPExcel_Shared_Drawing::cellDimensionToPixels($columnDimension->getWidth()) . 'px;' . "\r\n";
				if ($columnDimension->getVisible() === false) {
					$html .= '        display: none;' . "\r\n";
					$html .= '        visibility: hidden;' . "\r\n";
				}
				$html .= '      }' . "\r\n";
			}

			// Default row height
			$rowDimension = $sheet->getDefaultRowDimension();

			$html .= '      table.sheet' . $hashCode . ' tr {' . "\r\n";
			// height is disproportionately large
			$px_height = round( PHPExcel_Shared_Drawing::cellDimensionToPixels($rowDimension->getRowHeight()) / 12 );
			$html .= '        height: ' . $px_height . 'px;' . "\r\n";
			if ($rowDimension->getVisible() === false) {
				$html .= '        display: none;' . "\r\n";
				$html .= '        visibility: hidden;' . "\r\n";
			}
			$html .= '      }' . "\r\n";

			// Calculate row heights
			foreach ($sheet->getRowDimensions() as $rowDimension) {
				$html .= '      table.sheet' . $hashCode . ' tr.row' . ($rowDimension->getRowIndex() - 1)  . ' {' . "\r\n";
				// height is disproportionately large
				$px_height = round( PHPExcel_Shared_Drawing::cellDimensionToPixels($rowDimension->getRowHeight()) / 12 );
				$html .= '        height: ' . $px_height . 'px;' . "\r\n";
				if ($rowDimension->getVisible() === false) {
					$html .= '        display: none;' . "\r\n";
					$html .= '        visibility: hidden;' . "\r\n";
				}
				$html .= '      }' . "\r\n";
			}

			// Calculate cell style hashes
			$cellStyleHashes = new PHPExcel_HashTable();
			$aStyles = $sheet->getStyles();
			$cellStyleHashes->addFromSource( $aStyles );
			$addedStyles = array();
			foreach ($aStyles as $style) {
				if(isset($addedStyles[$style->getHashIndex()])) continue;
				$html .= $this->_createCSSStyle( $style );
				$addedStyles[$style->getHashIndex()] = true;
			}
		}

		// End styles
		if ($generateSurroundingHTML) {
			$html .= '    </style>' . "\r\n";
		}

		// Return
		return $html;
	}

	/**
	 * Create CSS style
	 *
	 * @param	PHPExcel_Style 		$pStyle			PHPExcel_Style
	 * @return	string
	 */
	private function _createCSSStyle(PHPExcel_Style $pStyle) {
		// Construct HTML
		$html = '';

		// Create CSS
		$html .= '      .style' . $pStyle->getHashIndex() . ' {' . "\r\n";
		$html .= $this->_createCSSStyleAlignment($pStyle->getAlignment());
		$html .= $this->_createCSSStyleFont($pStyle->getFont());
		$html .= $this->_createCSSStyleBorders($pStyle->getBorders());
		$html .= $this->_createCSSStyleFill($pStyle->getFill());
		$html .= '      }' . "\r\n";

		// Return
		return $html;
	}

	/**
	 * Create CSS style (PHPExcel_Style_Alignment)
	 *
	 * @param	PHPExcel_Style_Alignment 		$pStyle			PHPExcel_Style_Alignment
	 * @return	string
	 */
	private function _createCSSStyleAlignment(PHPExcel_Style_Alignment $pStyle) {
		// Construct HTML
		$html = '';

		// Create CSS
		$html .= '        vertical-align: ' 	. $this->_mapVAlign($pStyle->getVertical()) . ';' . "\r\n";
		$html .= '        text-align: ' 		. $this->_mapHAlign($pStyle->getHorizontal()) . ';' . "\r\n";

		// Return
		return $html;
	}

	/**
	 * Create CSS style (PHPExcel_Style_Font)
	 *
	 * @param	PHPExcel_Style_Font 		$pStyle			PHPExcel_Style_Font
	 * @return	string
	 */
	private function _createCSSStyleFont(PHPExcel_Style_Font $pStyle) {
		// Construct HTML
		$html = '';

		// Create CSS
		if ($pStyle->getBold()) {
			$html .= '        font-weight: bold;' . "\r\n";
		}
		if ($pStyle->getUnderline() != PHPExcel_Style_Font::UNDERLINE_NONE && $pStyle->getStriketrough()) {
			$html .= '        text-decoration: underline line-through;' . "\r\n";
		} else if ($pStyle->getUnderline() != PHPExcel_Style_Font::UNDERLINE_NONE) {
			$html .= '        text-decoration: underline;' . "\r\n";
		} else if ($pStyle->getStriketrough()) {
			$html .= '        text-decoration: line-through;' . "\r\n";
		}
		if ($pStyle->getItalic()) {
			$html .= '        font-style: italic;' . "\r\n";
		}

		$html .= '        color: ' 				. '#' . $pStyle->getColor()->getRGB() . ';' . "\r\n";
		$html .= '        font-family: ' 		. '\'' . $pStyle->getName() . '\';' . "\r\n";
		$html .= '        font-size: ' 			. $pStyle->getSize() . 'pt;' . "\r\n";

		// Return
		return $html;
	}

	/**
	 * Create CSS style (PHPExcel_Style_Borders)
	 *
	 * @param	PHPExcel_Style_Borders 		$pStyle			PHPExcel_Style_Borders
	 * @return	string
	 */
	private function _createCSSStyleBorders(PHPExcel_Style_Borders $pStyle) {
		// Construct HTML
		$html = '';

		// Create CSS
		$html .= '        border-bottom: ' 		. $this->_createCSSStyleBorder($pStyle->getBottom()) . ';' . "\r\n";
		$html .= '        border-top: ' 		. $this->_createCSSStyleBorder($pStyle->getTop()) . ';' . "\r\n";
		$html .= '        border-left: ' 		. $this->_createCSSStyleBorder($pStyle->getLeft()) . ';' . "\r\n";
		$html .= '        border-right: ' 		. $this->_createCSSStyleBorder($pStyle->getRight()) . ';' . "\r\n";

		// Return
		return $html;
	}

	/**
	 * Create CSS style (PHPExcel_Style_Border)
	 *
	 * @param	PHPExcel_Style_Border		$pStyle			PHPExcel_Style_Border
	 * @return	string
	 */
	private function _createCSSStyleBorder(PHPExcel_Style_Border $pStyle) {
		// Construct HTML
		$html = '';

		// Create CSS
		$html .= $this->_mapBorderStyle($pStyle->getBorderStyle()) . ' #' . $pStyle->getColor()->getRGB();

		// Return
		return $html;
	}

	/**
	 * Create CSS style (PHPExcel_Style_Fill)
	 *
	 * @param	PHPExcel_Style_Fill		$pStyle			PHPExcel_Style_Fill
	 * @return	string
	 */
	private function _createCSSStyleFill(PHPExcel_Style_Fill $pStyle) {
		// Construct HTML
		$html = '';

		// Create CSS
		$html .= '        background-color: ' 	. '#' . $pStyle->getStartColor()->getRGB() . ';' . "\r\n";

		// Return
		return $html;
	}

	/**
	 * Generate HTML footer
	 */
	public function generateHTMLFooter() {
		// Construct HTML
		$html = '';
		$html .= '  </body>' . "\r\n";
		$html .= '</html>' . "\r\n";

		// Return
		return $html;
	}

	/**
	 * Generate table header
	 *
	 * @param 	string	$pIdentifier	Identifier for the table
	 * @return	string
	 * @throws	Exception
	 */
	private function _generateTableHeader($pIdentifier = '') {
		// Construct HTML
		$html = '';
		$html .= '    <table border="0" cellpadding="0" cellspacing="0" class="sheet' . $pIdentifier . '">' . "\r\n";

		// Return
		return $html;
	}

	/**
	 * Generate table footer
	 *
	 * @throws	Exception
	 */
	private function _generateTableFooter() {
		// Construct HTML
		$html = '';
		$html .= '    </table>' . "\r\n";

		// Return
		return $html;
	}

	/**
	 * Generate row
	 *
	 * @param	PHPExcel_Worksheet 	$pSheet			PHPExcel_Worksheet
	 * @param	array				$pValues		Array containing cells in a row
	 * @param	int					$pRow			Row number
	 * @return	string
	 * @throws	Exception
	 */
	private function _generateRow(PHPExcel_Worksheet $pSheet, $pValues = null, $pRow = 0) {
		if (is_array($pValues)) {
			// Construct HTML
			$html = '';

			// Write row start
			$html .= '        <tr class="row' . $pRow . '">' . "\r\n";

			// Write cells
			$colNum = 0;
			foreach ($pValues as $cell) {
				$cellData = '&nbsp;';
				$cssClass = 'column' . $colNum;
				$colSpan = 1;
				$rowSpan = 1;
				$writeCell = true;	// Write cell

				// PHPExcel_Cell
				if ($cell instanceof PHPExcel_Cell) {
					// Value
					if ($cell->getValue() instanceof PHPExcel_RichText) {
						// Loop trough rich text elements
						$elements = $cell->getValue()->getRichTextElements();
						foreach ($elements as $element) {
							// Rich text start?
							if ($element instanceof PHPExcel_RichText_Run) {
								$cellData .= '<span style="' .
									str_replace("\r\n", '',
										$this->_createCSSStyleFont($element->getFont())
									) . '">';

								if ($element->getFont()->getSuperScript()) {
									$cellData .= '<sup>';
								} else if ($element->getFont()->getSubScript()) {
									$cellData .= '<sub>';
								}
							}

							// Convert UTF8 data to PCDATA
							$cellText = $element->getText();
							$cellData .= htmlspecialchars($cellText);

							if ($element instanceof PHPExcel_RichText_Run) {
								if ($element->getFont()->getSuperScript()) {
									$cellData .= '</sup>';
								} else if ($element->getFont()->getSubScript()) {
									$cellData .= '</sub>';
								}

								$cellData .= '</span>';
							}
						}
					} else {
						if ($this->_preCalculateFormulas) {
							$cellData = PHPExcel_Style_NumberFormat::toFormattedString(
								$cell->getCalculatedValue(),
								$pSheet->getstyle( $cell->getCoordinate() )->getNumberFormat()->getFormatCode()
							);
						} else {
							$cellData = PHPExcel_Style_NumberFormat::ToFormattedString(
								$cell->getValue(),
								$pSheet->getstyle( $cell->getCoordinate() )->getNumberFormat()->getFormatCode()
							);
						}

						// Convert UTF8 data to PCDATA
						$cellData = htmlspecialchars($cellData);
					}

					// Check value
					if ($cellData == '') {
						$cellData = '&nbsp;';
					}

					// Extend CSS class?
					if (array_key_exists($cell->getCoordinate(), $pSheet->getStyles())) {
						$cssClass .= ' style' . $pSheet->getStyle($cell->getCoordinate())->getHashIndex();
					}
				} else {
					$cell = new PHPExcel_Cell(
						PHPExcel_Cell::stringFromColumnIndex($colNum),
						($pRow + 1),
						'',
						null,
						null
					);
				}

				// Hyperlink?
				if ($cell->hasHyperlink() && !$cell->getHyperlink()->isInternal()) {
					$cellData = '<a href="' . htmlspecialchars($cell->getHyperlink()->getUrl()) . '" title="' . htmlspecialchars($cell->getHyperlink()->getTooltip()) . '">' . $cellData . '</a>';
				}

				// Column/rowspan
				foreach ($pSheet->getMergeCells() as $cells) {
					if ($cell->isInRange($cells)) {
						list($first, ) = PHPExcel_Cell::splitRange($cells);

						if ($first == $cell->getCoordinate()) {
							list($colSpan, $rowSpan) = PHPExcel_Cell::rangeDimension($cells);
						} else {
							$writeCell = false;
						}

						break;
					}
				}

				// Write
				if ($writeCell) {
					// Column start
					$html .= '          <td';
						$html .= ' class="' . $cssClass . '"';
						if ($colSpan > 1) {
							$html .= ' colspan="' . $colSpan . '"';
						}
						if ($rowSpan > 1) {
							$html .= ' rowspan="' . $rowSpan . '"';
						}
					$html .= '>';

					// Image?
					$html .= $this->_writeImageTagInCell($pSheet, $cell->getCoordinate());

					// Cell data
					$html .= $cellData;

					// Column end
					$html .= '</td>' . "\r\n";
				}

				// Next column
				++$colNum;
			}

			// Write row end
			$html .= '        </tr>' . "\r\n";

			// Return
			return $html;
		} else {
			throw new Exception("Invalid parameters passed.");
		}
	}


    /**
     * Get Pre-Calculate Formulas
     *
     * @return boolean
     */
    public function getPreCalculateFormulas() {
    	return $this->_preCalculateFormulas;
    }

    /**
     * Set Pre-Calculate Formulas
     *
     * @param boolean $pValue	Pre-Calculate Formulas?
     */
    public function setPreCalculateFormulas($pValue = true) {
    	$this->_preCalculateFormulas = $pValue;
    }

    /**
     * Get images root
     *
     * @return string
     */
    public function getImagesRoot() {
    	return $this->_imagesRoot;
    }

    /**
     * Set images root
     *
     * @param string $pValue
     */
    public function setImagesRoot($pValue = '.') {
    	$this->_imagesRoot = $pValue;
    }
}
