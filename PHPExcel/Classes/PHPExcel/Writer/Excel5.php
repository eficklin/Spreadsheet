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
 * @package    PHPExcel_Writer_Excel5
 * @copyright  Copyright (c) 2006 - 2009 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	1.6.5, 2009-01-05
 */


/** PHPExcel_IWriter */
require_once 'PHPExcel/Writer/IWriter.php';

/** PHPExcel_Cell */
require_once 'PHPExcel/Cell.php';

/** PHPExcel_Writer_Excel5_Workbook */
require_once 'PHPExcel/Writer/Excel5/Workbook.php';

/** PHPExcel_RichText */
require_once 'PHPExcel/RichText.php';

/** PHPExcel_HashTable */
require_once 'PHPExcel/HashTable.php';


/**
 * PHPExcel_Writer_Excel5
 *
 * @category   PHPExcel
 * @package    PHPExcel_Writer_Excel5
 * @copyright  Copyright (c) 2006 - 2009 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Writer_Excel5 implements PHPExcel_Writer_IWriter {
	/**
	 * PHPExcel object
	 *
	 * @var PHPExcel
	 */
	private $_phpExcel;

	/**
	 * Temporary storage directory
	 *
	 * @var string
	 */
	private $_tempDir = '';

	/**
	 * Color cache
	 */
	private $_colors = array();

	/**
	 * Create a new PHPExcel_Writer_Excel5
	 *
	 * @param	PHPExcel	$phpExcel	PHPExcel object
	 */
	public function __construct(PHPExcel $phpExcel) {
		$this->_phpExcel	= $phpExcel;
		$this->_tempDir		= '';
		$this->_colors		= array();
	}

	/**
	 * Save PHPExcel to file
	 *
	 * @param	string		$pFileName
	 * @throws	Exception
	 */
	public function save($pFilename = null) {

		// check for iconv support
		if (!function_exists('iconv')) {
			throw new Exception("Cannot write .xls file without PHP support for iconv");
		}

		$this->_colors		= array();

		$phpExcel = $this->_phpExcel;
		$workbook = new PHPExcel_Writer_Excel5_Workbook($pFilename, $phpExcel);
		$workbook->setVersion(8);

		// Set temp dir
		if ($this->_tempDir != '') {
			$workbook->setTempDir($this->_tempDir);
		}

		$saveDateReturnType = PHPExcel_Calculation_Functions::getReturnDateType();
		PHPExcel_Calculation_Functions::setReturnDateType(PHPExcel_Calculation_Functions::RETURNDATE_EXCEL);

		// Add empty sheets
		foreach ($phpExcel->getSheetNames() as $sheetIndex => $sheetName) {
			$phpSheet  = $phpExcel->getSheet($sheetIndex);
			$worksheet = $workbook->addWorksheet($sheetName, $phpSheet);
		}
		$allWorksheets = $workbook->worksheets();

		$formats = array();

		// Add full sheet data
		foreach ($phpExcel->getSheetNames() as $sheetIndex => $sheetName) {
			$phpSheet  = $phpExcel->getSheet($sheetIndex);
			$worksheet = $allWorksheets[$sheetIndex];
			
			// Default style
			$emptyStyle = $phpSheet->getDefaultStyle();

			$aStyles = $phpSheet->getStyles();
			
			// Calculate cell style hashes
			$cellStyleHashes = new PHPExcel_HashTable();
			$aStyles = $phpSheet->getStyles();
			$cellStyleHashes->addFromSource( $aStyles );

			$addedStyles = array();
			foreach ($aStyles as $style) {
				$styleHashIndex = $style->getHashIndex();

				if(isset($addedStyles[$styleHashIndex])) continue;
				
				$formats[$styleHashIndex] = $workbook->addFormat(array(
					'HAlign' => $style->getAlignment()->getHorizontal(),
					'VAlign' => $this->_mapVAlign($style->getAlignment()->getVertical()),
					'TextRotation' => $style->getAlignment()->getTextRotation(),

					'Bold' => $style->getFont()->getBold(),
					'FontFamily' => $style->getFont()->getName(),
					'Color' => $this->_addColor($workbook, $style->getFont()->getColor()->getRGB()),
					'Underline' => $this->_mapUnderline($style->getFont()->getUnderline()),
					'Size' => $style->getFont()->getSize(),
					//~ 'Script' => $style->getSuperscript(),

					'NumFormat' => $style->getNumberFormat()->getFormatCode(),

					'Bottom' => $this->_mapBorderStyle($style->getBorders()->getBottom()->getBorderStyle()),
					'Top' => $this->_mapBorderStyle($style->getBorders()->getTop()->getBorderStyle()),
					'Left' => $this->_mapBorderStyle($style->getBorders()->getLeft()->getBorderStyle()),
					'Right' => $this->_mapBorderStyle($style->getBorders()->getRight()->getBorderStyle()),
					'BottomColor' => $this->_addColor($workbook, $style->getBorders()->getBottom()->getColor()->getRGB()),
					'TopColor' => $this->_addColor($workbook, $style->getBorders()->getTop()->getColor()->getRGB()),
					'RightColor' => $this->_addColor($workbook, $style->getBorders()->getRight()->getColor()->getRGB()),
					'LeftColor' => $this->_addColor($workbook, $style->getBorders()->getLeft()->getColor()->getRGB()),

					'FgColor' => $this->_addColor($workbook, $style->getFill()->getStartColor()->getRGB()),
					'BgColor' => $this->_addColor($workbook, $style->getFill()->getEndColor()->getRGB()),
					'Pattern' => $this->_mapFillType($style->getFill()->getFillType()),

				));
				if ($style->getAlignment()->getWrapText()) {
					$formats[$styleHashIndex]->setTextWrap();
				}
				$formats[$styleHashIndex]->setIndent($style->getAlignment()->getIndent());
				if ($style->getAlignment()->getShrinkToFit()) {
					$formats[$styleHashIndex]->setShrinkToFit();
				}
				if ($style->getFont()->getItalic()) {
					$formats[$styleHashIndex]->setItalic();
				}
				if ($style->getFont()->getStriketrough()) {
					$formats[$styleHashIndex]->setStrikeOut();
				}
				if ($style->getProtection()->getLocked() == PHPExcel_Style_Protection::PROTECTION_UNPROTECTED) {
					$formats[$styleHashIndex]->setUnlocked();
				}
				if ($style->getProtection()->getHidden() == PHPExcel_Style_Protection::PROTECTION_PROTECTED) {
					$formats[$styleHashIndex]->setHidden();
				}
				
				$addedStyles[$style->getHashIndex()] = true;
			}

			// Active sheet
			if ($sheetIndex == $phpExcel->getActiveSheetIndex()) {
				$worksheet->activate();
			}

			// initialize first, last, row and column index, needed for DIMENSION record
			$firstRowIndex = 0;
			$lastRowIndex = -1;
			$firstColumnIndex = 0;
			$lastColumnIndex = -1;

			foreach ($phpSheet->getCellCollection() as $cell) {
				$row = $cell->getRow() - 1;
				$column = PHPExcel_Cell::columnIndexFromString($cell->getColumn()) - 1;

				// Don't break Excel!
				if ($row + 1 >= 65569) {
					break;
				}

				$firstRowIndex = min($firstRowIndex, $row);
				$lastRowIndex = max($lastRowIndex, $row);
				$firstColumnIndex = min($firstColumnIndex, $column);
				$lastColumnIndex = max($lastColumnIndex, $column);

				$style = $emptyStyle;
				if (isset($aStyles[$cell->getCoordinate()])) {
					$style = $aStyles[$cell->getCoordinate()];
				}
				$styleHashIndex = $style->getHashIndex();

				// Write cell value
				if ($cell->getValue() instanceof PHPExcel_RichText) {
					$worksheet->write($row, $column, $cell->getValue()->getPlainText(), $formats[$styleHashIndex]);
				} else {
					switch ($cell->getDatatype()) {

					case PHPExcel_Cell_DataType::TYPE_STRING:
						if ($cell->getValue() === '' or $cell->getValue() === null) {
							$worksheet->writeBlank($row, $column, $formats[$styleHashIndex]);
						} else {
							$worksheet->writeString($row, $column, $cell->getValue(), $formats[$styleHashIndex]);
						}
						break;

					case PHPExcel_Cell_DataType::TYPE_FORMULA:
						$worksheet->writeFormula($row, $column, $cell->getValue(), $formats[$styleHashIndex]);
						break;

					case PHPExcel_Cell_DataType::TYPE_BOOL:
						$worksheet->writeBoolErr($row, $column, $cell->getValue(), 0, $formats[$styleHashIndex]);
						break;

					case PHPExcel_Cell_DataType::TYPE_ERROR:
						$worksheet->writeBoolErr($row, $column, $this->_mapErrorCode($cell->getValue()), 1, $formats[$styleHashIndex]);
						break;

					default:
						$worksheet->write($row, $column, $cell->getValue(), $formats[$styleHashIndex], $style->getNumberFormat()->getFormatCode());
						break;
					}

					// Hyperlink?
					if ($cell->hasHyperlink()) {
						$worksheet->writeUrl($row, $column, str_replace('sheet://', 'internal:', $cell->getHyperlink()->getUrl()));
					}
				}
			}

			// set the worksheet dimension for the DIMENSION record
			$worksheet->setDimensions($firstRowIndex, $lastRowIndex, $firstColumnIndex, $lastColumnIndex);

			$phpSheet->calculateColumnWidths();

			// Column dimensions
			foreach ($phpSheet->getColumnDimensions() as $columnDimension) {
				$column = PHPExcel_Cell::columnIndexFromString($columnDimension->getColumnIndex()) - 1;
				if ($column < 256) {
					if ($columnDimension->getWidth() >= 0) {
						$width = $columnDimension->getWidth();
					} else if ($phpSheet->getDefaultColumnDimension()->getWidth() >= 0) {
						$width = $phpSheet->getDefaultColumnDimension()->getWidth();
					} else {
						$width = 8;
					}
					$worksheet->setColumn( $column, $column, $width, null, ($columnDimension->getVisible() ? '0' : '1'), $columnDimension->getOutlineLevel());
				}
			}

			// Row dimensions
			foreach ($phpSheet->getRowDimensions() as $rowDimension) {
				$worksheet->setRow( $rowDimension->getRowIndex() - 1, $rowDimension->getRowHeight(), null, ($rowDimension->getVisible() ? '0' : '1'), $rowDimension->getOutlineLevel() );
			}

			foreach ($phpSheet->getDrawingCollection() as $drawing) {
				list($column, $row) = PHPExcel_Cell::coordinateFromString($drawing->getCoordinates());

				if ($drawing instanceof PHPExcel_Worksheet_Drawing) {
					$filename = $drawing->getPath();
					list($imagesx, $imagesy, $imageFormat) = getimagesize($filename);
					switch ($imageFormat) {
						case 1: $image = imagecreatefromgif($filename); break;
						case 2: $image = imagecreatefromjpeg($filename); break;
						case 3: $image = imagecreatefrompng($filename); break;
						default: continue 2;
					}
				} else if ($drawing instanceof PHPExcel_Worksheet_MemoryDrawing) {
					$image = $drawing->getImageResource();
					$imagesx = imagesx($image);
					$imagesy = imagesy($image);
				}

				$worksheet->insertBitmap($row - 1, PHPExcel_Cell::columnIndexFromString($column) - 1, $image, $drawing->getOffsetX(), $drawing->getOffsetY(), $drawing->getWidth() / $imagesx, $drawing->getHeight() / $imagesy);
			}
		}

		PHPExcel_Calculation_Functions::setReturnDateType($saveDateReturnType);

		$workbook->close();
	}

	/**
	 * Add color
	 */
	private function _addColor($workbook, $rgb) {
		if (!isset($this->_colors[$rgb])) {
			$workbook->setCustomColor(8 + count($this->_colors), hexdec(substr($rgb, 0, 2)), hexdec(substr($rgb, 2, 2)), hexdec(substr($rgb, 4)));
			$this->_colors[$rgb] = 8 + count($this->_colors);
		}
		return $this->_colors[$rgb];
	}

	/**
	 * Map border style
	 */
	private function _mapBorderStyle($borderStyle) {
		switch ($borderStyle) {
			case PHPExcel_Style_Border::BORDER_NONE:				return 0x00;
			case PHPExcel_Style_Border::BORDER_THIN;				return 0x01;
			case PHPExcel_Style_Border::BORDER_MEDIUM;				return 0x02;
			case PHPExcel_Style_Border::BORDER_DASHED;				return 0x03;
			case PHPExcel_Style_Border::BORDER_DOTTED;				return 0x04;
			case PHPExcel_Style_Border::BORDER_THICK;				return 0x05;
			case PHPExcel_Style_Border::BORDER_DOUBLE;				return 0x06;
			case PHPExcel_Style_Border::BORDER_HAIR;				return 0x07;
			case PHPExcel_Style_Border::BORDER_MEDIUMDASHED;		return 0x08;
			case PHPExcel_Style_Border::BORDER_DASHDOT;				return 0x09;
			case PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT;		return 0x0A;
			case PHPExcel_Style_Border::BORDER_DASHDOTDOT;			return 0x0B;
			case PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT;	return 0x0C;
			case PHPExcel_Style_Border::BORDER_SLANTDASHDOT;		return 0x0D;
			default:												return 0x00;
		}
	}

	/**
	 * Map underline
	 */
	private function _mapUnderline($underline) {
		switch ($underline) {
			case PHPExcel_Style_Font::UNDERLINE_NONE:
				return 0;
			case PHPExcel_Style_Font::UNDERLINE_SINGLE:
				return 1;
			case PHPExcel_Style_Font::UNDERLINE_DOUBLE:
				return 2;
			case PHPExcel_Style_Font::UNDERLINE_SINGLEACCOUNTING:
				return 0x21;
			case PHPExcel_Style_Font::UNDERLINE_DOUBLEACCOUNTING:
				return 0x22;
			default:
				return 0; // map others to none
		}
	}

	/**
	 * Map fill type
	 */
	private function _mapFillType($fillType) {
		switch ($fillType) {
			case PHPExcel_Style_Fill::FILL_NONE:					return 0x00;
			case PHPExcel_Style_Fill::FILL_SOLID:					return 0x01;
			case PHPExcel_Style_Fill::FILL_PATTERN_MEDIUMGRAY:		return 0x02;
			case PHPExcel_Style_Fill::FILL_PATTERN_DARKGRAY:		return 0x03;
			case PHPExcel_Style_Fill::FILL_PATTERN_LIGHTGRAY:		return 0x04;
			case PHPExcel_Style_Fill::FILL_PATTERN_DARKHORIZONTAL:	return 0x05;
			case PHPExcel_Style_Fill::FILL_PATTERN_DARKVERTICAL:	return 0x06;
			case PHPExcel_Style_Fill::FILL_PATTERN_DARKDOWN:		return 0x07;
			case PHPExcel_Style_Fill::FILL_PATTERN_DARKUP:			return 0x08;
			case PHPExcel_Style_Fill::FILL_PATTERN_DARKGRID:		return 0x09;
			case PHPExcel_Style_Fill::FILL_PATTERN_DARKTRELLIS:		return 0x0A;
			case PHPExcel_Style_Fill::FILL_PATTERN_LIGHTHORIZONTAL:	return 0x0B;
			case PHPExcel_Style_Fill::FILL_PATTERN_LIGHTVERTICAL:	return 0x0C;
			case PHPExcel_Style_Fill::FILL_PATTERN_LIGHTDOWN:		return 0x0D;
			case PHPExcel_Style_Fill::FILL_PATTERN_LIGHTUP:			return 0x0E;
			case PHPExcel_Style_Fill::FILL_PATTERN_LIGHTGRID:		return 0x0F;
			case PHPExcel_Style_Fill::FILL_PATTERN_LIGHTTRELLIS:	return 0x10;
			case PHPExcel_Style_Fill::FILL_PATTERN_GRAY125:			return 0x11;
			case PHPExcel_Style_Fill::FILL_PATTERN_GRAY0625:		return 0x12;
			case PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR:		// does not exist in BIFF8
			case PHPExcel_Style_Fill::FILL_GRADIENT_PATH:		// does not exist in BIFF8
			default:												return 0x00;
		}
	}

	/**
	 * Map VAlign
	 */
	private function _mapVAlign($vAlign) {
		return ($vAlign == 'center' || $vAlign == 'justify' ? 'v' : '') . $vAlign;
	}

	/**
	 * Map Error code
	 */
	private function _mapErrorCode($errorCode) {
		switch ($errorCode) {
			case '#NULL!':	return 0x00;
			case '#DIV/0!':	return 0x07;
			case '#VALUE!':	return 0x0F;
			case '#REF!':	return 0x17;
			case '#NAME?':	return 0x1D;
			case '#NUM!':	return 0x24;
			case '#N/A':	return 0x2A;
		}

		return 0;
	}

	/**
	 * Get an array of all styles
	 *
	 * @param	PHPExcel				$pPHPExcel
	 * @return	PHPExcel_Style[]		All styles in PHPExcel
	 * @throws	Exception
	 */
	private function _allStyles(PHPExcel $pPHPExcel = null)
	{
		// Get an array of all styles
		$aStyles		= array();

		for ($i = 0; $i < $pPHPExcel->getSheetCount(); ++$i) {
			foreach ($pPHPExcel->getSheet($i)->getStyles() as $style) {
				$aStyles[] = $style;
			}
		}

		return $aStyles;
	}

	/**
	 * Get temporary storage directory
	 *
	 * @return string
	 */
	public function getTempDir() {
		return $this->_tempDir;
	}

	/**
	 * Set temporary storage directory
	 *
	 * @param	string	$pValue		Temporary storage directory
	 * @throws	Exception	Exception when directory does not exist
	 */
	public function setTempDir($pValue = '') {
		if (is_dir($pValue)) {
			$this->_tempDir = $pValue;
		} else {
			throw new Exception("Directory does not exist: $pValue");
		}
	}
}
