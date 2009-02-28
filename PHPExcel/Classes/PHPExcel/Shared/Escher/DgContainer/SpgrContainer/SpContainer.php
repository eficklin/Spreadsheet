<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2008 PHPExcel
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
 * You should have received a copy of tshhe GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel_Shared_Escher
 * @copyright  Copyright (c) 2006 - 2008 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/**
 * PHPExcel_Shared_Escher_DgContainer_SpgrContainer_SpContainer
 *
 * @category   PHPExcel
 * @package    PHPExcel_Shared_Escher
 * @copyright  Copyright (c) 2006 - 2008 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Shared_Escher_DgContainer_SpgrContainer_SpContainer
{
	/**
	 * Parent Shape Group Container
	 *
	 * @var PHPExcel_Shared_Escher_DgContainer_SpgrContainer
	 */
	private $_parent;

	/**
	 * Array of options
	 *
	 * @var array
	 */
	private $_OPT;

	/**
	 * Cell coordinates of upper-left corner of shape, e.g. 'A1'
	 *
	 * @var string
	 */
	private $_startCoordinates;

	/**
	 * Horizontal offset of upper-left corner of shape measured in 1/1024 of column width
	 *
	 * @var int
	 */
	private $_startOffsetX;

	/**
	 * Vertical offset of upper-left corner of shape measured in 1/256 of row height
	 *
	 * @var int
	 */
	private $_startOffsetY;

	/**
	 * Cell coordinates of bottom-right corner of shape, e.g. 'B2'
	 *
	 * @var string
	 */
	private $_endCoordinates;

	/**
	 * Horizontal offset of bottom-right corner of shape measured in 1/1024 of column width
	 *
	 * @var int
	 */
	private $_endOffsetX;

	/**
	 * Vertical offset of bottom-right corner of shape measured in 1/256 of row height
	 *
	 * @var int
	 */
	private $_endOffsetY;

	/**
	 * Set parent Shape Group Container
	 *
	 * @param PHPExcel_Shared_Escher_DgContainer_SpgrContainer $parent
	 */
	public function setParent($parent)
	{
		$this->_parent = $parent;
		$this->_OPT = array();
	}

	/**
	 * Set an option for the Shape Group Container
	 *
	 * @param int $property The number specifies the option
	 * @param mixed $value
	 */
	public function setOPT($property, $value)
	{
		$this->_OPT[$property] = $value;
	}

	/**
	 * Get an option for the Shape Group Container
	 *
	 * @param int $property The number specifies the option
	 * @return mixed
	 */
	public function getOPT($property)
	{
		if (isset($this->_OPT[$property])) {
			return $this->_OPT[$property];
		}
		return null;
	}

	/**
	 * Set cell coordinates of upper-left corner of shape
	 *
	 * @param string $value
	 */
	public function setStartCoordinates($value = 'A1')
	{
		$this->_startCoordinates = $value;
	}

	/**
	 * Get cell coordinates of upper-left corner of shape
	 *
	 * @return string
	 */
	public function getStartCoordinates()
	{
		return $this->_startCoordinates;
	}

	/**
	 * Set offset in x-direction of upper-left corner of shape measured in 1/1024 of column width
	 *
	 * @param int $startOffsetX
	 */
	public function setStartOffsetX($startOffsetX = 0)
	{
		$this->_startOffsetX = $startOffsetX;
	}

	/**
	 * Get offset in x-direction of upper-left corner of shape measured in 1/1024 of column width
	 *
	 * @return int
	 */
	public function getStartOffsetX()
	{
		return $this->_startOffsetX;
	}

	/**
	 * Set offset in y-direction of upper-left corner of shape measured in 1/256 of row height
	 *
	 * @param int $startOffsetY
	 */
	public function setStartOffsetY($startOffsetY = 0)
	{
		$this->_startOffsetY = $startOffsetY;
	}

	/**
	 * Get offset in y-direction of upper-left corner of shape measured in 1/256 of row height
	 *
	 * @return int
	 */
	public function getStartOffsetY()
	{
		return $this->_startOffsetY;
	}

	/**
	 * Set cell coordinates of bottom-right corner of shape
	 *
	 * @param string $value
	 */
	public function setEndCoordinates($value = 'A1')
	{
		$this->_endCoordinates = $value;
	}

	/**
	 * Get cell coordinates of bottom-right corner of shape
	 *
	 * @return string
	 */
	public function getEndCoordinates()
	{
		return $this->_endCoordinates;
	}

	/**
	 * Set offset in x-direction of bottom-right corner of shape measured in 1/1024 of column width
	 *
	 * @param int $startOffsetX
	 */
	public function setEndOffsetX($endOffsetX = 0)
	{
		$this->_endOffsetX = $endOffsetX;
	}

	/**
	 * Get offset in x-direction of bottom-right corner of shape measured in 1/1024 of column width
	 *
	 * @return int
	 */
	public function getEndOffsetX()
	{
		return $this->_endOffsetX;
	}

	/**
	 * Set offset in y-direction of bottom-right corner of shape measured in 1/256 of row height
	 *
	 * @param int $endOffsetY
	 */
	public function setEndOffsetY($endOffsetY = 0)
	{
		$this->_endOffsetY = $endOffsetY;
	}

	/**
	 * Get offset in y-direction of bottom-right corner of shape measured in 1/256 of row height
	 *
	 * @return int
	 */
	public function getEndOffsetY()
	{
		return $this->_endOffsetY;
	}

}
