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
 * PHPExcel_Shared_Escher_DggContainer
 *
 * @category   PHPExcel
 * @package    PHPExcel_Shared_Escher
 * @copyright  Copyright (c) 2006 - 2008 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Shared_Escher_DggContainer
{
	/**
	 * BLIP Store Container
	 *
	 * @var PHPExcel_Shared_Escher_DggContainer_BstoreContainer
	 */
	private $_bstoreContainer;

	/**
	 * Array of options for the drawing group
	 *
	 * @var array
	 */
	private $_OPT = array();

	/**
	 * Get BLIP Store Container
	 *
	 * @return PHPExcel_Shared_Escher_DggContainer_BstoreContainer
	 */
	public function getBstoreContainer()
	{
		return $this->_bstoreContainer;
	}

	/**
	 * Set BLIP Store Container
	 *
	 * @param PHPExcel_Shared_Escher_DggContainer_BstoreContainer $bstoreContainer
	 */
	public function setBstoreContainer($bstoreContainer)
	{
		$this->_bstoreContainer = $bstoreContainer;
	}

	/**
	 * Set an option for the drawing group
	 *
	 * @param int $property The number specifies the option
	 * @param mixed $value
	 */
	public function setOPT($property, $value)
	{
		$this->_OPT[$property] = $value;
	}

	/**
	 * Get an option for the drawing group
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

}
