<?php

/***************************************************************
 *
 *  Copyright notice
 *
 *  (c) 2015 Benjamin Wenzel <benjamin.wenzel@mail.de>
 *
 *  All rights reserved
 *
 *  This script is part of the TYPO3 project. The TYPO3 project is
 *  free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation; either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  The GNU General Public License can be found at
 *  http://www.gnu.org/copyleft/gpl.html.
 *
 *  This script is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  This copyright notice MUST APPEAR in all copies of the script!
 ***************************************************************/

namespace BW\XlsxWriter;

use TYPO3\CMS\Core\Utility\GeneralUtility;
use \BW\XlsxWriter\ExcelUtility;

/**
 * Class ExcelSheet
 * @package BW\XlsxWriter
 */
class ExcelSheet {

	/**
	 * @var string
	 */
	protected $xmlName = "";

	/**
	 * @var string
	 */
	protected $sheetName = "";

	/**
	 * @var array
	 */
	protected $tempFiles = array();

	/**
	 * @var int
	 */
	protected $rowCount = 0;

	/**
	 * @var array
	 */
	protected $columns = array();

	/**
	 * @var string
	 */
	protected $sheetData = "";

	/**
	 * @var array
	 */
	protected $columnWidths = array();

	/**
	 * @var \BW\XlsxWriter\ExcelDocument
	 */
	protected $doc = NULL;

	/**
	 * ExcelSheet constructor.
	 *
	 * @param \BW\XlsxWriter\ExcelDocument $doc
	 * @param string                                   $sheetName
	 * @param int                                      $sheetCount
	 */
	public function __construct( $doc, $sheetName, $sheetCount ) {
		$this->doc = $doc;
		$this->sheetName = $sheetName;
		$this->xmlName = "sheet" . ( $sheetCount + 1 ) . ".xml";
	}

	/**
	 * Writes a data row
	 *
	 * @param array $row
	 */
	public function writeRow( $row ) {
		if( empty( $this->columns ) ) {
			$this->columns = array_fill( $from = 0, $until = count( $row ), "0" );
		}
		$this->sheetData .= "<row collapsed=\"false\" customFormat=\"false\" customHeight=\"false\" hidden=\"false\" ht=\"12.1\" outlineLevel=\"0\" r=\"' . ($this->rowCount + 1) . '\">";

		/** @var int $columnCount */
		$columnCount = 0;
		/** @var string $cellValue */
		foreach( $row as $cellValue ) {
			$this->writeCell( $cellValue, $columnCount );
			$columnCount++;
		}

		$this->sheetData .= "</row>";
		$this->rowCount++;
	}

	/**
	 * Writes an excel cell
	 *
	 * @param string $cellValue
	 * @param int    $columnCount
	 */
	protected function writeCell( $cellValue, $columnCount ) {
		/** @var string $cellName */
		$cellName = ExcelUtility::getXlSCell( $this->rowCount, $columnCount );
		/** @var bool $isFormula */
		$isFormula = FALSE;
		/** @var string $type */
		if( !is_scalar( $cellValue ) || $cellValue === "" ) {
			$cellValue = "";
		} elseif( is_string( $cellValue ) && $cellValue{0} == "=" ) {
			$type = "s";
			$cellValue = ExcelUtility::xmlspecialchars( $cellValue );
			$isFormula = TRUE;
		} elseif( !is_string( $cellValue ) ) {
			$type = "n";
			$cellValue = $cellValue * 1;
		} elseif( $cellValue{0} != "0" && $cellValue{0} != "+" && filter_var( $cellValue, FILTER_VALIDATE_INT, array( "options" => array( "max_range" => 2147483647 ) ) ) ) {
			$type = "n";
			$cellValue = $cellValue * 1;
		} else {
			$type = "s";
			$cellValue = ExcelUtility::xmlspecialchars( $this->doc->setSharedString( $cellValue ) );
		}

		$this->sheetData .= "<c r=\"" . $cellName . "\" s=\"" . ( $this->columns[ $columnCount ] ) . "\" ";
		if( isset( $type ) && !empty( $type ) ) {
			$this->sheetData .= "t=\"" . $type . "\"";
		}
		$this->sheetData .= ">";
		if( !$isFormula ) {
			$this->sheetData .= "<v>" . $cellValue . "</v></c>";
		} else {
			$this->sheetData .= "<f>" . $cellValue . "</f></c>";
		}

	}

	/**
	 * @return string
	 */
	public function getXmlName() {
		return $this->xmlName;
	}

	/**
	 * @return string
	 */
	public function getSheetName() {
		return $this->sheetName;
	}

	/**
	 * @param $columnWidths
	 */
	public function setColumnWidths( $columnWidths ) {
		$this->columnWidths = $columnWidths;
	}

	/**
	 * Returns content of sheet xml file
	 *
	 * @return string
	 */
	public function buildXML() {
		$xml = "";
		$xml .= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		$xml .= "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ";
		$xml .= "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ";
		$xml .= "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" ";
		$xml .= "xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">";

		/** @var string $maxCell */
		$maxCell = ExcelUtility::getXlSCell( $this->rowCount - 1, count( $this->columns ) - 1 );

		$xml .= "<dimension ref=\"" . $maxCell . "\" />";
		$xml .= "<sheetViews>";
		$xml .= "<sheetView tabSelected=\"1\" workbookViewId=\"0\">";
		$xml .= "<selection activeCell=\"A1\" sqref=\"A1\"/>";
		$xml .= "</sheetView>";
		$xml .= "</sheetViews>";
		$xml .= "<sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>";

		if( !empty( $this->columnWidths ) ) {
			$xml .= "<cols>";
			/**
			 * @var int   $i
			 * @var float $width
			 */
			foreach( $this->columnWidths as $i => $width ) {
				$xml .= "<col min=\"" . ( $i + 1 ) . "\" max=\"" . ( $i + 1 ) . "\" width=\"" . $width . "\" bestFit=\"1\" customWidth=\"1\"/>";
			}
			$xml .= "</cols>";
		}

		$xml .= "<sheetData>";
		$xml .= $this->sheetData;
		$xml .= "</sheetData>";
		$xml .= "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.78740157499999996\" bottom=\"0.78740157499999996\" header=\"0.3\" footer=\"0.3\"/>";
		$xml .= "</worksheet>";

		return $xml;
	}
}
