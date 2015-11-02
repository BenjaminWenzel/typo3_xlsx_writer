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

/**
 * Class ExcelUtility
 * @package BW\XlsxWriter
 */
class ExcelUtility {

	/**
	 * @var int
	 */
	private static $excel2007maxRow = 1048576;

	/**
	 * @var int
	 */
	private static $excel2007maxCol = 16384;

	/**
	 * Ckecks a string for valid UTF8 encoding
	 *
	 * @param $string
	 *
	 * @return bool
	 */
	public static function isValidUTF8( $string ) {
		if( function_exists( "mb_check_encoding" ) ) {
			return mb_check_encoding( $string, "UTF-8" ) ? TRUE : FALSE;
		}

		return preg_match( "//u", $string ) ? TRUE : FALSE;
	}

	/**
	 * Replaces invalid xml chars
	 *
	 * @param string $val
	 *
	 * @return string
	 */
	public static function xmlspecialchars( $val ) {
		return str_replace( "'", "&#39;", htmlspecialchars( $val ) );
	}

	/**
	 * Removes invalid chars from filename
	 *
	 * @param $filename
	 *
	 * @return mixed
	 */
	public static function sanitizeFilename( $filename ) {
		/** @var array $nonPrinting */
		$nonPrinting = array_map( "chr", range( 0, 31 ) );
		/** @var array $invalidChars */
		$invalidChars = array( '<', '>', '?', '"', ':', '|', '\\', '/', '*', '&' );

		return str_replace( array_merge( $nonPrinting, $invalidChars ), "", $filename );
	}

	/**
	 * Returns cell label (ex: A1, C3)
	 *
	 * @param int $row
	 * @param int $column
	 *
	 * @return string
	 */
	public static function getXlSCell( $row, $column ) {
		for( $r = ""; $column >= 0; $column = intval( $column / 26 ) - 1 ) {
			$r = chr( $column % 26 + 0x41 ) . $r;
		}

		return $r . ( $row + 1 );
	}

	/**
	 * Returns the max cell label
	 *
	 * @return string
	 */
	public static function getMaxXLSCell() {
		return self::getXlSCell( self::$excel2007maxRow, self::$excel2007maxCol );
	}
}
