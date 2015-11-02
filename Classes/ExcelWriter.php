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

use BW\XlsxWriter\ExcelUtility;

/**
 * Class ExcelWriter
 * @package BW\XlsxWriter
 */
class ExcelWriter {

	/**
	 * @var resource
	 */
	protected $stream = NULL;

	/**
	 * @var string
	 */
	protected $buffer = "";

	/**
	 * @var bool
	 */
	protected $checkUTF8 = FALSE;

	/**
	 * ExcelWriter constructor.
	 *
	 * @param string     $filename
	 * @param string     $flags
	 * @param bool|FALSE $checkUTF8
	 *
	 * @throws \Exception
	 */
	public function __construct( $filename, $flags = "w", $checkUTF8 = FALSE ) {
		$this->checkUTF8 = $checkUTF8;
		$this->stream = fopen( $filename, $flags );
		if( $this->stream === FALSE ) {
			throw new \Exception( "Unable to open {$filename} for writing." );
		}
	}

	/**
	 * ExcelWriter deconstructor.
	 */
	public function __destruct() {
		$this->close();
	}

	/**
	 * Writes data to the stream
	 *
	 * @param $string
	 */
	public function write( $string ) {
		$this->buffer .= $string;
		if( isset( $this->buffer[ 8191 ] ) ) {
			$this->purge();
		}
	}

	/**
	 * Purges the buffer
	 */
	protected function purge() {
		if( $this->stream ) {
			if( $this->checkUTF8 && !ExcelUtility::isValidUTF8( $this->buffer ) ) {
				$this->checkUTF8 = FALSE;
			}
			fwrite( $this->stream, $this->buffer );
			$this->buffer = "";
		}
	}

	/**
	 * Closes the stream
	 */
	public function close() {
		$this->purge();
		if( $this->stream ) {
			fclose( $this->stream );
			$this->stream = NULL;
		}
	}

	/**
	 * Returns the current position of the file pointer
	 *
	 * @return int
	 */
	public function ftell() {
		if( $this->stream ) {
			$this->purge();

			return ftell( $this->stream );
		}

		return -1;
	}

	/**
	 * Moves the file pointer and returns the new position
	 *
	 * @param int $offset
	 *
	 * @return int
	 */
	public function fseek( $offset ) {
		if( $this->stream ) {
			$this->purge();

			return fseek( $this->stream, $offset );
		}

		return -1;
	}
}
