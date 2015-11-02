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

/**
 * Class ExcelDocument
 * @package BW\XlsxWriter
 */
class ExcelDocument {

	/**
	 * @var string
	 */
	protected $author = "Doc Author";

	/**
	 * @var \BW\XlsxWriter\ExcelSheet[]
	 */
	protected $sheets = array();

	/**
	 * @var array
	 */
	protected $tempFiles = array();

	/**
	 * Unique set
	 *
	 * @var array
	 */
	protected $sharedStrings = array();

	/**
	 * Count of non-unique references to the unique set
	 *
	 * @var int
	 */
	protected $sharedStringCount = 0;

	/**
	 * ExcelDocument constructor.
	 */
	public function __construct() { }

	/**
	 * ExcelDocument deconstructor.
	 */
	public function __destruct() {
		if( !empty( $this->tempFiles ) ) {
			/** @var string $tempFile */
			foreach( $this->tempFiles as $tempFile ) {
				@unlink( $tempFile );
			}
		}
	}

	/**
	 * Sends document to standard output
	 */
	public function writeToStdOut() {
		/** @var string $tempFile */
		$tempFile = $this->getTempFilename();
		$this->writeToFile( $tempFile );
		readfile( $tempFile );
	}

	/**
	 * Writes document to disc
	 *
	 * @param string $filename
	 *
	 * @throws \Exception
	 */
	public function writeToFile( $filename ) {
		if( file_exists( $filename ) ) {
			if( is_writable( $filename ) ) {
				@unlink( $filename );
			} else {
				throw new \Exception( "File is not " . $filename . " writeable." );
			}
		}

		/** @var \ZipArchive $zip */
		$zip = new \ZipArchive();

		if( !$zip->open( $filename, \ZipArchive::CREATE ) ) {
			throw new \Exception( "Unable to create zip." );
		}

		$zip->addEmptyDir( "docProps/" );
		$zip->addFromString( "docProps/app.xml", $this->buildAppXML() );
		$zip->addFromString( "docProps/core.xml", $this->buildCoreXML() );

		$zip->addEmptyDir( "_rels/" );
		$zip->addFromString( "_rels/.rels", $this->buildRelationshipsXML() );

		$zip->addEmptyDir( "xl/worksheets/" );
		/** @var \BW\XlsxWriter\ExcelSheet $sheet */
		foreach( $this->sheets as $sheet ) {
			$zip->addFromString( "xl/worksheets/" . $sheet->getXmlname(), $sheet->buildXML() );
		}

		if( !empty( $this->sharedStrings ) ) {
			$zip->addFromString( "xl/sharedStrings.xml", $this->buildSharedStringsXML() );
		}

		$zip->addFromString( "xl/workbook.xml", $this->buildWorkbookXML() );
		$zip->addFromString( "xl/styles.xml", $this->buildStylesXML() );
		$zip->addFromString( "[Content_Types].xml", $this->buildContentTypesXML() );

		$zip->addEmptyDir( "xl/_rels/" );
		$zip->addFromString( "xl/_rels/workbook.xml.rels", $this->buildWorkbookRelsXML() );
		$zip->close();
	}

	/**
	 * @param string $author
	 */
	public function setAuthor( $author ) {
		$this->author = $author;
	}

	/**
	 * @param array  $data
	 * @param string $sheetName
	 * @param array  $headerTypes
	 */
	public function writeSheet( $data, $sheetName, $headerTypes = array() ) {
		$sheetName = empty( $sheetName ) ? "Sheet1" : $sheetName;
		/** @var \BW\XlsxWriter\ExcelSheet $sheet */
		$sheet = $this->initializeSheet( $sheetName );

		$data = empty( $data ) ? array( array( "" ) ) : $data;
		if( !empty( $headerTypes ) ) {
			//$this->writeSheetHeader($sheet_name, $header_types);
		}
		foreach( $data as $i => $row ) {
			$sheet->writeRow( $row );
		}
	}

	/**
	 * Initializes a new excel sheet
	 *
	 * @param $sheetName
	 *
	 * @return \BW\XlsxWriter\ExcelSheet
	 */
	protected function initializeSheet( $sheetName ) {
		/** @var \BW\XlsxWriter\ExcelSheet $sheet */
		$sheet = GeneralUtility::makeInstance( "\\BW\\XlsxWriter\\ExcelSheet", $this, $sheetName, count( $this->sheets ) );

		$this->sheets[ $sheetName ] = $sheet;

		return $sheet;
	}

	/**
	 * TODO: Find proper description for function
	 *
	 * @param string $value
	 *
	 * @return string
	 */
	public function setSharedString( $value ) {
		/** @var string $stringValue */
		if( isset( $this->sharedStrings[ $value ] ) ) {
			$stringValue = $this->sharedStrings[ $value ];
		} else {
			$stringValue = count( $this->sharedStrings );
			$this->sharedStrings[ $value ] = $stringValue;
		}
		$this->sharedStringCount++;

		return $stringValue;
	}

	/**
	 * Returns content of "/xl/styles.xml"
	 *
	 * @return string
	 */
	protected function buildStylesXML() {
		/** @var string $xml */
		$xml = "";
		$xml .= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		$xml .= "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ";
		$xml .= "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" ";
		$xml .= "xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">";
		$xml .= "<fonts count=\"1\" x14ac:knownFonts=\"1\">";
		$xml .= "<font><name val=\"Arial\"/><charset val=\"1\"/><family val=\"2\"/><sz val=\"10\"/></font>";
		$xml .= "</fonts>";
		$xml .= "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>";
		$xml .= "<borders count=\"1\"><border diagonalDown=\"false\" diagonalUp=\"false\"><left/><right/><top/><bottom/><diagonal/></border></borders>";
		$xml .= "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>";
		$xml .= "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>";
		$xml .= "<cellStyles count=\"1\"><cellStyle name=\"Standard\" xfId=\"0\" builtinId=\"0\"/></cellStyles>";
		$xml .= "<dxfs count=\"0\"/>";
		$xml .= "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>";
		$xml .= "</styleSheet>";

		return $xml;
	}

	/**
	 * Returns content of "/xl/sharedStrings.xml"
	 *
	 * @return string
	 */
	protected function buildSharedStringsXML() {
		/** @var string $xml */
		$xml = "";
		$xml .= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		$xml .= "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"" . $this->sharedStringCount . "\" uniqueCount=\"" . count( $this->sharedStrings ) . "\">";

		/**
		 * @var string $key
		 * @var string $value
		 */
		foreach( $this->sharedStrings as $key => $value ) {
			$xml .= "<si><t>" . ExcelUtility::xmlspecialchars( $key ) . "</t></si>";
		}

		$xml .= "</sst>";

		return $xml;
	}

	/**
	 * Returns content of "/docProps/app.xml"
	 *
	 * @return string
	 */
	protected function buildAppXML() {
		/** @var string $xml */
		$xml = "";
		$xml .= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		$xml .= "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" ";
		$xml .= "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">";
		$xml .= "<TotalTime>0</TotalTime>";
		$xml .= "</Properties>";

		return $xml;
	}

	/**
	 * Returns content of "/docProps/core.xml"
	 *
	 * @return string
	 */
	protected function buildCoreXML() {
		/** @var string $xml */
		$xml = "";
		$xml .= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		$xml .= "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" ";
		$xml .= "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" ";
		$xml .= "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">";
		$xml .= "<dc:creator>" . ExcelUtility::xmlspecialchars( $this->author ) . "</dc:creator>";
		$xml .= "<cp:lastModifiedBy>" . ExcelUtility::xmlspecialchars( $this->author ) . "</cp:lastModifiedBy>";
		/** @var string $now */
		$now = date( "Y-m-d\TH:i:s.00\Z" );
		$xml .= "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" . $now . "</dcterms:created>";
		$xml .= "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" . $now . "</dcterms:modified>";
		$xml .= "</cp:coreProperties>";

		return $xml;
	}

	/**
	 * Returns content of "/_rels/.rels"
	 *
	 * @return string
	 */
	protected function buildRelationshipsXML() {
		/** @var string $xml */
		$xml = "";
		$xml .= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		$xml .= "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
		$xml .= "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>";
		$xml .= "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>";
		$xml .= "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>";
		$xml .= "</Relationships>";

		return $xml;
	}

	/**
	 * Returns content of "/xl/workbook.xml"
	 *
	 * @return string
	 */
	protected function buildWorkbookXML() {
		/** @var string $xml */
		$xml = "";
		$xml .= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		$xml .= "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ";
		$xml .= "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ";
		$xml .= "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x15\" ";
		$xml .= "xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\">";
		$xml .= "<fileVersion appName=\"xl\" lastEdited=\"6\" lowestEdited=\"6\" rupBuild=\"14420\"/>";
		$xml .= "<workbookPr defaultThemeVersion=\"153222\"/>";
		$xml .= "<bookViews>";
		$xml .= "<workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"24270\" windowHeight=\"12570\"/>";
		$xml .= "</bookViews>";
		$xml .= "<sheets>";

		/** @var int $i */
		$i = 0;
		/** @var  $sheet \BW\XlsxWriter\ExcelSheet */
		foreach( $this->sheets as $sheet ) {
			$xml .= "<sheet name=\"" . ExcelUtility::xmlspecialchars( $sheet->getSheetName() ) . "\" sheetId=\"" . ( $i + 1 ) . "\" r:id=\"rId" . ( $i + 2 ) . "\"/>";
			$i++;
		}

		$xml .= "</sheets>";
		$xml .= "<calcPr calcId=\"152511\"/>";
		$xml .= "</workbook>";

		return $xml;
	}

	/**
	 * Returns content of "/xl/_rels/workbook.xml.rels"
	 *
	 * @return string
	 */
	protected function buildWorkbookRelsXML() {
		/** @var string $xml */
		$xml = "";
		$xml .= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		$xml .= "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
		$xml .= "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>";

		/** @var int $i */
		$i = 2;
		/** @var  $sheet \BW\XlsxWriter\ExcelSheet */
		foreach( $this->sheets as $sheet ) {
			$xml .= "<Relationship Id=\"rId{$i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/{$sheet->getXmlName()}\"/>";
			$i++;
		}

		if( !empty( $this->sharedStrings ) ) {
			$xml .= "<Relationship Id=\"rId{$i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>";
		}

		$xml .= "</Relationships>";

		return $xml;
	}

	/**
	 * Returns content of "/[Content_Types].xml"
	 *
	 * @return string
	 */
	protected function buildContentTypesXML() {
		/** @var string $contentTypeXML */
		$contentTypeXML = "";
		$contentTypeXML .= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		$contentTypeXML .= "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">";
		$contentTypeXML .= "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>";
		$contentTypeXML .= "<Default Extension=\"xml\" ContentType=\"application/xml\"/>";
		$contentTypeXML .= "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>";

		/** @var  $sheet \BW\XlsxWriter\ExcelSheet */
		foreach( $this->sheets as $sheet ) {
			$contentTypeXML .= "<Override PartName=\"/xl/worksheets/{$sheet->getXmlName()}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
		}

		if( !empty( $this->sharedStrings ) ) {
			$contentTypeXML .= "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
		}

		$contentTypeXML .= "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>";
		$contentTypeXML .= "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
		$contentTypeXML .= "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>";
		$contentTypeXML .= "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>";
		$contentTypeXML .= "</Types>";

		return $contentTypeXML;
	}

	/**
	 * Returns a temporary filename
	 *
	 * @return string
	 */
	protected function getTempFilename() {
		$filename = tempnam( sys_get_temp_dir(), "xlsx_writer_" );
		$this->tempFiles[] = $filename;

		return $filename;
	}

}
