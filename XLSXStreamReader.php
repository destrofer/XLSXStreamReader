<?php
/*
 * Copyright 2016 Viacheslav Soroka
 * Author: Viacheslav Soroka
 * Version: 1.0.0
 * 
 * You can get latest version of this file at: https://github.com/destrofer/XLSXStreamReader
 * 
 * XLSXStreamReader is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * XLSXStreamReader is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public License
 * along with XLSXStreamReader.  If not, see <http://www.gnu.org/licenses/>.
 *
 * -----------------------------------------------------------------------------
 *
 * Reader depends on BigZip class (https://github.com/destrofer/BigZip)
 */

class XLSXStreamReader {
	const BLOCK_SIZE = 65536;

	const ERROR_INVALID_FORMAT_NOT_ZIP = 1;
	const ERROR_INVALID_FORMAT_NO_GLOBAL_RELS = 2;
	const ERROR_INVALID_FORMAT_NO_WORKBOOK_RELS = 3;
	const ERROR_INVALID_FORMAT_NO_WORKBOOK = 4;
	const ERROR_INVALID_FORMAT_BAD_WORKSHEET_REL = 5;
	const ERROR_CANNOT_CREATE_XML_PARSER = 6;
	const ERROR_CANNOT_READ_PACKED_FILE = 7;
	const ERROR_INVALID_XML = 8;
	const ERROR_WORKSHEET_NOT_OPEN = 9;
	const ERROR_FILE_CLOSED = 10;

	const CELL_FORMAT_GENERAL = 0;
	const CELL_FORMAT_PERCENT = 1;
	const CELL_FORMAT_DATE = 2;
	const CELL_FORMAT_DATE_TIME = 3;

	protected $file;
	protected $worksheets;
	protected $sharedStrings;
	protected $cellFormats;
	protected $calendarBasedOn1904 = false;

	protected $worksheetEntryFile = null;
	protected $worksheetXMLParser = null;
	protected $rowsQueue = [];
	protected $lastParsedRowIndex = 0;
	protected $parseRow = null;
	protected $parseCell = null;
	protected $parsingEnded = false;

	/**
	 * @var bool If TRUE readRow() method will return DateTime objects instead os strings for date and time cells.
	 */
	public $returnDatesAsDateTimeObjects = false;

	/**
	 * @var string Name of the class to use when creating DateTime objects.
	 */
	public $dateTimeClass = 'DateTime';

	/**
	 * XLSXStreamReader constructor.
	 * @param $filePath
	 * @throws Exception In case there was an error loading the file. Reason of exception can be determined by checking thrown exception code (ERROR_INVALID_FORMAT_NOT_ZIP, ERROR_INVALID_FORMAT_NO_GLOBAL_RELS, ERROR_INVALID_FORMAT_NO_WORKBOOK_RELS, ERROR_INVALID_FORMAT_NO_WORKBOOK, ERROR_INVALID_FORMAT_BAD_WORKSHEET_REL, ERROR_CANNOT_CREATE_XML_PARSER, ERROR_CANNOT_READ_PACKED_FILE or ERROR_INVALID_XML).
	 */
	public function __construct($filePath) {
		$this->file = BigZip::openForRead($filePath);
		if( !$this->file )
			throw new Exception('The file is damaged or its format is not XLSX', self::ERROR_INVALID_FORMAT_NOT_ZIP);

		$this->worksheets = [];
		$this->sharedStrings = [];
		$this->dateFormatStyles = [];
		$this->cellFormats = [];

		$workbookFile = 'xl/workbook.xml';
		$stylesFile = 'xl/styles.xml';
		$sharedStringsFile = 'xl/sharedStrings.xml';
		$worksheetFiles = [];

		// Read global file relationships to find workbook file path just in case it differs from default.
		$file = $this->file->entryOpen("_rels/.rels");
		if( !$file )
			throw new Exception('The file is not XLSX', self::ERROR_INVALID_FORMAT_NO_GLOBAL_RELS);
		$parser = $this->createXmlParser(function($parser, $name, $attr) use(&$workbookFile) {
			if( $name == 'RELATIONSHIP' && $attr['TYPE'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' )
				$workbookFile = $attr['TARGET'];
		});
		while( !$this->parseXml($parser, $file) );
		xml_parser_free($parser);
		$this->file->entryClose();

		$workbookPath = dirname($workbookFile);
		$workbookFileName = basename($workbookFile);

		// Read workbook relationships to find worksheets and shared strings files.
		$file = $this->file->entryOpen("{$workbookPath}/_rels/{$workbookFileName}.rels");
		if( !$file )
			throw new Exception('The file is damaged or its format is not XLSX', self::ERROR_INVALID_FORMAT_NO_WORKBOOK_RELS);
		$parser = $this->createXmlParser(function($parser, $name, $attr) use($workbookPath, &$stylesFile, &$sharedStringsFile, &$worksheetFiles) {
			if( $name == 'RELATIONSHIP' ) {
				if( $attr['TYPE'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' )
					$stylesFile = $workbookPath . '/' . $attr['TARGET'];
				else if( $attr['TYPE'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings' )
					$sharedStringsFile = $workbookPath . '/' . $attr['TARGET'];
				else if( $attr['TYPE'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' )
					$worksheetFiles[$attr['ID']] = $workbookPath . '/' . $attr['TARGET'];
			}
		});
		while( !$this->parseXml($parser, $file) );
		xml_parser_free($parser);
		$this->file->entryClose();

		// Read workbook to get info on worksheets and calendar.
		$file = $this->file->entryOpen($workbookFile);
		if( !$file )
			throw new Exception('The file is damaged or its format is not XLSX', self::ERROR_INVALID_FORMAT_NO_WORKBOOK);
		$parser = $this->createXmlParser(function($parser, $name, $attr) use(&$worksheetFiles) {
			if( $name == 'WORKBOOKPR' && isset($attr['DATE1904']) && ($attr['DATE1904'] == '1' || $attr['DATE1904'] == 'true') )
				$this->calendarBasedOn1904 = true;
			if( $name == 'SHEET' ) {
				if( !isset($worksheetFiles[$attr['R:ID']]) )
					throw new Exception('The file is damaged or its format is not XLSX', self::ERROR_INVALID_FORMAT_BAD_WORKSHEET_REL);
				$this->worksheets[] = [
					'id' => $attr['SHEETID'],
					'file' => $worksheetFiles[$attr['R:ID']],
					'name' => $attr['NAME'],
				];
			}
		});
		while( !$this->parseXml($parser, $file) );
		xml_parser_free($parser);
		$this->file->entryClose();

		// Read styles file to get cell formats.
		$file = $this->file->entryOpen($stylesFile);
		if( $file ) {
			$numFmts = false;
			$cellXfs = false;
			$numberFormats = [];
			$xfs = [];
			$parser = $this->createXmlParser(function($parser, $name, $attr) use(&$numFmts, &$cellXfs, &$numberFormats, &$xfs) {
				if( !$numFmts && $name == 'NUMFMTS' )
					$numFmts = true;
				else if( !$cellXfs && $name == 'CELLXFS' )
					$cellXfs = true;
				else if( $numFmts && $name == 'NUMFMT' && isset($attr['FORMATCODE']) && $attr['FORMATCODE'] != 'General' )
					$numberFormats[$attr['NUMFMTID']] = $attr['FORMATCODE'];
				else if( $cellXfs && $name == 'XF' && isset($attr['NUMFMTID']) )
					$xfs[] = $attr['NUMFMTID'];
			}, null, function($parser, $name) use(&$numFmts, &$cellXfs) {
				if( $numFmts && $name == 'NUMFMTS' )
					$numFmts = false;
				else if( $cellXfs && $name == 'CELLXFS' )
					$cellXfs = false;
			});
			while( !$this->parseXml($parser, $file) );
			xml_parser_free($parser);
			$this->file->entryClose();
			foreach( $xfs as $idx => $fmtId ) {
				if( isset($numberFormats[$fmtId]) ) {
					$fmt = preg_replace('#\[.*\]#sUu', '', $numberFormats[$fmtId]);
					if( preg_match('#(H|HH|S|SS)#isu', $fmt) )
						$this->cellFormats[$idx] = self::CELL_FORMAT_DATE_TIME;
					else if( preg_match('#(YY|YYYY|M{1,4}|D{1,5})#isu', $fmt) )
						$this->cellFormats[$idx] = self::CELL_FORMAT_DATE;
					else if( preg_match('#%#su', $fmt) )
						$this->cellFormats[$idx] = self::CELL_FORMAT_PERCENT;
					else
						$this->cellFormats[$idx] = self::CELL_FORMAT_GENERAL;
				}
				else
					$this->cellFormats[$idx] = self::CELL_FORMAT_GENERAL;
			}
		}

		// Read shared strings if such file exists. This will consume most of the memory.
		$file = $this->file->entryOpen($sharedStringsFile);
		if( $file ) {
			$currentString = null;
			$parser = $this->createXmlParser(function($parser, $name, $attr) use(&$currentString) {
				if( $currentString ) {
					if( $name == 'T' ) // we need only text, so ignore everything other than T nodes
						$currentString['collect'] = true;
				}
				else if( $name == 'SI' )
					$currentString = ['collect' => false, 'text' => ''];
			}, function($parser, $data) use(&$currentString) {
				if( $currentString && $currentString['collect'] )
					$currentString['text'] .= $data;
			}, function($parser, $name) use(&$currentString) {
				if( $currentString ) {
					if( $name == 'SI' ) {
						$this->sharedStrings[] = $currentString['text'];
						$currentString = null;
					}
					else if( $name == 'T' )
						$currentString['collect'] = false;
				}
			});
			while( !$this->parseXml($parser, $file) );
			xml_parser_free($parser);
			$this->file->entryClose();
		}
	}

	public function __destruct() {
		$this->close();
	}

	/**
	 * Closes the XLSX file.
	 */
	public function close() {
		if( $this->worksheetEntryFile ) {
			xml_parser_free($this->worksheetXMLParser);
			$this->file->entryClose();
		}
		$this->file->close();
		$this->file = null;
	}

	/**
	 * @param callable $nodeStart function($parser, $name, $attributes)
	 * @param callable|null $nodeData function($parser, $data)
	 * @param callable|null $nodeEnd function($parser, $name)
	 * @return resource
	 * @throws Exception
	 */
	private function createXmlParser($nodeStart, $nodeData = null, $nodeEnd = null) {
		$parser = xml_parser_create('UTF-8');
		if( !$parser )
			throw new Exception('Cannot create XML parser', self::ERROR_CANNOT_CREATE_XML_PARSER);
		xml_set_object($parser, $this);
		xml_parser_set_option($parser, XML_OPTION_TARGET_ENCODING, 'utf-8');
		xml_parser_set_option($parser, XML_OPTION_CASE_FOLDING, 1);
		xml_set_element_handler($parser, $nodeStart, $nodeEnd);
		xml_set_character_data_handler($parser, $nodeData);
		return $parser;
	}

	/**
	 * @param resource $parser
	 * @param resource $file
	 * @return bool
	 * @throws Exception
	 */
	private function parseXml($parser, $file) {
		if( feof($file) )
			return true;
		$data = fread($file, self::BLOCK_SIZE);
		if( $data === false )
			throw new Exception('There was an error while trying to read the data from packed file', self::ERROR_CANNOT_READ_PACKED_FILE);
		$isEOF = feof($file);
		if( !xml_parse($parser, $data, $isEOF) )
			throw new Exception('XML parsing error: [' . xml_get_error_code($parser) . '] ' . xml_error_string(xml_get_error_code($parser)) . ' on line ' . xml_get_current_line_number($parser), self::ERROR_INVALID_XML);
		return $isEOF;
	}

	/**
	 * Find worksheet index by its name. Case insensitive. Spaces at the beginning and end are ignored.
	 *
	 * @param string $name
	 * @return int|null Will return index of worksheet or NULL if worksheet was not found.
	 */
	public function findWorksheetByName($name) {
		if( !$this->file )
			return null;
		$name = trim(mb_strtolower($name, 'utf-8'));
		foreach( $this->worksheets as $idx => $worksheet )
			if( trim(mb_strtolower($worksheet['name'], 'utf-8')) == $name )
				return $idx;
		return null;
	}

	/**
	 * Opens worksheet for reading.
	 *
	 * Note: Only one worksheet may be opened at the same time. Opening another worksheet will automatically close
	 * currently open.
	 *
	 * @param int $index Index of the worksheet. It can be found out by name using findWorksheetByName() method.
	 * @return bool TRUE if worksheet opened successfully or FALSE otherwise.
	 * @throws Exception In case when XML parser could not be created (code=ERROR_CANNOT_CREATE_XML_PARSER) or file is already closed(code=ERROR_FILE_CLOSED).
	 */
	public function openWorksheet($index) {
		if( !$this->file )
			throw new Exception("File is closed", self::ERROR_FILE_CLOSED);
		if( $index === null || $index < 0 || $index >= count($this->worksheets) )
			return false;
		if( $this->worksheetEntryFile ) {
			xml_parser_free($this->worksheetXMLParser);
			$this->file->entryClose();
		}
		$this->rowsQueue = [];
		$this->worksheetEntryFile = $this->file->entryOpen($this->worksheets[$index]['file']);
		$this->parseRow = null;
		$this->parseCell = null;
		$this->parsingEnded = false;
		$this->lastParsedRowIndex = 0;

		if( !$this->worksheetEntryFile )
			return false;

		$this->worksheetXMLParser = $this->createXmlParser(function($parser, $name, $attr) {
			if( $this->parsingEnded )
				return;
			if( $name == 'ROW' ) {
				$index = intval($attr['R']);
				// fill empty rows just in case there are no nodes for rows in between
				for( $i = $this->lastParsedRowIndex + 1; $i < $index; $i++ )
					$this->rowsQueue[$i] = [];
				$this->lastParsedRowIndex = $index;
				$this->parseRow = [
					'index' => $index,
					'cells' => [],
				];
			}
			else if( $this->parseRow && $name == 'C' ) {
				$this->parseCell = [
					'column' => preg_replace('#[0-9]+$#su', '', $attr['R']),
					'type' => isset($attr['T']) ? $attr['T'] : null,
					'format' => isset($attr['S']) ? $attr['S'] : null,
					'collect' => false,
					'value' => '',
				];
			}
			else if( $this->parseCell && $name == 'V' ) {
				$this->parseCell['collect'] = true;
			}

		}, function($parser, $data) {
			if( $this->parseCell && $this->parseCell['collect'] )
				$this->parseCell['value'] .= $data;
		}, function($parser, $name) {
			if( $this->parseCell ) {
				if( $name == 'V' )
					$this->parseCell['collect'] = false;
				else if( $name == 'C' ) {
					$value = $this->parseCell['value'];
					if( $this->parseCell['type'] === null && is_numeric($value) )
						$this->parseCell['type'] = 'n';
					if( $this->parseCell['type'] == 's' )
						$value = isset($this->sharedStrings[$value]) ? $this->sharedStrings[$value] : $value;
					else if( $this->parseCell['type'] == 'b' )
						$value = !!$value;
					else if( $this->parseCell['type'] == 'n' ) {
						$format = $this->parseCell['format'];
						$format = isset($this->cellFormats[$format]) ? $this->cellFormats[$format] : self::CELL_FORMAT_GENERAL;
						switch( $format ) {
							case self::CELL_FORMAT_PERCENT: {
								$value = floatval($value) * 100;
								break;
							}
							case self::CELL_FORMAT_DATE:
							case self::CELL_FORMAT_DATE_TIME: {
								$timeFormat = 'Y-m-d' . (($format == self::CELL_FORMAT_DATE_TIME) ? ' H:i:s' : '');
								$value = floatval($value);
								if( $this->calendarBasedOn1904 )
									$base = 24107; // for 1904 based calendar
								else
									$base = ($value >= 60) ? 25569 : 25568; // for 1900 based calendar
								if( $value <= 0 ) {
									$hours = round($value * 24);
									$minutes = round($value * 1440) - round($hours * 60);
									$seconds = round($value * 86400) - round($hours * 3600) - round($minutes * 60);
									$value = gmdate($timeFormat, mktime($hours, $minutes, $seconds));
								}
								else
									$value = gmdate($timeFormat, round(($value - $base) * 86400));
								if( $this->returnDatesAsDateTimeObjects )
									$value = new $this->dateTimeClass($value);
								break;
							}
							default: {
								$value = floatval($value);
							}
						}
					}

					if( $value !== '' )
						$this->parseRow['cells'][$this->parseCell['column']] = $value;
					$this->parseCell = null;
				}
			}
			else if( $this->parseRow ) {
				if( $name == 'ROW' ) {
					$this->rowsQueue[$this->parseRow['index']] = $this->parseRow['cells'];
					$this->parseRow = null;
				}
			}
			else if( $name == 'SHEETDATA' )
				$this->parsingEnded = true;
		});

		return true;
	}

	/**
	 * Reads next row from the open worksheet.
	 *
	 * Note: Be careful when checking returned value. The method may return an empty array, which also evaluates to
	 * FALSE. Loop "while($row = $reader->readRow()) ..." has a chance of stopping on the first empty row. It is better
	 * to use "while(($row = $reader->readRow()) !== null) ..." instead.
	 *
	 * @return array|null Returns either an array containing all cell data of the row or NULL if no more rows are available in the worksheet.
	 * @throws Exception Thrown if file is closed (code=ERROR_FILE_CLOSED), no worksheet currently open (code=ERROR_WORKSHEET_NOT_OPEN), there was an error while reading (code=ERROR_CANNOT_READ_PACKED_FILE) or parsing (code=ERROR_CANNOT_READ_PACKED_FILE) packed XML.
	 */
	public function readRow() {
		if( !$this->file )
			throw new Exception("File is closed", self::ERROR_FILE_CLOSED);
		if( !$this->worksheetEntryFile )
			throw new Exception('Use openWorksheet() method before reading rows', self::ERROR_WORKSHEET_NOT_OPEN);

		while( !$this->parsingEnded && empty($this->rowsQueue) ) {
			if( $this->parseXml($this->worksheetXMLParser, $this->worksheetEntryFile) )
				$this->parsingEnded = true;
		}

		if( empty($this->rowsQueue) )
			return null;

		$row = array_shift($this->rowsQueue);
		// possibly need cell reindexing to numbers instead of letters
		return $row;
	}
}