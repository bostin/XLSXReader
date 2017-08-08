<?php
/**
 * Created by PhpStorm.
 *  ELXS 文件读取类
 *
 * User: Bostin
 * Date: 2017/8/8
 * Time: 13:52
 */

namespace Bostin\Office\Excel;

use Exception;
use ZipArchive;


/**
 * Class XLSXReader
 */
class XLSXReader
{
    protected $sheets        = [];
    protected $sharedstrings = [];
    protected $sheetInfo;
    protected $zip;
    public    $config        = [
        'removeTrailingRows' => true,
    ];

    // XML schemas
    const SCHEMA_OFFICEDOCUMENT              = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
    const SCHEMA_RELATIONSHIP                = 'http://schemas.openxmlformats.org/package/2006/relationships';
    const SCHEMA_OFFICEDOCUMENT_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const SCHEMA_SHAREDSTRINGS               = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
    const SCHEMA_WORKSHEETRELATION           = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

    public function __construct($filePath, $config = [])
    {
        $this->config = array_merge($this->config, $config);
        $this->zip    = new ZipArchive();
        $status       = $this->zip->open($filePath);
        if ($status === true) {
            $this->parse();
        } else {
            throw new Exception("Failed to open $filePath with zip error code: $status");
        }
    }

    // get a file from the zip
    protected function getEntryData($name)
    {
        $data = $this->zip->getFromName($name);
        if ($data === false) {
            throw new Exception("File $name does not exist in the Excel file");
        } else {
            return $data;
        }
    }

    // extract the shared string and the list of sheets
    protected function parse()
    {
        $sheets           = [];
        $relationshipsXML = simplexml_load_string($this->getEntryData("_rels/.rels"));
        foreach ($relationshipsXML->Relationship as $rel) {
            if ($rel['Type'] == self::SCHEMA_OFFICEDOCUMENT) {
                $workbookDir = dirname($rel['Target']) . '/';
                $workbookXML = simplexml_load_string($this->getEntryData($rel['Target']));
                foreach ($workbookXML->sheets->sheet as $sheet) {
                    $r                       = $sheet->attributes('r', true);
                    $sheets[(string) $r->id] = array(
                        'sheetId' => (int) $sheet['sheetId'],
                        'name'    => (string) $sheet['name'],
                    );

                }
                $workbookRelationsXML = simplexml_load_string($this->getEntryData($workbookDir . '_rels/' . basename($rel['Target']) . '.rels'));
                foreach ($workbookRelationsXML->Relationship as $wrel) {
                    switch ($wrel['Type']) {
                        case self::SCHEMA_WORKSHEETRELATION:
                            $sheets[(string) $wrel['Id']]['path'] = $workbookDir . (string) $wrel['Target'];
                            break;
                        case self::SCHEMA_SHAREDSTRINGS:
                            $sharedStringsXML = simplexml_load_string($this->getEntryData($workbookDir . (string) $wrel['Target']));
                            foreach ($sharedStringsXML->si as $val) {
                                if (isset($val->t)) {
                                    $this->sharedStrings[] = (string) $val->t;
                                } elseif (isset($val->r)) {
                                    $this->sharedStrings[] = XLSXWorksheet::parseRichText($val);
                                }
                            }
                            break;
                    }
                }
            }
        }
        $this->sheetInfo = [];
        foreach ($sheets as $rid => $info) {
            $this->sheetInfo[$info['name']] = array(
                'sheetId' => $info['sheetId'],
                'rid'     => $rid,
                'path'    => $info['path'],
            );
        }
    }

    /**
     * 返回sheet的名称数组，格式：
     *  [
     *      'sheet_id' => 'sheet_name',
     *  ]
     *
     * @return array
     */
    public function getSheetNames()
    {
        $res = array();
        foreach ($this->sheetInfo as $sheetName => $info) {
            $res[$info['sheetId']] = $sheetName;
        }
        return $res;
    }

    /**
     * 获取sheet数量
     *
     * @return int
     */
    public function getSheetCount()
    {
        return count($this->sheetInfo);
    }

    // instantiates a sheet object (if needed) and returns an array of its data
    public function getSheetData($sheetNameOrId)
    {
        $sheet = $this->getSheet($sheetNameOrId);
        return $sheet->getData();
    }

    // instantiates a sheet object (if needed) and returns the sheet object
    public function getSheet($sheet)
    {
        if (is_numeric($sheet)) {
            $sheet = $this->getSheetNameById($sheet);
        } elseif (!is_string($sheet)) {
            throw new Exception("Sheet must be a string or a sheet Id");
        }
        if (!array_key_exists($sheet, $this->sheets)) {
            $this->sheets[$sheet] = new XLSXWorksheet($this->getSheetXML($sheet), $sheet, $this);

        }
        return $this->sheets[$sheet];
    }

    public function getSheetNameById($sheetId)
    {
        foreach ($this->sheetInfo as $sheetName => $sheetInfo) {
            if ($sheetInfo['sheetId'] === $sheetId) {
                return $sheetName;
            }
        }
        throw new Exception("Sheet ID $sheetId does not exist in the Excel file");
    }

    protected function getSheetXML($name)
    {
        return simplexml_load_string($this->getEntryData($this->sheetInfo[$name]['path']));
    }

    // converts an Excel date field (a number) to a unix timestamp (granularity: seconds)
    public static function toUnixTimeStamp($excelDateTime)
    {
        if (!is_numeric($excelDateTime)) {
            return $excelDateTime;
        }
        $d = floor($excelDateTime); // seconds since 1900
        $t = $excelDateTime - $d;
        return ($d > 0) ? ($d - 25569) * 86400 + $t * 86400 : $t * 86400;
    }

}
