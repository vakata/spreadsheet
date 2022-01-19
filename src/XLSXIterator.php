<?php

namespace vakata\spreadsheet;

class XLSXIterator implements \Iterator
{
    protected $zip;
    protected $strings = [];
    protected $path;
    protected $stream;
    protected $rest = '';
    protected $def  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet>';
    protected $dim  = null;
    protected $row  = null;
    protected $ind  = -1;

    public function __construct(string $file, $desiredSheet = null)
    {
        $this->zip = new \ZipArchive();
        if (!$this->zip->open($file)) {
            throw new Exception('Document not readable');
        }
        $sheets = [];
        // this should be safe to extract as string (not stream)
        $relationships = simplexml_load_string($this->zip->getFromName("_rels/.rels"));
        foreach ($relationships->Relationship as $relationship) {
            if ($relationship['Type'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument') {
                // this should be safe to extract as string (not stream)
                $workbook = simplexml_load_string($this->zip->getFromName($relationship['Target']));
                foreach ($workbook->sheets->sheet as $sheet) {
                    $sheets[(string)$sheet->attributes('r', true)->id] = array(
                        'id'   => (int)$sheet['sheetId'],
                        'name' => (string)$sheet['name']
                    );
                }
                // this should be safe to extract as string (not stream)
                $workbookRelations = simplexml_load_string(
                    $this->zip->getFromName(
                        dirname($relationship['Target']) . '/_rels/' . basename($relationship['Target']) . '.rels'
                    )
                );
                foreach ($workbookRelations->Relationship as $workbookRelation) {
                    switch ($workbookRelation['Type']) {
                        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet':
                            $sheets[(string)$workbookRelation['Id']]['path'] = dirname($relationship['Target']) . '/' . (string)$workbookRelation['Target'];
                            break;
                        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings':
                            // THIS MIGHT NOT BE SAFE TO EXTRACT AS A STRING - PROBABLY SWITCH TO INCREMENTAL READ?
                            $sharedStrings = simplexml_load_string(
                                $this->zip->getFromName(
                                    dirname($relationship['Target']) . '/' . (string)$workbookRelation['Target']
                                )
                            );
                            foreach ($sharedStrings->si as $val) {
                                if (isset($val->t)) {
                                    $this->strings[] = (string)$val->t;
                                } elseif (isset($val->r)) {
                                    $temp = [];
                                    foreach ($val->r as $part) {
                                        $temp[] = (string)$part->t;
                                    }
                                    $this->strings[] = implode(' ', $temp);
                                }
                            }
                            break;
                    }
                }
                break;
            }
        }
        if (!count($sheets)) {
            throw new Exception('No sheets found');
        }
        $desiredSheet = $desiredSheet ?? '0';
        foreach ($sheets as $sheet) {
            if ($sheet['name'] === $desiredSheet) {
                $this->path = $sheet['path'];
            }
        }
        if (!isset($this->path) && is_numeric($desiredSheet)) {
            $sheets = array_values($sheets);
            if (isset($sheets[(int)$desiredSheet])) {
                $this->path = $sheets[(int)$desiredSheet]['path'];
            }
        }
        if (!isset($this->path)) {
            throw new Exception('Sheet not found');
        }
        $this->stream = $this->zip->getStream($this->path);
        $beg = $this->read('<worksheet');
        if (isset($beg)) {
            $end = $this->read('>');
            $this->def = $beg . $end;
        }
        $beg = $this->read('<dimension');
        if (isset($beg)) {
            $end = $this->read('>');
            $this->dim = $this->getCellIndex(explode(':', explode('"', $end)[1])[1]);
        }
    }
    protected function getCellIndex($cell)
    {
        $matches = [];
        if (!preg_match("/([A-Z]+)(\d+)/", $cell, $matches)) {
            throw new Exception('Invalid column');
        }
        $col = $matches[1];
        $len = strlen($col);
        $ind = 0;
        for ($i = $len - 1; $i >= 0; $i--) {
            $ind += (ord($col[$i]) - 64) * pow(26, $len - $i - 1);
        }
        return $ind;
    }
    protected function getCellValue($cell)
    {
        // $cell['t'] is the cell type
        switch ((string)$cell["t"]) {
            case "s": // Value is a shared string
                return (string)$cell->v !== '' ? ($this->strings[intval($cell->v)] ?? '') : '';
            case "b": // Value is boolean
                $value = (string)$cell->v;
                if ($value === '0') {
                    return false;
                }
                if ($value === '1') {
                    return true;
                }
                return (bool)$cell->v;
            case "inlineStr": // Value is rich text inline
                $value = $cell->is;
                if (isset($value->t)) {
                    return (string)$value->t;
                }
                if (isset($value->r)) {
                    $temp = [];
                    foreach ($value->r as $part) {
                        $temp[] = (string)$part->t;
                    }
                    return implode(' ', $temp);
                }
                return '';
            case "e": // Value is an error message
                return (string)$cell->v;
            default:
                if(!isset($cell->v)) {
                    return null;
                }
                $value = (string)$cell->v;
                // Check for numeric values
                if (is_numeric($value)) {
                    if ($value == (int)$value) {
                        $value = (int)$value;
                    } else if ($value == (float)$value) {
                        $value = (float)$value;
                    } else if ($value == (double)$value) {
                        $value = (double)$value;
                    }
                }
                return $value;
        }
    }
    protected function read(string $end)
    {
        $size = strlen($end);
        $data = $this->rest;
        while (!feof($this->stream)) {
            $data .= fread($this->stream, $size);
            if (strpos($data, $end) !== false) {
                $temp = explode($end, $data, 2);
                $this->rest = $temp[1];
                return $temp[0] . $end;
            }
        }
        return null;
    }
    public function current()
    {
        return $this->row;
    }
    public function key()
    {
        return $this->ind;
    }
    public function next()
    {
        $temp = $this->read('<row ');
        if (!isset($temp)) {
            $this->row = null;
            return;
        }
        $temp = $this->read('</row>');
        if (!isset($temp)) {
            $this->row = null;
            return;
        }
        $data = simplexml_load_string($this->def . '<row ' . $temp . '</worksheet>');
        $this->ind = (int)$data->row['r'];
        $this->row = array_fill(0, $this->dim, null);
        foreach ($data->row->c as $cell) {
            $this->row[$this->getCellIndex($cell['r']) - 1] = $this->getCellValue($cell);
        }
    }
    public function rewind()
    {
        $this->stream = $this->zip->getStream($this->path);
        $this->rest = '';
        $this->row = null;
        $this->ind = -1;
        $this->next();
    }
    public function valid()
    {
        return $this->row !== null;
    }
}
