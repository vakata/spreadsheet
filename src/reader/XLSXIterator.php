<?php

namespace vakata\spreadsheet\reader;

use vakata\spreadsheet\Exception;

class XLSXIterator implements \Iterator
{
    protected \ZipArchive $zip;
    protected array $strings = [];
    protected array $styles = [];
    protected string $path;
    protected mixed $stream;
    protected string $rest = '';
    protected string $def  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet>';
    protected ?int $dim  = null;
    protected ?array $row  = null;
    protected int $ind  = -1;

    public function __construct(string $file, mixed $desiredSheet = null)
    {
        $this->zip = new \ZipArchive();
        if (!$this->zip->open($file)) {
            throw new Exception('Document not readable');
        }
        $sheets = [];
        // this should be safe to extract as string (not stream)
        $relationships = simplexml_load_string(
            $this->zip->getFromName("_rels/.rels") ?: throw new Exception('Could not read zip')
        ) ?: throw new Exception('Could not parse xml');
        foreach ($relationships->Relationship as $relationship) {
            // phpcs:ignore
            if ($relationship['Type'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument') {
                // this should be safe to extract as string (not stream)
                $workbook = simplexml_load_string(
                    // phpcs:ignore
                    $this->zip->getFromName((string)$relationship['Target']) ?: throw new Exception('Could not read zip')
                ) ?: throw new Exception('Could not parse xml');
                foreach ($workbook->sheets->sheet as $sheet) {
                    $sheets[(string)$sheet->attributes('r', true)?->id] = array(
                        'id'   => (int)$sheet['sheetId'],
                        'name' => (string)$sheet['name']
                    );
                }
                // this should be safe to extract as string (not stream)
                $workbookRelations = simplexml_load_string(
                    $this->zip->getFromName(
                        dirname((string)$relationship['Target']) .
                        '/_rels/' .
                        basename((string)$relationship['Target']) . '.rels'
                    ) ?: throw new Exception('Could not read zip')
                ) ?: throw new Exception('Could not parse xml');
                foreach ($workbookRelations->Relationship as $workbookRelation) {
                    switch ($workbookRelation['Type']) {
                        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet':
                            $sheets[(string)$workbookRelation['Id']]['path'] =
                                dirname((string)$relationship['Target']) . '/' .
                                (string)$workbookRelation['Target'];
                            break;
                        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings':
                            // THIS MIGHT NOT BE SAFE TO EXTRACT AS A STRING - PROBABLY SWITCH TO INCREMENTAL READ?
                            $sharedStrings = simplexml_load_string(
                                $this->zip->getFromName(
                                    dirname((string)$relationship['Target']) . '/' . (string)$workbookRelation['Target']
                                    // phpcs:ignore
                                ) ?: throw new Exception('Could not read zip')
                            ) ?: throw new Exception('Could not parse xml');
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
                        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles':
                            $styles = simplexml_load_string(
                                $this->zip->getFromName(
                                    dirname((string)$relationship['Target']) . '/' . (string)$workbookRelation['Target']
                                ) ?: throw new Exception('Could not read zip')
                            ) ?: throw new Exception('Could not parse xml');
                            $i = 0;
                            foreach ($styles?->cellXfs[0]?->xf ?? [] as $k => $v) {
                                if (in_array((int)$v['numFmtId'], [14,20,22]) && (int)$v['applyNumberFormat']) {
                                    $this->styles[$i] = (int)$v['numFmtId'];
                                }
                                $i++;
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
            if ($sheet['name'] === $desiredSheet && isset($sheet['path'])) {
                $this->path = $sheet['path'];
            }
        }
        if (!isset($this->path) && is_numeric($desiredSheet)) {
            $sheets = array_values($sheets);
            if (isset($sheets[(int)$desiredSheet]) && isset($sheets[(int)$desiredSheet]['path'])) {
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
            $this->dim = (int)$this->getCellIndex(explode(':', explode('"', (string)$end)[1] ?? '')[1] ?? '');
        }
    }
    protected function getCellIndex(string $cell): int
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
    protected function getCellValue(mixed $cell): mixed
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
                if (!isset($cell->v)) {
                    return null;
                }
                $value = (string)$cell->v;
                switch ($this->styles[(int)$cell['s']] ?? 0) {
                    case 14:
                        $value = gmdate('Y-m-d', self::timestamp((float)$value));
                        break;
                    case 20:
                        $value = gmdate('H:i', self::timestamp((float)$value));
                        break;
                    case 22:
                        $value = gmdate('Y-m-d H:i:s', self::timestamp((float)$value));
                        break;
                    default:
                        // check for numeric values
                        if (is_numeric($value)) {
                            if ($value == (int)$value) {
                                $value = (int)$value;
                            } elseif ($value == (float)$value) {
                                $value = (float)$value;
                            } elseif ($value == (double)$value) {
                                $value = (double)$value;
                            }
                        }
                        break;
                }
                return $value;
        }
    }
    protected function read(string $end): ?string
    {
        $size = strlen($end);
        $data = $this->rest;
        while (!feof($this->stream)) {
            $data .= fread($this->stream, $size);
            if (strpos($data, $end) !== false && $end) {
                $temp = explode($end, $data, 2);
                $this->rest = $temp[1];
                return $temp[0] . $end;
            }
        }
        return null;
    }
    public function current(): mixed
    {
        return $this->row;
    }
    public function key(): mixed
    {
        return $this->ind;
    }
    public function next(): void
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
        $data = simplexml_load_string($this->def . '<row ' . $temp . '</worksheet>') ?: throw new Exception('XML');
        $this->ind = (int)$data->row['r'];
        $this->row = array_fill(0, $this->dim ?? 0, null);
        foreach ($data->row->c as $cell) {
            $this->row[$this->getCellIndex((string)$cell['r']) - 1] = $this->getCellValue($cell);
        }
    }
    public function rewind(): void
    {
        $this->stream = $this->zip->getStream($this->path);
        $this->rest = '';
        $this->row = null;
        $this->ind = -1;
        $this->next();
    }
    public function valid(): bool
    {
        return $this->row !== null;
    }
    public static function timestamp(float|int $excelDateTime): int
    {
        $d = floor($excelDateTime); // days since 1900 or 1904
        $t = $excelDateTime - $d;

        $t = (abs($d) > 0) ? ($d - 25569) * 86400 + round($t * 86400) : round($t * 86400);

        return (int)$t;
    }
}
