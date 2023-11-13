<?php

namespace vakata\spreadsheet\writer;

use DateTime;
use ZipArchive;
use vakata\spreadsheet\Exception;

class XLSXWriter implements DriverInterface
{
    /**
     * @var resource
     */
    protected mixed $stream;
    /**
     * @var array<string,mixed>
     */
    protected array $options;
    protected string $temp;
    protected array $sheets = [];
    protected ?int $activeSheet = null;
    protected array $sharedStrings = [];

    /**
     * @param mixed $stream
     * @param array<string,mixed> $options
     * @return void
     */
    public function __construct(mixed $stream, array $options = [])
    {
        $this->stream = $stream;
        $this->options = array_merge(
            [
                'temp' => sys_get_temp_dir(),
                'user' => 'XLSXWriter',
                'created' => date('c'),
                'defaultSheet' => null,
                'sharedStrings' => true,
                'autoWidth' => true,
                'minWidth' => 4,
                'maxWidth' => null
            ],
            $options
        );
        do {
            $this->temp = $this->options['temp'] . DIRECTORY_SEPARATOR . 'xlsx_' . microtime(true) . '_' . uniqid();
        } while (is_dir($this->temp));
        mkdir($this->temp);
        chmod($this->temp, 0775);
        foreach ([ '_rels', 'docProps', 'xl/_rels', 'xl/worksheets' ] as $dir) {
            mkdir($this->temp . DIRECTORY_SEPARATOR . $dir, 0775, true);
            chmod($this->temp . DIRECTORY_SEPARATOR . $dir, 0775);
        }
        if (isset($this->options['defaultSheet']) && is_string($this->options['defaultSheet'])) {
            $this->addSheet($this->options['defaultSheet']);
        }
    }

    protected function escape(string $input, bool $attr = false): string
    {
        return $attr ?
            htmlspecialchars($input, ENT_XML1 | ENT_COMPAT, 'UTF-8') :
            htmlspecialchars($input, ENT_XML1, 'UTF-8');
    }

    public function addSheet(string $name, array $options = []): DriverInterface
    {
        $id = count($this->sheets) + 1;
        $fp = fopen($this->temp . '/xl/worksheets/sheet' . $id . '.xml', 'w');
        $tm = fopen($this->temp . '/xl/worksheets/sheet' . $id . '.xml.tmp', 'w');
        if (!$fp) {
            throw new Exception('Could not open temp file');
        }
        $this->sheets[$id] = [
            'id'            => $id,
            'name'          => $name,
            'count'         => 0,
            'stream'        => $fp,
            'streamd'       => $tm,
            'minRow'        => 0,
            'maxRow'        => 0,
            'minCell'       => 0,
            'maxCell'       => 0,
            'freezeRow'     => null,
            'filterRow'     => null,
            'widths'        => $options['widths'] ?? [],
            'autoWidth'     => $options['autoWidth'] ?? $this->options['autoWidth'] ?? true,
            'minWidth'      => $options['minWidth'] ?? $this->options['minWidth'] ?? 4,
            'maxWidth'      => $options['maxWidth'] ?? $this->options['maxWidth'] ?? null
        ];
        fwrite(
            $fp,
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\r\n" .
                '<worksheet ' .
                ' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' .
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' .
                '>'
        );
        fwrite($tm, '<sheetData>');
        $this->activeSheet = $id;
        return $this;
    }
    protected function getSharedStringValue(string $key) : int
    {
        if (!isset($this->sharedStrings[$key])) {
            $this->sharedStrings[$key] = [
                'count' => 0,
                'index' => count($this->sharedStrings)
            ];
        }
        $this->sharedStrings[$key]['count']++;

        return $this->sharedStrings[$key]['index'];
    }
    protected function getCellFromIndex(int $k) : string
    {
        $cell = '';
        $index = $k + 1;
        while ($index !== 0) {
            $temp = (($index - 1) % 26);
            $index = (int) (($index - $temp) / 25);
            $cell = chr(65 + $temp) . $cell;
        }

        return $cell;
    }
    public function addRow(
        array $data,
        bool $header = false,
        bool $filter = false,
        bool $freeze = false
    ): DriverInterface
    {
        if (!$this->activeSheet) {
            $this->addSheet('Sheet1');
        }
        $fp = $this->sheets[$this->activeSheet]['streamd'];
        $content = '<row r="' . (++$this->sheets[$this->activeSheet]['count']) . '" ';
        //$content .= ' spans="1:' . count($data) . '"';
        $content .= '>';
        if ($this->sheets[$this->activeSheet]['maxRow'] < $this->sheets[$this->activeSheet]['count']) {
            $this->sheets[$this->activeSheet]['maxRow'] = $this->sheets[$this->activeSheet]['count'];
        }
        if ($filter) {
            $this->sheets[$this->activeSheet]['filterRow'] = $this->sheets[$this->activeSheet]['count'];
        }
        if ($freeze) {
            $this->sheets[$this->activeSheet]['freezeRow'] = $this->sheets[$this->activeSheet]['count'];
        }
        foreach (array_values($data) as $k => $value) {
            if ($header) {
                $this->options['sharedStrings'] = true;
                if (!is_array($value)) {
                    $value = (string)$value;
                }
            }
            $v = is_array($value) ? (string)(array_values($value)[0]) : (string)$value;
            if (
                $this->options['autoWidth'] &&
                (
                    !isset($this->sheets[$this->activeSheet]['widths'][$k]) ||
                    mb_strlen($v, 'UTF-8') + ($filter ? 1 : 0) > $this->sheets[$this->activeSheet]['widths'][$k]
                )
            ) {
                $this->sheets[$this->activeSheet]['widths'][$k] = mb_strlen($v, 'UTF-8') + ($filter ? 1 : 0);
            }
            $type = null;
            $style = null;
            if ($value instanceof DateTime) {
                $value = $value->format('c');
            }
            switch (gettype($value)) {
                case 'boolean':
                    $type = 'b';
                    break;
                case 'integer':
                case 'double':
                    $type = 'n';
                    break;
                default:
                    $format = '';
                    if (is_array($value) && count($value) === 2) {
                        $format = array_values($value)[1];
                        $value = array_values($value)[0];
                    }
                    if ($header) {
                        $format .= 'b';
                    }
                    $value = (string)$value;
                    if (
                        (
                            preg_match('(^\d\d\.\d\d\.\d\d\d\d$)', $value) ||
                            preg_match('(^\d\d\d\d\-\d\d-\d\d$)', $value)
                        ) && $v = strtotime($value)
                    ) {
                        $type = null;
                        $style = 1;
                        $value = $this->excelDate((int)date('Y', $v), (int)date('n', $v), (int)date('j', $v));
                    } elseif (
                        // phpcs:ignore
                        preg_match('(^\d\d\d\d-\d\d-\d\d[ T]\d\d:\d\d:\d\d(\.\d+)? ?([\+-][0-9]{2}(:[0-9]{2})?|Z|[a-z/_]+)?$)i', $value) &&
                        $v = strtotime($value)
                    ) {
                        $type = null;
                        $style = 3;
                        $value = $this->excelDate(
                            (int)date('Y', $v),
                            (int)date('n', $v),
                            (int)date('j', $v),
                            (int)date('G', $v),
                            (int)ltrim(date('i', $v), '0'),
                            (int)ltrim(date('s', $v), '0')
                        );
                    } elseif (
                        (preg_match('(^\d\d:\d\d:\d\d(\.\d+)$)i', $value) || preg_match('(^\d\d:\d\d$)i', $value)) &&
                        $v = strtotime($value)
                    ) {
                        $type = null;
                        $style = 2;
                        $value = $this->excelDate(0, 0, 0, (int)date('G', $v), (int)date('i', $v), (int)date('s', $v));
                    } else {
                        if ($this->options['sharedStrings']) {
                            $type = 's';
                            $temp = '';
                            if (strlen($format)) {
                                $temp = chr(0) . $format;
                            }
                            $value = $this->getSharedStringValue($value . $temp);
                        } else {
                            $type = 'inlineStr';
                        }
                    }
                    break;
            }
            if ($this->sheets[$this->activeSheet]['maxCell'] < $k) {
                $this->sheets[$this->activeSheet]['maxCell'] = $k;
            }
            $cell = $this->getCellFromIndex($k);
            $cell .= $this->sheets[$this->activeSheet]['count'];
            $content .= '<c r="' . $this->escape((string)$cell, true) . '"';
            if (isset($type)) {
                $content .= ' t="' . $this->escape((string)$type, true) . '"';
            }
            if (isset($style)) {
                $content .= ' s="' . $this->escape((string)$style, true) . '"';
            }
            $content .= '>';
            if ($type !== 'inlineStr') {
                $content .= '<v>' . $this->escape((string)$value) . '</v>';
            } else {
                $content .= '<is><t>' . $this->escape((string)$value) . '</t></is>';
            }
            $content .= '</c>';
        }
        $content .= '</row>';
        fwrite($fp, $content);
        return $this;
    }

    public function close(): void
    {
        // phpcs:disable
        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\r\n" .
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' .
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' .
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' .
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' .
        '</Relationships>';
        file_put_contents($this->temp . '/_rels/.rels', $content);

        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
            <TotalTime>0</TotalTime>
            <Application>XLSXWriter</Application></Properties>';
        file_put_contents($this->temp . '/docProps/app.xml', $content);

        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\r\n" .
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' .
                '<dc:creator>' . $this->escape($this->options['user']) . '</dc:creator>' .
                '<cp:lastModifiedBy>' . $this->escape($this->options['user']) . '</cp:lastModifiedBy>' .
                '<dcterms:created xsi:type="dcterms:W3CDTF">' . $this->escape($this->options['created']) . '</dcterms:created>' .
                '<dcterms:modified xsi:type="dcterms:W3CDTF">' . $this->escape($this->options['created']) . '</dcterms:modified>' .
            '</cp:coreProperties>';
        file_put_contents($this->temp . '/docProps/core.xml', $content);

        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\r\n" .
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' .
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' .
                '<Default Extension="xml" ContentType="application/xml"/>' .
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' .
                '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' .
                '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' .
                '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' .
                '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
        foreach ($this->sheets as $sheet) {
            $content .= '<Override PartName="/xl/worksheets/sheet' . $sheet['id'] . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        $content .= '</Types>';
        file_put_contents($this->temp . '/[Content_Types].xml', $content);

        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\r\n" .
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        foreach ($this->sheets as $sheet) {
            $content .= '<Relationship Id="rId' . $sheet['id'] . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $sheet['id'] . '.xml"/>';
        }
        $content .= '<Relationship Id="rId' . (count($this->sheets) + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        $content .= '<Relationship Id="rId' . (count($this->sheets) + 3) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
        $content .= '</Relationships>';
        file_put_contents($this->temp . '/xl/_rels/workbook.xml.rels', $content);

        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fonts count="1"><font><name val="Calibri"/><family val="2"/></font></fonts>
        <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
        <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs>
        <cellXfs count="4">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
            <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />
            <xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />
            <xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />
        </cellXfs>
        <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
        </cellStyles>
        </styleSheet>';
        file_put_contents($this->temp . '/xl/styles.xml', $content);

        if ($this->options['sharedStrings']) {
            $count = array_sum(
                array_map(
                    function (array $item) {
                        return $item['count'];
                    },
                    $this->sharedStrings
                )
            );
            $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\r\n" .
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' . $count . '" uniqueCount="' . count($this->sharedStrings) . '">';
            foreach ($this->sharedStrings as $value => $item) {
                $value = explode(chr(0), $value, 2);
                $content .= '<si>';
                    if (isset($value[1]) && strlen($value[1])) {
                        $content .= '<r><rPr>';
                        if (strpos($value[1], 'b') !== false) {
                            $content .= '<b />';
                        }
                        if (strpos($value[1], 'i') !== false) {
                            $content .= '<i />';
                        }
                        if (strpos($value[1], 'u') !== false) {
                            $content .= '<u />';
                        }
                        $content .= '</rPr>';
                    }
                    $content .='<t>' . $this->escape($value[0]) . '</t>';
                    if (isset($value[1]) && strlen($value[1])) {
                        $content .= '</r>';
                    }
                $content .= '</si>';
            }
            $content .= '</sst>';
            file_put_contents($this->temp . '/xl/sharedStrings.xml', $content);
        }

        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\r\n" .
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' .
            '<sheets>';
        foreach ($this->sheets as $sheet) {
            $content .= '<sheet name="' . $sheet['name'] . '" sheetId="' . $sheet['id'] . '" r:id="rId' . $sheet['id'] . '"/>';
        }
        $content .= '</sheets>' . '</workbook>';
        file_put_contents($this->temp . '/xl/workbook.xml', $content);
        // phpcs:enable

        foreach ($this->sheets as $sheet) {
            fwrite(
                $sheet['stream'],
                '<dimension ref="' .
                $this->escape((string) $this->getCellFromIndex($sheet['minCell']) . '1') .
                ':' .
                $this->escape((string) $this->getCellFromIndex($sheet['maxCell']) . $sheet['maxRow']) .
                '" />'
            );
            if ($sheet['freezeRow']) {
                fwrite(
                    $sheet['stream'],
                    '<sheetViews><sheetView tabSelected="1" workbookViewId="0">' .
                    '<pane ySplit="'.$sheet['freezeRow'].'" topLeftCell="A'.($sheet['freezeRow']+1).'" activePane="bottomLeft" state="frozen"/>' .
                    '<selection pane="bottomLeft"/>' .
                    '</sheetView></sheetViews>'
                );
            }
            if (count($sheet['widths'])) {
                fwrite($sheet['stream'], '<cols>');
                foreach ($sheet['widths'] as $kw => $w) {
                    // 5 is the font-width
                    $w = ((($w * 5 + 5) / 5 * 256) / 256);
                    $w = max($w, $sheet['minWidth'], 1);
                    if (isset($sheet['maxWidth'])) {
                        $w = min($w, $sheet['maxWidth']);
                    }
                    fwrite($sheet['stream'], '<col min="'.($kw+1).'" max="'.($kw+1).'" width="'.sprintf('%01.3f', $w).'" bestFit="1" customWidth="1" />');
                }
                fwrite($sheet['stream'], '</cols>');
            }
            fwrite(
                $sheet['streamd'],
                '</sheetData>'
            );
            fclose($sheet['streamd']);
            stream_copy_to_stream(
                fopen($this->temp . '/xl/worksheets/sheet' . $sheet['id'] . '.xml.tmp', 'r'),
                $sheet['stream']
            );
            unlink($this->temp . '/xl/worksheets/sheet' . $sheet['id'] . '.xml.tmp');
            $content = '';
            if ($sheet['filterRow']) {
                $content .= '<autoFilter ref="' .
                    $this->escape((string) $this->getCellFromIndex($sheet['minCell']) . $sheet['filterRow']) .
                    ':' .
                    $this->escape((string) $this->getCellFromIndex($sheet['maxCell']) . $sheet['maxRow']) .
                    '"></autoFilter>';
            }
            $content .= '</worksheet>';
            fwrite($sheet['stream'], $content);
            fclose($sheet['stream']);
        }

        $path = $this->temp . '/xlsx.zip';
        $zip = new ZipArchive();
        $zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE);
        $iterator = new \RecursiveIteratorIterator(
            new \RecursiveDirectoryIterator($this->temp, \RecursiveDirectoryIterator::SKIP_DOTS),
            \RecursiveIteratorIterator::LEAVES_ONLY
        );
        $temp = realpath($this->temp) ?: throw new Exception('Could not get temp file');
        $index = 0;
        foreach ($iterator as $file) {
            if ($file->isFile() && $file->getFilename() !== 'xlsx.zip') {
                $filePath = $file->getRealPath();
                $relativePath = substr($filePath, strlen($temp) + 1);
                $zip->addFile($filePath, str_replace('\\', '/', $relativePath));
                $zip->setCompressionIndex($index, ZipArchive::CM_DEFLATE);
                $index++;
            }
        }
        $zip->close();

        $zip = fopen($path, 'r') ?: throw new Exception('Could not open temp file');
        stream_copy_to_stream($zip, $this->stream);
        fclose($this->stream);
        fclose($zip);

        $iterator = new \RecursiveIteratorIterator(
            new \RecursiveDirectoryIterator($this->temp, \RecursiveDirectoryIterator::SKIP_DOTS),
            \RecursiveIteratorIterator::CHILD_FIRST
        );
        foreach ($iterator as $file) {
            if ($file->isFile()) {
                unlink($file->getPathname());
            } else {
                rmdir($file->getPathname());
            }
        }
        rmdir($this->temp);
    }

    public static function excelDate(
        int $year,
        int $month,
        int $day,
        int $hours = 0,
        int $minutes = 0,
        int $seconds = 0
    ): int|float {
        $excelTime = (($hours * 3600) + ($minutes * 60) + $seconds) / 86400;
        if ((int)$year === 0) {
            return $excelTime;
        }
        $excel1900isLeapYear = true;
        if (($year === 1900) && ($month <= 2)) {
            $excel1900isLeapYear = false;
        }
        $myExcelBaseDate = 2415020;

        // Julian base date Adjustment
        if ($month > 2) {
            $month -= 3;
        } else {
            $month += 9;
            --$year;
        }
        $century = substr((string)$year, 0, 2);
        $decade = substr((string)$year, 2, 2);
        $excelDate = floor((146097 * (int)$century) / 4) +
            floor((1461 * (int)$decade) / 4) +
            floor((153 * $month + 2) / 5) +
            $day + 1721119 - $myExcelBaseDate + ($excel1900isLeapYear ? 1 : 0);

        return (float)$excelDate + $excelTime;
    }
}
