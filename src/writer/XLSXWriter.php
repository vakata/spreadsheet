<?php

namespace vakata\spreadsheet\writer;

use DateTime;
use RuntimeException;
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
    protected array $styles = [];

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
        $this->styles[] = '/////';
        $this->styles[] = 'd/////';
        $this->styles[] = 't/////';
        $this->styles[] = 'dt/////';
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
    protected function getStyle(
        string $type = '',
        string $format = '',
        string $borders = 'LTRB',
        string $fill = '',
        string $alignment = 'B',
        string $color = ''
    ): int {
        $type = in_array(strtolower($type), ['d', 't', 'dt']) ? strtolower($type) : '';
        $format = str_split(preg_replace('([^bui])', '', strtolower($format)) ?? '');
        sort($format);
        $format = implode('', $format);
        $borders = str_split(strtoupper($borders));
        sort($borders);
        $borders = implode('', $borders);
        if (strlen($fill) && strlen($fill) < 8) {
            $fill = strtoupper(str_pad($fill, 8, 'F', STR_PAD_LEFT));
        }
        $alignment = str_split(preg_replace('([^LRCTBM])', '', str_replace('B', '', strtoupper($alignment))) ?? '');
        sort($alignment);
        $alignment = implode($alignment);
        if (strlen($color) && strlen($color) < 8) {
            $color = strtoupper(str_pad($color, 8, 'F', STR_PAD_LEFT));
        }
        $key = implode('/', [$type, $format, $borders, $fill, $alignment, $color]);
        if (array_search($key, $this->styles) === false) {
            $this->styles[] = $key;
        }
        return (int)array_search($key, $this->styles);
    }

    public function addSheet(string $name, array $options = []): DriverInterface
    {
        $id = count($this->sheets) + 1;
        $fp = fopen($this->temp . '/xl/worksheets/sheet' . $id . '.xml', 'w');
        $tm = fopen($this->temp . '/xl/worksheets/sheet' . $id . '.xml.tmp', 'w');
        if (!$fp || !$tm) {
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
            'merges'        => [],
            'merged'        => [],
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
    protected function getSharedStringValue(string $key): int
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
    protected function getCellFromIndex(int $k): string
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
    public function addHeaderRow(
        array $data,
        bool $filter = true,
        bool $freeze = true,
        string $format = 'b',
        string $borders = 'LTBR',
        string $fill = 'EBEBEB',
        string $alignment = 'B',
        string $color = ''
    ): DriverInterface {
        $row = $this->addRow($data, $format, $borders, $fill, $alignment, $color);
        if ($filter) {
            $this->sheets[$this->activeSheet]['filterRow'] = $this->sheets[$this->activeSheet]['count'];
        }
        if ($freeze) {
            $this->sheets[$this->activeSheet]['freezeRow'] = $this->sheets[$this->activeSheet]['count'];
        }
        return $row;
    }
    public function addRow(
        array $data,
        string $format = '',
        string $borders = 'LTBR',
        string $fill = '',
        string $alignment = 'B',
        string $color = ''
    ): DriverInterface {
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
        foreach (array_values($data) as $k => $value) {
            $ft = $format;
            $bs = $borders;
            $fl = $fill;
            $al = $alignment;
            $cl = $color;
            $v = $value;
            $cell = $this->getCellFromIndex($k) . $this->sheets[$this->activeSheet]['count'];
            if (is_array($v)) {
                $v = array_values($value)[0];
                $ft = array_values($value)[1] ?? $ft;
                $bs = array_values($value)[2] ?? $bs;
                $fl = array_values($value)[3] ?? $fl;
                $al = array_values($value)[4] ?? $al;
                $cl = array_values($value)[5] ?? $cl;
                $spanC = array_values($value)[6] ?? 1;
                $spanR = array_values($value)[7] ?? 1;
                $style = $this->getStyle('', $ft, $bs, $fl, $al, $cl);
                if ($spanC > 1 || $spanR > 1) {
                    $this->sheets[$this->activeSheet]['merges'][] = $cell . ':' .
                        $this->getCellFromIndex($k + $spanC - 1) .
                        ($this->sheets[$this->activeSheet]['count'] + $spanR - 1);
                    for ($mr = 0; $mr < $spanR; $mr++) {
                        for ($mc = 0; $mc < $spanC; $mc++) {
                            $key  = $this->getCellFromIndex($k + $mc);
                            $key .= ($this->sheets[$this->activeSheet]['count'] + $mr);
                            $this->sheets[$this->activeSheet]['merged'][$key] = $style;
                        }
                    }
                    unset($this->sheets[$this->activeSheet]['merged'][$cell]);
                }
            }
            if (isset($this->sheets[$this->activeSheet]['merged'][$cell])) {
                $content .= '<c r="' . $this->escape((string)$cell, true) . '"';
                if ($this->sheets[$this->activeSheet]['merged'][$cell]) {
                    $content .= ' s="' .
                        $this->escape((string)$this->sheets[$this->activeSheet]['merged'][$cell], true) .
                        '"';
                }
                $content .= ' />';
                continue;
            }
            if (
                $this->options['autoWidth'] &&
                (
                    !isset($this->sheets[$this->activeSheet]['widths'][$k]) ||
                    mb_strlen($v, 'UTF-8') + 4 > $this->sheets[$this->activeSheet]['widths'][$k]
                )
            ) {
                $this->sheets[$this->activeSheet]['widths'][$k] = mb_strlen($v, 'UTF-8') + 4;
            }
            $type = null;
            $style = '';
            if ($v instanceof DateTime) {
                $v = $v->format('c');
            }
            switch (gettype($v)) {
                case 'boolean':
                    $type = 'b';
                    break;
                case 'integer':
                case 'double':
                    $type = 'n';
                    break;
                default:
                    $v = (string)$v;
                    if (
                        (
                            preg_match('(^\d\d\.\d\d\.\d\d\d\d$)', $v) ||
                            preg_match('(^\d\d\d\d\-\d\d-\d\d$)', $v)
                        ) && $tmp = strtotime($v)
                    ) {
                        $type = null;
                        $style = 'd';
                        $v = $this->excelDate((int)date('Y', $tmp), (int)date('n', $tmp), (int)date('j', $tmp));
                    } elseif (
                        // phpcs:ignore
                        preg_match('(^\d\d\d\d-\d\d-\d\d[ T]\d\d:\d\d:\d\d(\.\d+)? ?([\+-][0-9]{2}(:[0-9]{2})?|Z|[a-z/_]+)?$)i', $v) &&
                        $tmp = strtotime($v)
                    ) {
                        $type = null;
                        $style = 'dt';
                        $v = $this->excelDate(
                            (int)date('Y', $tmp),
                            (int)date('n', $tmp),
                            (int)date('j', $tmp),
                            (int)date('G', $tmp),
                            (int)ltrim(date('i', $tmp), '0'),
                            (int)ltrim(date('s', $tmp), '0')
                        );
                    } elseif (
                        (preg_match('(^\d\d:\d\d:\d\d(\.\d+)$)i', $v) || preg_match('(^\d\d:\d\d$)i', $v)) &&
                        $tmp = strtotime($v)
                    ) {
                        $type = null;
                        $style = 't';
                        $v = $this->excelDate(
                            0,
                            0,
                            0,
                            (int)date('G', $tmp),
                            (int)date('i', $tmp),
                            (int)date('s', $tmp)
                        );
                    } else {
                        if ($this->options['sharedStrings']) {
                            $type = 's';
                            $temp = '';
                            // if (strlen($ft)) {
                            //     $temp = chr(0) . $ft;
                            // }
                            $v = $this->getSharedStringValue($v . $temp);
                        } else {
                            $type = 'inlineStr';
                        }
                    }
                    break;
            }
            if ($this->sheets[$this->activeSheet]['maxCell'] < $k) {
                $this->sheets[$this->activeSheet]['maxCell'] = $k;
            }
            $content .= '<c r="' . $this->escape((string)$cell, true) . '"';
            if (isset($type)) {
                $content .= ' t="' . $this->escape((string)$type, true) . '"';
            }
            $style = $this->getStyle($style, $ft, $bs, $fl, $al, $cl);
            if ($style) {
                $content .= ' s="' . $this->escape((string)$style, true) . '"';
            }
            $content .= '>';
            if ($type !== 'inlineStr') {
                $content .= '<v>' . $this->escape((string)$v) . '</v>';
            } else {
                $content .= '<is><t>' . $this->escape((string)$v) . '</t></is>';
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

        $fonts = [
            '/' => '<font><name val="Calibri"/><family val="2"/></font>'
        ];
        $colors = [];
        $fills = [
            '' => '<fill><patternFill patternType="none"/></fill>',
            'gray' => '<fill><patternFill patternType="gray125"/></fill>'
        ];
        $borders = [
            '' => '<border><left/><right/><top/><bottom/><diagonal/></border>'
        ];
        $styles = [];
        foreach ($this->styles as $k => $style) {
            $parts = explode('/', $style);
            $nm = ([''=>0,'d'=>14,'t'=>20,'dt'=>22])[$parts[0]];
            $ft = $parts[1];
            $bs = $parts[2];
            if (!isset($borders[$bs])) {
                $borders[$bs] = '' .
                    '<border>' .
                    (strpos($bs, 'L') !== false ? '<left style="thin"><color indexed="64"/></left>' : '<left/>') .
                    (strpos($bs, 'R') !== false ? '<right style="thin"><color indexed="64"/></right>' : '<right/>') .
                    (strpos($bs, 'T') !== false ? '<top style="thin"><color indexed="64"/></top>' : '<top/>') .
                    (strpos($bs, 'B') !== false ? '<bottom style="thin"><color indexed="64"/></bottom>' : '<bottom/>') .
                    (strpos($bs, 'D') !== false ?
                        '<diagonal style="thin"><color indexed="64"/></diagonal>' : '<diagonal/>'
                    ) .
                    '</border>';
            }
            $bs = array_search($bs, array_keys($borders));
            $fl = $parts[3];
            if (!isset($fills[$fl])) {
                $fills[$fl] = '' .
                    '<fill><patternFill patternType="solid"><fgColor rgb="'.$fl.'"/></patternFill></fill>';
            }
            $fl = array_search($fl, array_keys($fills));
            $cl = $ft . '/' . $parts[5];
            if (!isset($fonts[$cl])) {
                $fonts[$cl] = '' .
                    '<font>' .
                    (strpos($ft, 'b') !== false ? '<b/>' : '') .
                    (strpos($ft, 'u') !== false ? '<u/>' : '') .
                    (strpos($ft, 'i') !== false ? '<i/>' : '') .
                    '<name val="Calibri"/><family val="2"/>' .
                    ($parts[5] ? '<color rgb="'.$parts[5].'"/>' : '' ) .
                    '</font>';
            }
            if (strlen($parts[5])) {
                $colors[] = '<color rgb="'.$parts[5].'"/>';
            }
            $cl = array_search($cl, array_keys($fonts));
            $styles[$k] = '<xf xfId="0" fontId="'.$cl.'" ' .
                ($cl ? ' applyFont="1" ' : '') .
                ($nm ? ' numFmtId="'.$nm.'" applyNumberFormat="1" ' : '') .
                ($bs ? ' borderId="'.$bs.'" applyBorder="1" ' : '') .
                ($fl ? ' fillId="'.$fl.'" applyFill="1" ' : '');
            if ($parts[4]) {
                $styles[$k] .= ' applyAlignment="1">';
                $styles[$k] .= '<alignment ';
                if (strpos($parts[4], 'L') !== false) {
                    $styles[$k] .= ' horizontal="left" ';
                } elseif (strpos($parts[4], 'R') !== false) {
                    $styles[$k] .= ' horizontal="right" ';
                } elseif (strpos($parts[4], 'C') !== false) {
                    $styles[$k] .= ' horizontal="center" ';
                }
                if (strpos($parts[4], 'T') !== false) {
                    $styles[$k] .= ' vertical="top" ';
                } elseif (strpos($parts[4], 'B') !== false) {
                    $styles[$k] .= ' vertical="bottom" ';
                } elseif (strpos($parts[4], 'M') !== false) {
                    $styles[$k] .= ' vertical="center" ';
                }
                $styles[$k] .= ' /></xf>';
            } else {
                $styles[$k] .= '/>';
            }
        }
        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fonts count="'.count($fonts).'">'.implode('', $fonts).'</fonts>
        <fills count="'.count($fills).'">'.implode('', $fills).'</fills>
        <borders count="'.count($borders).'">'.implode('', $borders).'</borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs>
        <cellXfs count="'.count($styles).'">'.implode('', $styles).'</cellXfs>
        <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
        </cellStyles>
        <colors><mruColors>'.implode('', $colors).'</mruColors></colors>
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
                    '<pane ySplit="' . $sheet['freezeRow'] . '" topLeftCell="A' . ($sheet['freezeRow'] + 1) . '" ' .
                    ' activePane="bottomLeft" state="frozen"/>' .
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
                    fwrite(
                        $sheet['stream'],
                        '<col min="' . ($kw + 1) . '" max="' . ($kw + 1) . '" width="' . sprintf('%01.3f', $w) . '" ' .
                            ' bestFit="1" customWidth="1" />'
                    );
                }
                fwrite($sheet['stream'], '</cols>');
            }
            fwrite(
                $sheet['streamd'],
                '</sheetData>'
            );
            fclose($sheet['streamd']);
            stream_copy_to_stream(
                fopen(
                    $this->temp . '/xl/worksheets/sheet' . $sheet['id'] . '.xml.tmp',
                    'r'
                ) ?: throw new RuntimeException(),
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
            if (count($sheet['merges'])) {
                $content .= '<mergeCells count="' . count($sheet['merges']) . '">';
                foreach ($sheet['merges'] as $merge) {
                    $content .= '<mergeCell ref="' . $merge . '" />';
                }
                $content .= '</mergeCells>';
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
