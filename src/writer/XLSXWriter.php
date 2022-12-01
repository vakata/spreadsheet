<?php

namespace vakata\spreadsheet\writer;

use ZipArchive;

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
                'created' => date('c')
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
    }

    protected function escape(string $input, bool $attr = false): string
    {
        return $attr ?
            htmlspecialchars($input, ENT_XML1 | ENT_COMPAT, 'UTF-8') :
            htmlspecialchars($input, ENT_XML1, 'UTF-8');
    }

    public function addSheet(string $name): DriverInterface
    {
        $id = count($this->sheets) + 1;
        $fp = fopen($this->temp . '/xl/worksheets/sheet' . $id . '.xml', 'w');
        $this->sheets[$id] = [
            'id'   => $id,
            'name' => $name,
            'count' => 0,
            'stream' => $fp
        ];
        fwrite(
            $fp,
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\r\n" .
                '<worksheet ' .
                ' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' .
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' .
                '><sheetData>'
        );
        $this->activeSheet = $id;
        return $this;
    }
    public function addRow(array $data): DriverInterface
    {
        if (!$this->activeSheet) {
            $this->addSheet('Sheet1');
        }
        $fp = $this->sheets[$this->activeSheet]['stream'];
        $content = '<row r="' . (++$this->sheets[$this->activeSheet]['count']) . '" ';
        $content .= ' spans="1:' . count($data) . '">';
        foreach (array_values($data) as $k => $value) {
            $type = null;
            switch (gettype($value)) {
                case 'boolean':
                    $type = 'b';
                    break;
                case 'integer':
                case 'double':
                    $type = 'n';
                    break;
                default:
                    $type = 'inlineStr';
                    $value = (string)$value;
                    if (
                        (
                            preg_match('(^\d\d\.\d\d\.\d\d\d\d$)', $value) ||
                            preg_match('(^\d\d\d\d\-\d\d-\d\d$)', $value) ||
                            preg_match('(^\d\d\d\d-\d\d-\d\d[ T]\d\d:\d\d:\d\d(\.\d+)? ?([\+-][0-9]{2}(:[0-9]{2})?|Z|[a-z/_]+)?$)i', $value)
                        ) && strtotime($value)
                    ) {
                        $type = 'c';
                        $value = date('c', strtotime($value));
                    }
                    break;
            }
            $cell = '';
            $index = $k + 1;
            while ($index !== 0) {
                $temp = (($index - 1) % 26);
                $index = (int) (($index - $temp) / 25);
                $cell = chr(65 + $temp) . $cell;
            }
            $cell .= $this->sheets[$this->activeSheet]['count'];
            $content .= '<c r="' . $this->escape($cell, true) . '" t="' . $this->escape($type, true) . '">';
            if ($type !== 'inlineStr') {
                $content .= '<v>' . $this->escape($value) . '</v>';
            } else {
                $content .= '<is><t>' . $this->escape($value) . '</t></is>';
            }
            $content .= '</c>';
        }
        $content .= '</row>';
        fwrite($fp, $content);
        return $this;
    }

    public function close() : void
    {
        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\r\n" .
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' .
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' .
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' .
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' .
        '</Relationships>';
        file_put_contents($this->temp . '/_rels/.rels', $content);

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
                '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
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
        $content .= '</Relationships>';
        file_put_contents($this->temp . '/xl/_rels/workbook.xml.rels', $content);

        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\r\n" .
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' .
            '<sheets>';
        foreach ($this->sheets as $sheet) {
            $content .= '<sheet name="' . $sheet['name'] . '" sheetId="' . $sheet['id'] . '" r:id="rId' . $sheet['id'] . '"/>';
        }
        $content .= '</sheets>' . '</workbook>';
        file_put_contents($this->temp . '/xl/workbook.xml', $content);

        foreach ($this->sheets as $sheet) {
            $content = '</sheetData>' . '</worksheet>';
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
        $temp = realpath($this->temp);
        $index = 0;
        foreach ($iterator as $file) {
            if ($file->isFile() && $file->getFilename() !== 'xlsx.zip') {
                $filePath = $file->getRealPath();
                $relativePath = substr($filePath, strlen($temp) + 1);
                $zip->addFile($filePath, $relativePath);
                $zip->setCompressionIndex($index, ZipArchive::CM_STORE);
                $index++;
            }
        }
        $zip->close();

        $zip = fopen($path, 'r');
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
}