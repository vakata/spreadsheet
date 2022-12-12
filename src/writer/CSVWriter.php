<?php

namespace vakata\spreadsheet\writer;

use vakata\spreadsheet\Exception;

class CSVWriter implements DriverInterface
{
    /**
     * @var resource
     */
    protected mixed $stream;
    /**
     * @var array<string,mixed>
     */
    protected array $options;

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
                'separator' => ',',
                'enclosure' => '"',
                'escape' => '\\',
                'eol' => "\n",
                'excel' => false
            ],
            $options
        );
        if ($this->options['excel']) {
            $this->options['separator'] = ';';
            $this->options['enclosure'] = '"';
            $this->options['escape'] = '\\';
            $this->options['eol'] = "\r\n";
            fputs($this->stream, (chr(0xEF) . chr(0xBB) . chr(0xBF)));
        }
    }
    public function setSheetName(string $name): DriverInterface
    {
        throw new Exception('Operation not supported');
    }
    public function addSheet(string $name): DriverInterface
    {
        throw new Exception('Operation not supported');
    }
    public function addRow(array $data): DriverInterface
    {
        fputcsv(
            $this->stream,
            $data,
            $this->options['separator'],
            $this->options['enclosure'],
            $this->options['escape'],
            $this->options['eol']
        );
        return $this;
    }
    public function close(): void
    {
        fclose($this->stream);
    }
}
