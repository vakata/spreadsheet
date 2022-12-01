<?php

namespace vakata\spreadsheet;

use vakata\spreadsheet\writer\CSVWriter;
use vakata\spreadsheet\writer\DriverInterface;
use vakata\spreadsheet\writer\XLSXWriter;
use XMLWriter;

class Writer
{
    protected DriverInterface $driver;

    public function __construct(mixed $stream, string $format, array $options = [])
    {
        $this->stream = $stream;
        switch ($format) {
            case 'csv':
                $this->driver = new CSVWriter($stream, $options);
                break;
            case 'xml':
                $this->driver = new XMLWriter($stream, $options);
                break;
            case 'xlsx':
                $this->driver = new XLSXWriter($stream, $options);
                break;
            default:
                throw new Exception('Unsupported format');
        }
    }
    public static function toFile(string $path, ?string $format = null, array $options = [])
    {
        return new static(fopen($path, 'wb'), $format ?? strtolower(substr($path, strrpos($path, '.') + 1)), $options);
    }
    public static function toStream(mixed $stream, string $format = 'xlsx', array $options = [])
    {
        return new static($stream, $format, $options);
    }
    public static function toBrowser(string $format = 'xlsx', array $options = [])
    {
        return static::toStream(fopen('php://output', 'wb'), $format, $options);
    }

    public function getDriver(): DriverInterface
    {
        return $this->driver;
    }
    public function fromArray(array $data): void
    {
        $this->fromIterable($data);
    }
    public function fromIterable(iterable $data): void
    {
        foreach ($data as $row) {
            $this->driver->addRow($row);
        }
        $this->driver->close();
    }
}
