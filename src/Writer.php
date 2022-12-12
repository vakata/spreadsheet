<?php

namespace vakata\spreadsheet;

use vakata\spreadsheet\writer\CSVWriter;
use vakata\spreadsheet\writer\DriverInterface;
use vakata\spreadsheet\writer\XLSXWriter;
use vakata\spreadsheet\writer\XMLWriter;

/** @phpstan-consistent-constructor */
class Writer
{
    protected DriverInterface $driver;

    public function __construct(mixed $stream, string $format, array $options = [])
    {
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
    public static function toFile(string $path, ?string $format = null, array $options = []): static
    {
        return new static(fopen($path, 'wb'), $format ?? strtolower(substr($path, strrpos($path, '.') + 1)), $options);
    }
    public static function toStream(mixed $stream, string $format = 'xlsx', array $options = []): static
    {
        return new static($stream, $format, $options);
    }
    public static function toBrowser(string $format = 'xlsx', array $options = [], ?string $filename = null): static
    {
        if ($filename) {
            foreach (static::headers($format, $filename) as $k => $v) {
                header($k . ': ' . $v);
            }
        }
        return static::toStream(fopen('php://output', 'wb'), $format, $options);
    }

    public static function headers(string $format, ?string $filename = null): array
    {
        $headers = [];
        switch ($format) {
            case 'xlsx':
                $headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                break;
            case 'csv':
                $headers['Content-Type'] = 'text/csv; charset=utf-8';
                break;
            case 'xml':
                $headers['Content-Type'] = 'text/xml; charset=utf-8';
                break;
            default:
                throw new Exception('Unsupported format');
        }
        if ($filename) {
            $headers['Content-Disposition'] = 'attachment; ' .
                'filename="' . preg_replace('([^a-z0-9.-]+)i', '_', $filename) . '"; ' .
                'filename*=UTF-8\'\'' . rawurlencode($filename);
        }
        return $headers;
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
