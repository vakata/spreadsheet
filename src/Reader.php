<?php

namespace vakata\spreadsheet;

use Traversable;
use vakata\spreadsheet\reader\CSVIterator;
use vakata\spreadsheet\reader\XLSIterator;
use vakata\spreadsheet\reader\XLSXIterator;
use vakata\spreadsheet\reader\XMLIterator;

/** @phpstan-consistent-constructor */
class Reader implements \IteratorAggregate
{
    protected string $path;
    protected string $format;
    protected array $options;
    protected ?Traversable $iterator = null;

    public function __construct(string $path, string $format, array $options = [])
    {
        $this->path = $path;
        $this->format = $format;
        $this->options = $options;
    }
    public static function fromFile(string $path, array $options = []): static
    {
        return new static($path, strtolower(substr($path, strrpos($path, '.') + 1)), $options);
    }

    public function getIterator(): Traversable
    {
        if (isset($this->iterator)) {
            return $this->iterator;
        }
        switch ($this->format) {
            case 'csv':
            case 'tsv':
            case 'dsv':
            case 'ssv':
            case 'psv':
            case 'txt':
                $iterator = new CSVIterator($this->path, $this->options);
                break;
            case 'xls':
                $iterator = new XLSIterator($this->path, $this->options['active'] ?? null);
                break;
            case 'xlsx':
                $iterator = new XLSXIterator($this->path, $this->options['active'] ?? null);
                break;
            case 'xml':
                $iterator = new XMLIterator($this->path, $this->options);
                break;
            default:
                throw new Exception('Unsupported format');
        }
        return $this->iterator = $iterator;
    }
    public function toArray(): array
    {
        return iterator_to_array($this);
    }
}
