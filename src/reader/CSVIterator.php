<?php

namespace vakata\spreadsheet\reader;

use vakata\spreadsheet\Exception;

class CSVIterator implements \Iterator
{
    protected mixed $stream;
    protected array $options = [];
    protected mixed $row = null;
    protected int $ind = -1;

    public function __construct(string $file, array $options = [])
    {
        $this->stream = fopen($file, 'r');
        $this->options = $options;
        if ($this->stream === false) {
            throw new Exception('Document not readable');
        }
    }
    public function __destruct()
    {
        fclose($this->stream);
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
        $this->row = fgetcsv(
            $this->stream,
            0,
            $this->options['delimiter'] ?? ',',
            $this->options['enclosure'] ?? '"',
            $this->options['escape'] ?? '\\'
        );
        $this->ind++;
    }
    public function rewind(): void
    {
        rewind($this->stream);
        $this->row = null;
        $this->ind = -1;
        $this->next();
    }
    public function valid(): bool
    {
        return $this->row !== false;
    }
}
