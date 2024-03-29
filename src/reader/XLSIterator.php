<?php

namespace vakata\spreadsheet\reader;

class XLSIterator implements \Iterator
{
    protected array $sheet = [];

    public function __construct(string $file, mixed $desiredSheet = null)
    {
        $this->sheet = (new XLSHelper($file, $desiredSheet))->getSheet();
    }

    public function current(): mixed
    {
        $temp = current($this->sheet['cells']);
        if (is_array($temp) && isset($this->sheet['numCols'])) {
            for ($i = 0; $i < $this->sheet['numCols']; $i++) {
                if (!isset($temp[$i])) {
                    $temp[$i] = null;
                }
            }
        }
        return $temp;
    }
    public function key(): mixed
    {
        return key($this->sheet['cells']);
    }
    public function next(): void
    {
        next($this->sheet['cells']);
    }
    public function rewind(): void
    {
        reset($this->sheet['cells']);
    }
    public function valid(): bool
    {
        return current($this->sheet['cells']) !== false;
    }
}
