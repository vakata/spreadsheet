<?php

namespace vakata\spreadsheet;

class XLSIterator implements \Iterator
{
    protected $sheet = [];

    public function __construct(string $file, $sheet = null)
    {
        $this->sheet = (new XLSHelper($file))->getSheet();
    }

    public function current()
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
    public function key()
    {
        return key($this->sheet['cells']);
    }
    public function next()
    {
        next($this->sheet['cells']);
    }
    public function rewind()
    {
        reset($this->sheet['cells']);
    }
    public function valid()
    {
        return current($this->sheet['cells']) !== false;
    }
}
