<?php

namespace vakata\spreadsheet;

class CSVIterator implements \Iterator
{
    protected $stream;
    protected $options = [];
    protected $row;
    protected $ind = -1;

    public function __construct($file, array $options = [])
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
    public function current()
    {
        return $this->row;
    }
    public function key()
    {
        return $this->ind;
    }
    public function next()
    {
        $this->row = fgetcsv($this->stream, 0, $this->options['delimiter'] ?? ',', $this->options['enclosure'] ?? '"', $this->options['escape'] ?? '\\');
        $this->ind++;
    }
    public function rewind()
    {
        rewind($this->stream);
        $this->row = null;
        $this->ind = -1;
        $this->next();
    }
    public function valid()
    {
        return $this->row !== false;
    }
}
