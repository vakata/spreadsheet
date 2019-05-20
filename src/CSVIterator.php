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
        if (!$this->options['delimiter']) {
            $delimiters = [];
            foreach ([',',';',"\t",'|'] as $delimiter) {
                $delimiters[$delimiter] = [
                    fgetcsv($this->stream, 0, $delimiter),
                    fgetcsv($this->stream, 0, $delimiter)
                ];
                fseek($this->stream, 0);
            }
            foreach ($delimiters as $delimiter => $rows) {
                if (count($rows[1]) !== 0 && count($rows[0]) !== count($rows[1])) {
                    unset($delimiters[$delimiter]);
                    break;
                }
                if (!count($rows[0])) {
                    unset($delimiters[$delimiter]);
                    break;
                }
                $delimiters[$delimiter] = count($rows[0]);
            }
            if (count($delimiters)) {
                $this->options['delimiter'] = array_search(max($delimiters), $delimiters);
            }
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
