<?php

namespace vakata\spreadsheet\reader;

use vakata\spreadsheet\Exception;

class CSVIterator implements \Iterator
{
    protected mixed $stream;
    protected array $options = [];
    protected mixed $row = null;
    protected int $ind = -1;
    protected string $bom = "";

    public function __construct(string $file, array $options = [])
    {
        $this->stream = fopen($file, 'r');
        $this->options = $options;
        if ($this->stream === false) {
            throw new Exception('Document not readable');
        }
        // detect and skip bom
        $tmp = fread($this->stream, 4);
        foreach ([ "\xEF\xBB\xBF", "\xFF\xFE", "\xFE\xFF", "\xFF\xFE\x00\x00", "\x00\x00\xFE\xFF" ] as $b) {
            if (strncmp($tmp, $b, strlen($b)) === 0) {
                $this->bom = $b;
                break;
            }
        }
        rewind($this->stream);
        if ($this->bom) {
            fread($this->stream, strlen($this->bom));
        }
        // detect delimiter
        if (!isset($this->options['delimiter'])) {
            $pos = ftell($this->stream);
            $del = null;
            foreach ([",", ";", "\t", "|", ":"] as $delimiter) {
                fseek($this->stream, $pos);
                $cols = [];
                for ($i = 0; $i < 15; $i++) {
                    $tmp = fgetcsv(
                        $this->stream,
                        0,
                        $delimiter,
                        $this->options['enclosure'] ?? '"',
                        $this->options['escape'] ?? '\\'
                    );
                    if ($tmp && is_array($tmp) && count($tmp) > 1) {
                        $cols[] = count($tmp);
                    }
                }
                $cols = array_unique($cols);
                if (count($cols) && (!isset($del) || $del > count($cols))) {
                    $del = count($cols);
                    $this->options['delimiter'] = $delimiter;
                }
            }
            fseek($this->stream, $pos);
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
        if ($this->row !== false && is_array($this->row)) {
            $this->row[count($this->row) - 1] = rtrim($this->row[count($this->row) - 1], "\r");
        }
        $this->ind++;
    }
    public function rewind(): void
    {
        rewind($this->stream);
        $this->row = null;
        $this->ind = -1;
        if ($this->bom) {
            fread($this->stream, strlen($this->bom));
        }
        $this->next();
    }
    public function valid(): bool
    {
        return $this->row !== false;
    }
}
