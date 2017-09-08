<?php

namespace vakata\spreadsheet;

class Reader implements \IteratorAggregate
{
    protected $file;
    protected $format;
    protected $active;
    protected $options;
    protected $iterator;

    public static function fromFile(string $file, string $format = null, $active = null, array $options = [])
    {
        $instance          = new static();
        $instance->file    = $file;
        $instance->format  = $format ?? strtolower(substr($file, strrpos($file, '.') + 1));
        $instance->active  = $active;
        $instance->options = $options;
        return $instance;
    }

    public function __destruct()
    {
        @fclose($this->stream);
    }

    public function getIterator()
    {
        if (isset($this->iterator)) {
            return $this->iterator;
        }
        switch ($this->format) {
            case 'csv':
                $iterator = new CSVIterator($this->file, $this->options);
                break;
            case 'xls':
                $iterator = new XLSIterator($this->file, $this->active);
                break;
            case 'xlsx':
                $iterator = new XLSXIterator($this->file, $this->active);
                break;
            default:
                throw new \Exception('Unsupported format');
        }
        return $this->iterator = $iterator;
    }
}
