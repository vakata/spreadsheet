<?php

namespace vakata\spreadsheet\writer;

use vakata\spreadsheet\Exception;

class XMLWriter implements DriverInterface
{
    /**
     * @var resource
     */
    protected mixed $stream;
    /**
     * @var array<string,mixed>
     */
    protected array $options;
    protected bool $hasSheet = false;

    /**
     * @param mixed $stream 
     * @param array<string,mixed> $options 
     * @return void 
     */
    public function __construct(mixed $stream, array $options = [])
    {
        $this->stream = $stream;
        $this->options = array_merge(
            [
                'namespaces' => [],
                'sheet' => null,
                'root' => 'root',
                'row' => 'item',
                'cell' => 'cell'
            ],
            $options
        );
        fwrite($this->stream, '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>');
        fwrite($this->stream, '<' . $this->options['root']);
        foreach ($this->options['namespaces'] as $k => $v) {
            fwrite($this->stream, ' ' . $k . '=' . htmlspecialchars($v));
        }
        fwrite($this->stream, '>');
    }
    public function addSheet(string $name): DriverInterface
    {
        if (!$this->options['sheet']) {
            throw new Exception('Operation not supported');
        }
        if ($this->hasSheet) {
            fwrite($this->stream, '</' . $this->options['sheet'] . '>');
        }
        fwrite(
            $this->stream,
            '<' . $this->options['sheet'] . ' name="'.htmlspecialchars($name, ENT_XML1 | ENT_COMPAT, 'UTF-8').'">'
        );
        $this->hasSheet = true;
        return $this;
    }
    public function addRow(array $data): DriverInterface
    {
        fwrite($this->stream, '<' . $this->options['row'] . '>');
        foreach ($data as $k => $v) {
            $tag = preg_match('(^[a-zA-Z_]([a-zA-Z0-9_:.])*$)', $k) ? $k : $this->options['cell'];
            fwrite($this->stream, '<' . $tag . '>');
            fwrite($this->stream, htmlspecialchars($v, ENT_XML1, 'UTF-8'));
            fwrite($this->stream, '</' . $tag . '>');
        }
        fwrite($this->stream, '</' . $this->options['row'] . '>');
        return $this;
    }
    public function close(): void
    {
        if ($this->hasSheet) {
            fwrite($this->stream, '</' . $this->options['sheet'] . '>');
        }
        fwrite($this->stream, '</' . $this->options['root'] . '>');
        fclose($this->stream);
    }
}