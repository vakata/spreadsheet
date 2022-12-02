<?php

namespace vakata\spreadsheet\reader;

use vakata\spreadsheet\Exception;
use Traversable;

class XMLIterator implements \IteratorAggregate
{
    protected \SimpleXMLElement $xml;
    protected array $items;
    protected array $options = [];

    public function __construct(string $file, array $options = [])
    {
        $this->xml = simplexml_load_file(
            $file,
            'SimpleXMLElement',
            LIBXML_NOCDATA
        ) ?: throw new Exception('Could not load');
        $this->options = $options;
        foreach ($this->xml->getDocNamespaces() ?: [] as $prefix => $name) {
            if (strlen($prefix) === 0) {
                $prefix = 'ns';
            }
            $this->xml->registerXPathNamespace($prefix, $name);
        }
        foreach ($options['namespaces'] ?? [] as $k => $v) {
            $this->xml->registerXPathNamespace($k, $v);
        }
        $this->items = ($this->xml->xpath($options['selector'] ?? '//item') ?? false) ?: [];
        foreach ($this->items as $k => $v) {
            $this->items[$k] = $this->process($v);
        }
    }
    public function getIterator(): Traversable
    {
        return new \ArrayIterator($this->items);
    }
    protected function process(mixed $xml): array
    {
        return json_decode(json_encode($xml) ?: 'NULL', true);
    }
}
