<?php

namespace vakata\spreadsheet\reader;

use Traversable;

class XMLIterator implements \IteratorAggregate
{
    protected $xml;
    protected $items;
    protected $options = [];

    public function __construct($file, array $options = [])
    {
        $this->xml = simplexml_load_file($file, 'SimpleXMLElement', LIBXML_NOCDATA);
        $this->options = $options;
        foreach($this->xml->getDocNamespaces() as $prefix => $name) {
            if (strlen($prefix) === 0) {
                $prefix = 'ns';
            }
            $this->xml->registerXPathNamespace($prefix, $name);
        }
        foreach ($options['namespaces'] ?? [] as $k => $v) {
            $this->xml->registerXPathNamespace($k, $v);
        }
        $this->items = $this->xml->xpath($options['selector'] ?? '//item');
        foreach ($this->items as $k => $v) {
            $this->items[$k] = $this->process($v);
        }
    }
    public function getIterator(): Traversable
    {
        return new \ArrayIterator($this->items);
    }
    protected function process($xml)
    {
        return json_decode(json_encode($xml), true);
    }
}
