<?php

namespace vakata\spreadsheet\writer;

interface DriverInterface
{
    public function addSheet(string $name): DriverInterface;
    public function addHeaderRow(array $data): DriverInterface;
    public function addRow(array $data): DriverInterface;
    public function close(): void;
}
