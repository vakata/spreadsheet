<?php
namespace vakata\spreadsheet;

class OLEHelper
{
    protected $data = '';
    protected $excl = '';
    protected $bigBlocks = [];
    protected $smallBlocks = [];

    public static function getInt4d($data, $pos)
    {
        $value = ord($data[$pos]) | (ord($data[$pos+1])	<< 8) | (ord($data[$pos+2]) << 16) | (ord($data[$pos+3]) << 24);
        if ($value >= 4294967294) {
            $value = -2;
        }
        return $value;
    }

    public function __construct($sFileName)
    {
        $this->data = @file_get_contents($sFileName);
        if (!$this->data) {
            throw new Exception('Could not read data');
        }
        if (substr($this->data, 0, 8) != pack("CCCCCCCC",0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1)) {
            throw new Exception('Invalid file');
        }
        $numBigBlockDepotBlocks = self::getInt4d($this->data, 0x2c);
        $sbdStartBlock          = self::getInt4d($this->data, 0x3c);
        $rootStartBlock         = self::getInt4d($this->data, 0x30);
        $extensionBlock         = self::getInt4d($this->data, 0x44);
        $numExtensionBlocks     = self::getInt4d($this->data, 0x48);

        $bigBlockDepotBlocks = array();
        $pos = 0x4c;
        $bbdBlocks = $numBigBlockDepotBlocks;
        if ($numExtensionBlocks != 0) {
            $bbdBlocks = (0x200 - 0x4c)/4;
        }

        for ($i = 0; $i < $bbdBlocks; $i++) {
            $bigBlockDepotBlocks[$i] = self::getInt4d($this->data, $pos);
            $pos += 4;
        }

        for ($j = 0; $j < $numExtensionBlocks; $j++) {
            $pos = ($extensionBlock + 1) * 0x200;
            $blocksToRead = min($numBigBlockDepotBlocks - $bbdBlocks, 0x200 / 4 - 1);

            for ($i = $bbdBlocks; $i < $bbdBlocks + $blocksToRead; $i++) {
                $bigBlockDepotBlocks[$i] = self::getInt4d($this->data, $pos);
                $pos += 4;
            }

            $bbdBlocks += $blocksToRead;
            if ($bbdBlocks < $numBigBlockDepotBlocks) {
                $extensionBlock = self::getInt4d($this->data, $pos);
            }
        }

        // readBigBlockDepot
        for ($i = 0; $i < $numBigBlockDepotBlocks; $i++) {
            $pos = ($bigBlockDepotBlocks[$i] + 1) * 0x200;
            for ($j = 0 ; $j < 0x200 / 4; $j++) {
                $this->bigBlocks[] = self::getInt4d($this->data, $pos);
                $pos += 4;
            }
        }

        $sbdBlock = $sbdStartBlock;
        while ($sbdBlock != -2) {
            $pos = ($sbdBlock + 1) * 0x200;
            for ($j = 0; $j < 0x200 / 4; $j++) {
                $this->smallBlocks[] = self::getInt4d($this->data, $pos);
                $pos += 4;
            }
            $sbdBlock = $this->bigBlocks[$sbdBlock];
        }
        
        $properties = $this->readData($rootStartBlock);
        $root = null;
        $book = null;
        $pos = 0;
        while ($pos < strlen($properties)) {
            $temp = substr($properties, $pos, 0x80);
            $lngt = ord($temp[0x40]) | (ord($temp[0x40+1]) << 8);
            $type = ord($temp[0x42]);
            $strt = self::getInt4d($temp, 0x74);
            $size = self::getInt4d($temp, 0x78);
            $name = '';
            for ($i = 0; $i < $lngt ; $i++) {
                $name .= $temp[$i];
            }
            $name = str_replace("\x00", "", $name);
            if (strtolower($name) == "workbook" || strtolower($name) == "book") {
                $book = [
                    'name' => $name,
                    'type' => $type,
                    'strt' => $strt,
                    'size' => $size
                ];
            }
            if ($name == "Root Entry") {
                $root = [
                    'name' => $name,
                    'type' => $type,
                    'strt' => $strt,
                    'size' => $size
                ];
            }
            $pos += 0x80;
        }
        if ($book['size'] < 0x1000){
            $rootdata = $this->readData($root['strt']);
            $block = $book['strt'];
            while ($block != -2) {
                $pos = $block * 0x40;
                $this->excl .= substr($rootdata, $pos, 0x40);
                $block = $this->smallBlocks[$block];
            }
        } else {
            $numBlocks = $book['size'] / 0x200;
            if ($book['size'] % 0x200 != 0) {
                $numBlocks++;
            }
            if ($numBlocks > 0) {
                $block = $book['strt'];
                while ($block != -2) {
                    $pos = ($block + 1) * 0x200;
                    $this->excl .= substr($this->data, $pos, 0x200);
                    $block = $this->bigBlocks[$block];
                }
            }
        }
    }

    protected function readData($block)
    {
        $data = '';
        while ($block != -2)  {
            $pos   = ($block + 1) * 0x200;
            $data  = $data . substr($this->data, $pos, 0x200);
            $block = $this->bigBlocks[$block];
        }
        return $data;
    }

    public function workbook()
    {
        return $this->excl;
    }
}
