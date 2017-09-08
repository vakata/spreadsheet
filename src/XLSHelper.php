<?php
namespace vakata\spreadsheet;

class XLSHelper
{
    const READER_BIFF8 = 0x600;
    const READER_BIFF7 = 0x500;
    const READER_WORKBOOKGLOBALS = 0x5;
    const READER_WORKSHEET = 0x10;
    const READER_TYPE_BOF = 0x809;
    const READER_TYPE_EOF = 0x0a;
    const READER_TYPE_BOUNDSHEET = 0x85;
    const READER_TYPE_DIMENSION = 0x200;
    const READER_TYPE_ROW = 0x208;
    const READER_TYPE_DBCELL = 0xd7;
    const READER_TYPE_FILEPASS = 0x2f;
    const READER_TYPE_NOTE = 0x1c;
    const READER_TYPE_TXO = 0x1b6;
    const READER_TYPE_RK = 0x7e;
    const READER_TYPE_RK2 = 0x27e;
    const READER_TYPE_MULRK = 0xbd;
    const READER_TYPE_MULBLANK = 0xbe;
    const READER_TYPE_INDEX = 0x20b;
    const READER_TYPE_SST = 0xfc;
    const READER_TYPE_EXTSST = 0xff;
    const READER_TYPE_CONTINUE = 0x3c;
    const READER_TYPE_LABEL = 0x204;
    const READER_TYPE_LABELSST = 0xfd;
    const READER_TYPE_NUMBER = 0x203;
    const READER_TYPE_NAME = 0x18;
    const READER_TYPE_ARRAY = 0x221;
    const READER_TYPE_STRING = 0x207;
    const READER_TYPE_FORMULA = 0x406;
    const READER_TYPE_FORMULA2 = 0x6;
    const READER_TYPE_FORMAT = 0x41e;
    const READER_TYPE_XF = 0xe0;
    const READER_TYPE_BOOLERR = 0x205;
    const READER_TYPE_FONT = 0x0031;
    const READER_TYPE_PALETTE = 0x0092;
    const READER_TYPE_UNKNOWN = 0xffff;
    const READER_TYPE_NINETEENFOUR = 0x22;
    const READER_TYPE_MERGEDCELLS = 0xE5;
    const READER_UTCOFFSETDAYS = 25569;
    const READER_UTCOFFSETDAYS1904 = 24107;
    const READER_MSINADAY = 86400;
    const READER_TYPE_HYPER = 0x01b8;
    const READER_TYPE_COLINFO = 0x7d;
    const READER_TYPE_DEFCOLWIDTH = 0x55;
    const READER_TYPE_STANDARDWIDTH = 0x99;
    const READER_DEF_NUM_FORMAT = "%s";

    protected $data;
    protected $sst = [];
    protected $sheet;
    protected $nineteenFour;
    protected $formatRecords = [];
    protected $xfRecords = [];

    public function __construct($file, $sheet = null)
    {
        $this->data = (new OLEHelper($file))->workbook();
        $this->parseWorkbook($sheet);
    }
    public function getSheet()
    {
        return $this->sheet;
    }

    protected function parseWorkbook($sheet = null)
    {
        $pos = 0;
        $data = $this->data;

        $length = $this->v($data,$pos+2);
        $version = $this->v($data,$pos+4);
        $substreamType = $this->v($data,$pos+6);

        if (($version != self::READER_BIFF8) &&
            ($version != self::READER_BIFF7)) {
            return false;
        }

        if ($substreamType != self::READER_WORKBOOKGLOBALS){
            return false;
        }

        $pos += $length + 4;

        $code = $this->v($data,$pos);
        $length = $this->v($data,$pos+2);

        $sheets = [];
        while ($code != self::READER_TYPE_EOF) {
            switch ($code) {
                case self::READER_TYPE_SST:
                    $spos = $pos + 4;
                    $limitpos = $spos + $length;
                    $uniqueStrings = $this->getInt4d($data, $spos+4);
                    $spos += 8;
                    for ($i = 0; $i < $uniqueStrings; $i++) {
                        $formattingRuns = 0;
                        $extendedRunLength = 0;
                        // Read in the number of characters
                        if ($spos == $limitpos) {
                            $opcode = $this->v($data,$spos);
                            $conlength = $this->v($data,$spos+2);
                            if ($opcode != 0x3c) {
                                return -1;
                            }
                            $spos += 4;
                            $limitpos = $spos + $conlength;
                        }
                        $numChars = ord($data[$spos]) | (ord($data[$spos+1]) << 8);
                        $spos += 2;
                        $optionFlags = ord($data[$spos]);
                        $spos++;
                        $asciiEncoding = (($optionFlags & 0x01) == 0) ;
                        $extendedString = ( ($optionFlags & 0x04) != 0);

                        // See if string contains formatting information
                        $richString = ( ($optionFlags & 0x08) != 0);

                        if ($richString) {
                            // Read in the crun
                            $formattingRuns = $this->v($data,$spos);
                            $spos += 2;
                        }

                        if ($extendedString) {
                            // Read in cchExtRst
                            $extendedRunLength = $this->getInt4d($data, $spos);
                            $spos += 4;
                        }

                        $len = ($asciiEncoding)? $numChars : $numChars*2;
                        if ($spos + $len < $limitpos) {
                            $retstr = substr($data, $spos, $len);
                            $spos += $len;
                        } else{
                            // found countinue
                            $retstr = substr($data, $spos, $limitpos - $spos);
                            $bytesRead = $limitpos - $spos;
                            $charsLeft = $numChars - (($asciiEncoding) ? $bytesRead : ($bytesRead / 2));
                            $spos = $limitpos;

                            while ($charsLeft > 0){
                                $opcode = $this->v($data,$spos);
                                $conlength = $this->v($data,$spos+2);
                                if ($opcode != 0x3c) {
                                    return -1;
                                }
                                $spos += 4;
                                $limitpos = $spos + $conlength;
                                $option = ord($data[$spos]);
                                $spos += 1;
                                if ($asciiEncoding && ($option == 0)) {
                                    $len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
                                    $retstr .= substr($data, $spos, $len);
                                    $charsLeft -= $len;
                                    $asciiEncoding = true;
                                } elseif (!$asciiEncoding && ($option != 0)) {
                                    $len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
                                    $retstr .= substr($data, $spos, $len);
                                    $charsLeft -= $len/2;
                                    $asciiEncoding = false;
                                } elseif (!$asciiEncoding && ($option == 0)) {
                                    // Bummer - the string starts off as Unicode, but after the
                                    // continuation it is in straightforward ASCII encoding
                                    $len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
                                    for ($j = 0; $j < $len; $j++) {
                                        $retstr .= $data[$spos + $j].chr(0);
                                    }
                                    $charsLeft -= $len;
                                    $asciiEncoding = false;
                                } else {
                                    $newstr = '';
                                    for ($j = 0; $j < strlen($retstr); $j++) {
                                        $newstr = $retstr[$j].chr(0);
                                    }
                                    $retstr = $newstr;
                                    $len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
                                    $retstr .= substr($data, $spos, $len);
                                    $charsLeft -= $len/2;
                                    $asciiEncoding = false;
                                }
                                $spos += $len;
                            }
                        }
                        $retstr = ($asciiEncoding) ? $retstr : $this->encodeUTF16($retstr);
                        if ($richString) {
                            $spos += 4 * $formattingRuns;
                        }
                        // For extended strings, skip over the extended string data
                        if ($extendedString) {
                            $spos += $extendedRunLength;
                        }
                        $this->sst[]=$retstr;
                    }
                    break;
                case self::READER_TYPE_FILEPASS:
                    throw new \Exception('Workbook incomplete');
                case self::READER_TYPE_NAME:
                    break;
                case self::READER_TYPE_FORMAT:
                    $indexCode = $this->v($data,$pos+4);
                    if ($version == self::READER_BIFF8) {
                        $numchars = $this->v($data,$pos+6);
                        if (ord($data[$pos+8]) == 0){
                            $formatString = substr($data, $pos+9, $numchars);
                        } else {
                            $formatString = substr($data, $pos+9, $numchars*2);
                        }
                    } else {
                        $numchars = ord($data[$pos+6]);
                        $formatString = substr($data, $pos+7, $numchars*2);
                    }
                    $this->formatRecords[$indexCode] = $formatString;
                    break;
                
                case self::READER_TYPE_XF:
                    $indexCode = ord($data[$pos+6]) | ord($data[$pos+7]) << 8;
                    $xf = [];
                    $xf['formatIndex'] = $indexCode;

                    $dateFormats = [
                        0xe  => "d.m.Y",
                        0xf  => "d-M-Y",
                        0x10 => "d-M",
                        0x11 => "M-Y",
                        0x12 => "h:i a",
                        0x13 => "h:i:s a",
                        0x14 => "H:i",
                        0x15 => "H:i:s",
                        0x16 => "d.m.Y H:i",
                        0x2d => "i:s",
                        0x2e => "H:i:s",
                        0x2f => "i:s.S"
                    ];
                    $numberFormats = [
                        0x1 => "0",
                        0x2 => "0.00",
                        0x3 => "#,##0",
                        0x4 => "#,##0.00",
                        0x5 => "\$#,##0;(\$#,##0)",
                        0x6 => "\$#,##0;[Red](\$#,##0)",
                        0x7 => "\$#,##0.00;(\$#,##0.00)",
                        0x8 => "\$#,##0.00;[Red](\$#,##0.00)",
                        0x9 => "0%",
                        0xa => "0.00%",
                        0xb => "0.00E+00",
                        0x25 => "#,##0;(#,##0)",
                        0x26 => "#,##0;[Red](#,##0)",
                        0x27 => "#,##0.00;(#,##0.00)",
                        0x28 => "#,##0.00;[Red](#,##0.00)",
                        0x29 => "#,##0;(#,##0)",  // Not exactly
                        0x2a => "\$#,##0;(\$#,##0)",  // Not exactly
                        0x2b => "#,##0.00;(#,##0.00)",  // Not exactly
                        0x2c => "\$#,##0.00;(\$#,##0.00)",  // Not exactly
                        0x30 => "##0.0E+0"
                    ];
                    if (array_key_exists($indexCode, $dateFormats)) {
                        $xf['type'] = 'date';
                        $xf['format'] = $dateFormats[$indexCode];
                    } elseif (array_key_exists($indexCode, $numberFormats)) {
                        $xf['type'] = 'number';
                        $xf['format'] = $numberFormats[$indexCode];
                    } else {
                        $isdate = FALSE;
                        $formatstr = '';
                        if ($indexCode > 0){
                            if (isset($this->formatRecords[$indexCode])) {
                                $formatstr = $this->formatRecords[$indexCode];
                            }
                            if ($formatstr!="") {
                                $tmp = preg_replace("/\;.*/","",$formatstr);
                                $tmp = preg_replace("/^\[[^\]]*\]/","",$tmp);
                                if (preg_match("/[^hmsday\/\-:\s\\\,AMP]/i", $tmp) == 0) { // found day and time format
                                    $isdate = TRUE;
                                    $formatstr = strtolower($tmp);
                                    $formatstr = str_replace(array('am/pm','mmmm','mmm'), array('a','F','M'), $formatstr);
                                    // m/mm are used for both minutes and months - oh SNAP!
                                    // This mess tries to fix for that.
                                    // 'm' == minutes only if following h/hh or preceding s/ss
                                    $formatstr = preg_replace("/(h:?)mm?/","$1i", $formatstr);
                                    $formatstr = preg_replace("/mm?(:?s)/","i$1", $formatstr);
                                    // A single 'm' = n in PHP
                                    $formatstr = preg_replace("/(^|[^m])m([^m]|$)/", '$1n$2', $formatstr);
                                    $formatstr = preg_replace("/(^|[^m])m([^m]|$)/", '$1n$2', $formatstr);
                                    // else it's months
                                    $formatstr = str_replace('mm', 'm', $formatstr);
                                    // Convert single 'd' to 'j'
                                    $formatstr = preg_replace("/(^|[^d])d([^d]|$)/", '$1j$2', $formatstr);
                                    $formatstr = str_replace(array('dddd','ddd','dd','yyyy','yy','hh','h'), array('l','D','d','Y','y','H','g'), $formatstr);
                                    $formatstr = preg_replace("/ss?/", 's', $formatstr);
                                }
                            }
                        }
                        if ($isdate) {
                            $xf['type'] = 'date';
                            $xf['format'] = $formatstr;
                        } else {
                            // If the format string has a 0 or # in it, we'll assume it's a number
                            if (preg_match("/[0#]/", $formatstr)) {
                                $xf['type'] = 'number';
                            } else {
                                $xf['type'] = 'other';
                            }
                            $xf['format'] = $formatstr;
                            $xf['code'] = $indexCode;
                        }
                    }
                    $this->xfRecords[] = $xf;
                    break;
                case self::READER_TYPE_NINETEENFOUR:
                    $this->nineteenFour = (ord($data[$pos+4]) == 1);
                    break;
                case self::READER_TYPE_BOUNDSHEET:
                    $rec_offset = $this->getInt4d($data, $pos+4);
                    $rec_length = ord($data[$pos+10]);

                    $rec_name = '';
                    if ($version == self::READER_BIFF8){
                        $chartype =  ord($data[$pos+11]);
                        if ($chartype == 0){
                            $rec_name = substr($data, $pos+12, $rec_length);
                        } else {
                            $rec_name = $this->encodeUTF16(substr($data, $pos+12, $rec_length*2));
                        }
                    } elseif ($version == self::READER_BIFF7){
                        $rec_name = substr($data, $pos+11, $rec_length);
                    }
                    $sheets[] = [ 'name' => $rec_name, 'offset' => $rec_offset ];
                    break;
            }

            $pos += $length + 4;
            $code = ord($data[$pos]) | ord($data[$pos+1])<<8;
            $length = ord($data[$pos+2]) | ord($data[$pos+3])<<8;
        }

        $offset = null;
        foreach ($sheets as $key => $val) {
            if (isset($sheet) && $val['name'] === $sheet) {
                $offset = $val['offset'];
                break;
            }
        }
        $sheet = $sheet ?? 0;
        if (!isset($offset) && isset($sheets[$sheet])) {
            $offset = $sheets[$sheet]['offset'];
        }
        if (!isset($offset)) {
            throw new Exception('Invalid sheet');
        }
        $this->parseSheet($offset);
    }
    protected function parseSheet($spos)
    {
        $cont = true;
        $data = $this->data;
        // read BOF
        $code = ord($data[$spos]) | ord($data[$spos+1])<<8;
        $length = ord($data[$spos+2]) | ord($data[$spos+3])<<8;

        $version = ord($data[$spos + 4]) | ord($data[$spos + 5])<<8;
        $substreamType = ord($data[$spos + 6]) | ord($data[$spos + 7])<<8;

        if (($version != self::READER_BIFF8) && ($version != self::READER_BIFF7)) {
            return -1;
        }

        if ($substreamType != self::READER_WORKSHEET){
            return -2;
        }
        $spos += $length + 4;
        $previousRow = 0;
        $previousCol = 0;
        while ($cont) {
            $lowcode = ord($data[$spos]);
            if ($lowcode == self::READER_TYPE_EOF) break;
            $code = $lowcode | ord($data[$spos+1])<<8;
            $length = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
            $spos += 4;
            switch ($code) {
                case self::READER_TYPE_DIMENSION:
                    if (!isset($this->numRows)) {
                        if (($length == 10) ||  ($version == self::READER_BIFF7)){
                            $this->sheet['numRows'] = ord($data[$spos+2]) | ord($data[$spos+3]) << 8;
                            $this->sheet['numCols'] = ord($data[$spos+6]) | ord($data[$spos+7]) << 8;
                        } else {
                            $this->sheet['numRows'] = ord($data[$spos+4]) | ord($data[$spos+5]) << 8;
                            $this->sheet['numCols'] = ord($data[$spos+10]) | ord($data[$spos+11]) << 8;
                        }
                    }
                    break;
                case self::READER_TYPE_MERGEDCELLS:
                    break;
                case self::READER_TYPE_RK:
                case self::READER_TYPE_RK2:
                    $row = ord($data[$spos]) | ord($data[$spos+1])<<8;
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $rknum = $this->getInt4d($data, $spos + 6);
                    $numValue = $this->getIEEE754($rknum);
                    $info = $this->getCellDetails($spos,$numValue);
                    $this->addCell($row, $column, $info['string']);
                    break;
                case self::READER_TYPE_LABELSST:
                    $row     = ord($data[$spos]) | ord($data[$spos+1])<<8;
                    $column  = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $index   = $this->getInt4d($data, $spos + 6);
                    $this->addCell($row, $column, $this->sst[$index]);
                    break;
                case self::READER_TYPE_MULRK:
                    $row      = ord($data[$spos]) | ord($data[$spos+1])<<8;
                    $colFirst = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $colLast  = ord($data[$spos + $length - 2]) | ord($data[$spos + $length - 1])<<8;
                    $columns  = $colLast - $colFirst + 1;
                    $tmppos   = $spos+4;
                    for ($i = 0; $i < $columns; $i++) {
                        $numValue = $this->getIEEE754($this->getInt4d($data, $tmppos + 2));
                        $info = $this->getCellDetails($tmppos-4,$numValue);
                        $tmppos += 6;
                        $this->addCell($row, $colFirst + $i, $info['string']);
                    }
                    break;
                case self::READER_TYPE_NUMBER:
                    $row	= ord($data[$spos]) | ord($data[$spos+1])<<8;
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $tmp = unpack("ddouble", substr($data, $spos + 6, 8)); // It machine machine dependent
                    if ($this->isDate($spos)) {
                        $numValue = $tmp['double'];
                    }
                    else {
                        $numValue = $this->createNumber($spos);
                    }
                    $info = $this->getCellDetails($spos,$numValue);
                    $this->addCell($row, $column, $info['string']);
                    break;
                case self::READER_TYPE_FORMULA:
                case self::READER_TYPE_FORMULA2:
                    $row	= ord($data[$spos]) | ord($data[$spos+1])<<8;
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    if ((ord($data[$spos+6])==0) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
                        //String formula. Result follows in a STRING record
                        // This row/col are stored to be referenced in that record
                        // http://code.google.com/p/php-excel-reader/issues/detail?id=4
                        $previousRow = $row;
                        $previousCol = $column;
                    } elseif ((ord($data[$spos+6])==1) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
                        //Boolean formula. Result is in +2; 0=false,1=true
                        // http://code.google.com/p/php-excel-reader/issues/detail?id=4
                        if (ord($this->data[$spos+8])==1) {
                            $this->addCell($row, $column, "TRUE");
                        } else {
                            $this->addCell($row, $column, "FALSE");
                        }
                    } elseif ((ord($data[$spos+6])==2) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
                        //Error formula. Error code is in +2;
                    } elseif ((ord($data[$spos+6])==3) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
                        //Formula result is a null string.
                        $this->addCell($row, $column, '');
                    } else {
                        // result is a number, so first 14 bytes are just like a _NUMBER record
                        $tmp = unpack("ddouble", substr($data, $spos + 6, 8)); // It machine machine dependent
                              if ($this->isDate($spos)) {
                                $numValue = $tmp['double'];
                              }
                              else {
                                $numValue = $this->createNumber($spos);
                              }
                        $info = $this->getCellDetails($spos,$numValue);
                        $this->addCell($row, $column, $info['string']);
                    }
                    break;
                case self::READER_TYPE_BOOLERR:
                    $row	= ord($data[$spos]) | ord($data[$spos+1])<<8;
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $string = ord($data[$spos+6]);
                    $this->addCell($row, $column, $string);
                    break;
                case self::READER_TYPE_STRING:
                    $retstr = '';
                    // http://code.google.com/p/php-excel-reader/issues/detail?id=4
                    if ($version == self::READER_BIFF8) {
                        // Unicode 16 string, like an SST record
                        $xpos = $spos;
                        $numChars =ord($data[$xpos]) | (ord($data[$xpos+1]) << 8);
                        $xpos += 2;
                        $optionFlags =ord($data[$xpos]);
                        $xpos++;
                        $asciiEncoding = (($optionFlags &0x01) == 0) ;
                        $extendedString = (($optionFlags & 0x04) != 0);
                        // See if string contains formatting information
                        $richString = (($optionFlags & 0x08) != 0);
                        if ($richString) {
                            // Read in the crun
                            $formattingRuns = ord($data[$xpos]) | (ord($data[$xpos+1]) << 8);
                            $xpos += 2;
                        }
                        if ($extendedString) {
                            // Read in cchExtRst
                            $extendedRunLength =$this->getInt4d($this->data, $xpos);
                            $xpos += 4;
                        }
                        $len = ($asciiEncoding)?$numChars : $numChars*2;
                        $retstr =substr($data, $xpos, $len);
                        $xpos += $len;
                        $retstr = ($asciiEncoding)? $retstr : $this->encodeUTF16($retstr);
                    } elseif ($version == self::READER_BIFF7){
                        // Simple byte string
                        $xpos = $spos;
                        $numChars =ord($data[$xpos]) | (ord($data[$xpos+1]) << 8);
                        $xpos += 2;
                        $retstr = substr($data, $xpos, $numChars);
                    }
                    $this->addCell($previousRow, $previousCol, $retstr);
                    break;
                case self::READER_TYPE_ROW:
                    break;
                case self::READER_TYPE_DBCELL:
                    break;
                case self::READER_TYPE_MULBLANK:
                    $row = ord($data[$spos]) | ord($data[$spos+1])<<8;
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $cols = ($length / 2) - 3;
                    for ($c = 0; $c < $cols; $c++) {
                        $this->addCell($row, $column + $c, "");
                    }
                    break;
                case self::READER_TYPE_LABEL:
                    $row	= ord($data[$spos]) | ord($data[$spos+1])<<8;
                    $column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
                    $this->addCell($row, $column, substr($data, $spos + 8, ord($data[$spos + 6]) | ord($data[$spos + 7])<<8));
                    break;
                case self::READER_TYPE_EOF:
                    $cont = false;
                    break;
                case self::READER_TYPE_HYPER:
                    //  Only handle hyperlinks to a URL
                    $row	= ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
                    $row2   = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
                    $column = ord($this->data[$spos+4]) | ord($this->data[$spos+5])<<8;
                    $column2 = ord($this->data[$spos+6]) | ord($this->data[$spos+7])<<8;
                    $linkdata = [];
                    $flags = ord($this->data[$spos + 28]);
                    $udesc = "";
                    $ulink = "";
                    $uloc = 32;
                    $linkdata['flags'] = $flags;
                    if (($flags & 1) > 0 ) {   // is a type we understand
                        //  is there a description ?
                        if (($flags & 0x14) == 0x14 ) {   // has a description
                            $uloc += 4;
                            $descLen = ord($this->data[$spos + 32]) | ord($this->data[$spos + 33]) << 8;
                            $udesc = substr($this->data, $spos + $uloc, $descLen * 2);
                            $uloc += 2 * $descLen;
                        }
                        $ulink = $this->read16bitstring($this->data, $spos + $uloc + 20);
                        if ($udesc == "") {
                            $udesc = $ulink;
                        }
                    }
                    $linkdata['desc'] = $udesc;
                    $linkdata['link'] = $this->encodeUTF16($ulink);
                    for ($r=$row; $r<=$row2; $r++) { 
                        for ($c=$column; $c<=$column2; $c++) {
                            $this->sheet['cellsInfo'][$r+1][$c+1]['hyperlink'] = $linkdata;
                        }
                    }
                    break;
                case self::READER_TYPE_COLINFO:
                    break;
                default:
                    break;
            }
            $spos += $length;
        }
    }
    protected function getCellDetails($spos, $numValue)
    {
        $xfindex = ord($this->data[$spos+4]) | ord($this->data[$spos+5]) << 8;
        $xfrecord = $this->xfRecords[$xfindex];
        $type = $xfrecord['type'];

        $format = $xfrecord['format'];
        $formatIndex = $xfrecord['formatIndex'];
        $rectype = '';
        $string = '';
        $raw = '';

        if ($type == 'date') {
            // See http://groups.google.com/group/php-excel-reader-discuss/browse_frm/thread/9c3f9790d12d8e10/f2045c2369ac79de
            $rectype = 'date';
            // Convert numeric value into a date
            $utcDays = floor($numValue - ($this->nineteenFour ? self::READER_UTCOFFSETDAYS1904 : self::READER_UTCOFFSETDAYS));
            $utcValue = ($utcDays) * self::READER_MSINADAY;
            $dateinfo = array_combine(
                ['seconds','minutes','hours','mday','wday','mon','year','yday','weekday','month',0],
                explode(":", gmdate('s:i:G:j:w:n:Y:z:l:F:U', $utcValue ?? time()))
            );

            $raw = $numValue;
            $fractionalDay = $numValue - floor($numValue) + .0000001; // The .0000001 is to fix for php/excel fractional diffs

            $totalseconds = floor(self::READER_MSINADAY * $fractionalDay);
            $secs = $totalseconds % 60;
            $totalseconds -= $secs;
            $hours = floor($totalseconds / (60 * 60));
            $mins = floor($totalseconds / 60) % 60;
            $string = date ($format, mktime($hours, $mins, $secs, $dateinfo["mon"], $dateinfo["mday"], $dateinfo["year"]));
        } else if ($type == 'number') {
            $rectype = 'number';
            $string = $numValue;
            $raw = $numValue;
        } else {
            if ($format=="") {
                $format = self::READER_DEF_NUM_FORMAT;
            }
            $rectype = 'unknown';
            $string = $numValue;
            $raw = $numValue;
        }

        return array(
            'string'      => $string,
            'raw'         => $raw,
            'rectype'     => $rectype,
            'format'      => $format,
            'formatIndex' => $formatIndex,
            'xfIndex'     => $xfindex
        );
    }
    protected function addCell($row, $col, $string)
    {
        $this->sheet['cells'][$row][$col] = $string;
    }

    protected function isDate($spos)
    {
        $xfindex = ord($this->data[$spos+4]) | ord($this->data[$spos+5]) << 8;
        return ($this->xfRecords[$xfindex]['type'] == 'date');
    }
    protected function createNumber($spos)
    {
        $rknumhigh = $this->getInt4d($this->data, $spos + 10);
        $rknumlow = $this->getInt4d($this->data, $spos + 6);
        $sign = ($rknumhigh & 0x80000000) >> 31;
        $exp =  ($rknumhigh & 0x7ff00000) >> 20;
        $mantissa = (0x100000 | ($rknumhigh & 0x000fffff));
        $mantissalow1 = ($rknumlow & 0x80000000) >> 31;
        $mantissalow2 = ($rknumlow & 0x7fffffff);
        $value = $mantissa / pow( 2 , (20- ($exp - 1023)));
        if ($mantissalow1 != 0) {
            $value += 1 / pow (2 , (21 - ($exp - 1023)));
        }
        $value += $mantissalow2 / pow (2 , (52 - ($exp - 1023)));
        if ($sign) {
            $value = -1 * $value;
        }
        return  $value;
    }
    protected function getIEEE754($rknum)
    {
        if (($rknum & 0x02) != 0) {
            $value = $rknum >> 2;
        } else {
            //mmp
            // I got my info on IEEE754 encoding from
            // http://research.microsoft.com/~hollasch/cgindex/coding/ieeefloat.html
            // The RK format calls for using only the most significant 30 bits of the
            // 64 bit floating point value. The other 34 bits are assumed to be 0
            // So, we use the upper 30 bits of $rknum as follows...
            $sign = ($rknum & 0x80000000) >> 31;
            $exp = ($rknum & 0x7ff00000) >> 20;
            $mantissa = (0x100000 | ($rknum & 0x000ffffc));
            $value = $mantissa / pow( 2 , (20- ($exp - 1023)));
            if ($sign) {
                $value = -1 * $value;
            }
            //end of changes by mmp
        }
        if (($rknum & 0x01) != 0) {
            $value /= 100;
        }
        return $value;
    }
    protected function encodeUTF16($string)
    {
        if (function_exists('iconv')) {
            return iconv('UTF-16LE', 'UTF-8', $string);
        }
        if (function_exists('mb_convert_encoding')) {
            return mb_convert_encoding($string, 'UTF-8', 'UTF-16LE');
        }
        return $string;
    }
    protected function getInt4d($data, $pos)
    {
        $_or_24 = ord($data[$pos+3]);
        if ($_or_24>=128) {
            $_ord_24 = -abs((256-$_or_24) << 24);
        } else {
            $_ord_24 = ($_or_24&127) << 24;
        }
        return ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | $_ord_24;
    }
    protected function read16bitstring($data, $start) {
        $len = 0;
        while (ord($data[$start + $len]) + ord($data[$start + $len + 1]) > 0) {
            $len++;
        }
        return substr($data, $start, $len);
    }
    protected function v($data, $pos)
    {
        return ord($data[$pos]) | ord($data[$pos+1])<<8;
    }
}