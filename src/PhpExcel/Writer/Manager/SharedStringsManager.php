<?php

namespace Threshold\PhpExcel\Writer\Manager;

use SimpleXMLElement;
use Threshold\PhpExcel\Writer\Helper\Escaper\XLSX;

class SharedStringsManager
{
    const SHARED_STRINGS_FILE_NAME = 'sharedStrings.xml';
    const SHARED_STRINGS_XML_FILE_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" count="9999999999999" uniqueCount="9999999999999"></sst>';
    private int $numSharedStrings = 0;
    private XLSX $stringsEscaper;
    private string $xlFolder;
    private SimpleXMLElement $xml;

    public function __construct(string $xlFolder, XLSX $stringsEscaper)
    {
        $this->xlFolder       = $xlFolder;
        $this->stringsEscaper = $stringsEscaper;
        $this->xml            = new SimpleXMLElement(self::SHARED_STRINGS_XML_FILE_HEADER);
    }

    public function writeString(string $string): int
    {
        $si = $this->xml->addChild("si");
        $t  = $si->addChild("t", $this->stringsEscaper->escape($string));
        $t->addAttribute("xml:space", "preserve", "http://www.w3.org/XML/1998/namespace");
        unset($si);
        unset($t);
        $this->numSharedStrings++;

        // Shared string ID is zero-based
        return ($this->numSharedStrings - 1);
    }

    public function close(): void
    {
        $this->xml->attributes()["count"]       = $this->numSharedStrings;
        $this->xml->attributes()["uniqueCount"] = $this->numSharedStrings;
        $this->xml->saveXML($this->xlFolder . '/' . self::SHARED_STRINGS_FILE_NAME);
    }
}
