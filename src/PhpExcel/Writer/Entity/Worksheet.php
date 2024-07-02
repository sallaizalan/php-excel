<?php

namespace Threshold\PhpExcel\Writer\Entity;

use Threshold\PhpExcel\Writer\Entity\Style\Style;
use Threshold\PhpExcel\Writer\Helper\ExtendedSimpleXMLElement;

class Worksheet
{
    const SHEET_XML_FILE_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></worksheet>';
    
    private string $filePath;
    private Sheet $externalSheet;
    private int $maxNumColumns;
    private int $lastWrittenRowIndex;
    private ?int $firstRowIndex = null;
    protected array $rows = [];
    private Style $fontStyle;
    private ?ExtendedSimpleXMLElement $xml = null;

    public function __construct(string $worksheetFilePath, Sheet $externalSheet)
    {
        $this->filePath            = $worksheetFilePath;
        $this->externalSheet       = $externalSheet;
        $this->maxNumColumns       = 0;
        $this->lastWrittenRowIndex = 0;
    }

    public function getFilePath(): string
    {
        return $this->filePath;
    }

    public function getExternalSheet(): Sheet
    {
        return $this->externalSheet;
    }

    public function getMaxNumColumns(): int
    {
        return $this->maxNumColumns;
    }

    public function setMaxNumColumns(int $maxNumColumns): self
    {
        $this->maxNumColumns = $maxNumColumns;
        return $this;
    }

    public function getLastWrittenRowIndex(): int
    {
        return $this->lastWrittenRowIndex;
    }

    public function setLastWrittenRowIndex(int $lastWrittenRowIndex): self
    {
        $this->lastWrittenRowIndex = $lastWrittenRowIndex;
        return $this;
    }
    
    public function getFirstRowIndex(): ?int
    {
        return $this->firstRowIndex;
    }
    
    public function setFirstRowIndex(?int $firstRowIndex): void
    {
        $this->firstRowIndex = $firstRowIndex;
    }

    public function getId(): int
    {
        // sheet index is zero-based, while ID is 1-based
        return $this->externalSheet->getIndex() + 1;
    }
    
    public function getRowAtIndex(int $rowIndex): ?Row
    {
        return $this->rows[$rowIndex] ?? null;
    }
    
    public function addRow(int $index, Row $row): self
    {
        $this->rows[$index] = $row;
        return $this;
    }
    
    public function getFontStyle(): Style
    {
        return $this->fontStyle;
    }
    
    public function setFontStyle(Style $fontStyle): self
    {
        $this->fontStyle = $fontStyle;
        return $this;
    }
    
    public function getXml(): ?ExtendedSimpleXMLElement
    {
        if (!$this->xml) {
            if (file_exists($this->filePath)) {
                $this->xml = simplexml_load_file($this->filePath, ExtendedSimpleXMLElement::class);
            }
            else {
                $this->xml = new ExtendedSimpleXMLElement(self::SHEET_XML_FILE_HEADER);
            }
            $this->xml->registerXPathNamespace("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        }
        
        if (!isset($this->xml->sheetData)) {
            $this->xml->addChild("sheetData");
        }
        
        return $this->xml;
    }
}