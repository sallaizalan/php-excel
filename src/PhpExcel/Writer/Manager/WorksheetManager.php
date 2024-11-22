<?php

namespace Threshold\PhpExcel\Writer\Manager;

use Exception;
use Threshold\PhpExcel\Writer\Entity\{Cell, Options, Row, Style\Style, Worksheet};
use Threshold\PhpExcel\Writer\Exception\{InvalidArgumentException};
use Threshold\PhpExcel\Writer\Helper\{CellHelper,
    ColumnHelper,
    Escaper\XLSX,
    ExtendedSimpleXMLElement,
    StringHelper,
    WorksheetHelper};
use Threshold\PhpExcel\Writer\Manager\Style\{StyleManager, StyleMerger};

class WorksheetManager
{
    /*
     * Maximum number of characters a cell can contain
     * @see https://support.office.com/en-us/article/Excel-specifications-and-limits-16c69c74-3d6a-4aaf-ba35-e6eb276e8eaa [Excel 2007]
     * @see https://support.office.com/en-us/article/Excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3 [Excel 2010]
     * @see https://support.office.com/en-us/article/Excel-specifications-and-limits-ca36e2dc-1f09-4620-b726-67c00b05040f [Excel 2013/2016]
     */
    const MAX_CHARACTERS_PER_CELL = 32767;

    private bool $shouldUseInlineStrings;
    private RowManager $rowManager;
    private StyleManager $styleManager;
    private StyleMerger $styleMerger;
    private SharedStringsManager $sharedStringsManager;
    private XLSX $stringsEscaper;
    private StringHelper $stringHelper;

    public function __construct(OptionsManager $optionsManager, RowManager $rowManager, StyleManager $styleManager,
                                StyleMerger $styleMerger, SharedStringsManager $sharedStringsManager, XLSX $stringsEscaper,
                                StringHelper $stringHelper)
    {
        $this->shouldUseInlineStrings = $optionsManager->getOption(Options::SHOULD_USE_INLINE_STRINGS);
        $this->rowManager             = $rowManager;
        $this->styleManager           = $styleManager;
        $this->styleMerger            = $styleMerger;
        $this->sharedStringsManager   = $sharedStringsManager;
        $this->stringsEscaper         = $stringsEscaper;
        $this->stringHelper           = $stringHelper;
    }

    public function getSharedStringsManager(): SharedStringsManager
    {
        return $this->sharedStringsManager;
    }
    
    /**
     * @throws Exception
     */
    public function startSheet(Worksheet $worksheet): void
    {
        $rowIndexes = [];
        foreach ($worksheet->getXML()->xpath("xmlns:sheetData/xmlns:row") as $xmlRowElement) {
            if (isset($xmlRowElement["r"])) {
                $rowIndexes[] = (int)$xmlRowElement["r"];
            }
        }
        if (!empty($rowIndexes)) {
            $worksheet->setLastWrittenRowIndex(max($rowIndexes));
        }
        $worksheet->setFirstRowIndex($worksheet->getLastWrittenRowIndex());
    }
    
    /**
     * @throws InvalidArgumentException
     */
    public function addRow(Worksheet $worksheet, Row $row, ?int $rowIndex = null): void
    {
        if (!$this->rowManager->isEmpty($row)) {
            $this->addNonEmptyRow($worksheet, $row, $rowIndex);
        }

        $worksheet->setLastWrittenRowIndex($worksheet->getLastWrittenRowIndex() + 1);
    }

    /**
     * @throws InvalidArgumentException
     */
    private function addNonEmptyRow(Worksheet $worksheet, Row $row, ?int $rowIndex = null): void
    {
        $rowStyle         = $row->getStyle();
        $rowIndexOneBased = $rowIndex ?? $worksheet->getLastWrittenRowIndex() + 1;
        $numCells         = $row->getNumCells();
        
        if ($rowIndex) {
            $rowXML = null;
            
            if (!empty($xmlRow = $worksheet->getXML()->xpath('xmlns:sheetData/xmlns:row[@r="' . $rowIndex . '"]'))) {
                $dom = dom_import_simplexml($xmlRow[0]);
                $dom->parentNode->removeChild($dom);
            }
            
            foreach ($worksheet->getXML()->xpath("xmlns:sheetData/xmlns:row") as $xmlIndex => $xmlRowElement) {
                if (isset($xmlRowElement["r"]) && (int)$xmlRowElement["r"] > $rowIndex) {
                    $rowXML = $worksheet->getXML()->sheetData->insertChild("row", "", $xmlIndex);
                    break;
                }
            }
            
            if ($rowXML === null) {
                $rowXML = $worksheet->getXML()->sheetData->addChild("row");
            }
        }
        else {
            $rowXML = $worksheet->getXML()->sheetData->addChild("row");
        }
        
        $rowXML->addAttribute("r", $rowIndexOneBased);
        $rowXML->addAttribute("spans", ("1:" . $numCells));
        
        if ($rowIndex && $rowIndex > $worksheet->getLastWrittenRowIndex()) {
            $worksheet->setLastWrittenRowIndex($rowIndex - 1);
        }
        
        foreach ($row->getCells() as $columnIndexZeroBased => $cell) {
            $worksheet->addRow($rowIndexOneBased, $row);
            $registeredStyle = $this->applyStyleAndRegister($cell, $rowStyle);
            $cellStyle       = $registeredStyle->getStyle();
            
            if ($registeredStyle->isMatchingRowStyle()) {
                $rowStyle = $cellStyle; // Replace actual rowStyle (possibly with null id) by registered style (with id)
            }
            
            $this->addCellToRowXML($rowXML, $rowIndexOneBased, $columnIndexZeroBased, $cell, $cellStyle->getId());
        }
    }
    
    /**
     * @throws InvalidArgumentException
     */
    private function addCellToRowXML(ExtendedSimpleXMLElement $rowXML, int $rowIndexOneBased, int $columnIndexZeroBased, Cell $cell, int $styleId): void
    {
        $columnLetters = CellHelper::getColumnLettersFromColumnIndex($columnIndexZeroBased);
        $cellXML       = $rowXML->addChild("c");
        $cellXML->addAttribute("r", $columnLetters . $rowIndexOneBased);
        $cellXML->addAttribute("s", $styleId);
        
        if ($cell->isString()) {
            if ($this->stringHelper->getStringLength($cell->getValue()) > self::MAX_CHARACTERS_PER_CELL) {
                throw new InvalidArgumentException('Trying to add a value that exceeds the maximum number of characters allowed in a cell (32,767)');
            }
            if ($this->shouldUseInlineStrings) {
                $cellXML->addAttribute("t", "inlineStr");
                $is = $cellXML->addChild("is");
                $is->addChild("t", $this->stringsEscaper->escape($cell->getValue()));
            }
            else {
                $cellXML->addAttribute("t", "s");
                $cellXML->addChild("v", $this->sharedStringsManager->writeString($cell->getValue()));
            }
        }
        elseif ($cell->isBoolean()) {
            $cellXML->addAttribute("t", "b");
            $cellXML->addChild("v", (int)$cell->getValue());
        }
        elseif ($cell->isNumeric()) {
            $cellXML->addChild("v", $this->stringHelper->formatNumericValue($cell->getValue()));
        }
        elseif ($cell->isError() && is_string($cell->getValueEvenIfError())) {
            $cellXML->addAttribute("t", "e");
            $cellXML->addChild("v", $cell->getValueEvenIfError());
            // only writes the error value if it's a string
        }
        elseif (!$cell->isEmpty()) {
            throw new InvalidArgumentException('Trying to add a value with an unsupported type: ' . gettype($cell->getValue()));
        }
    }
    
    /**
     * @throws InvalidArgumentException
     */
    private function applyStyleAndRegister(Cell $cell, Style $rowStyle): RegisteredStyle
    {
        $isMatchingRowStyle = false;
        
        if ($cell->getStyle()->isEmpty()) {
            $cell->setStyle($rowStyle);

            $possiblyUpdatedStyle = $this->styleManager->applyExtraStylesIfNeeded($cell);

            if ($possiblyUpdatedStyle->isUpdated()) {
                $registeredStyle = $this->styleManager->registerStyle($possiblyUpdatedStyle->getStyle());
            }
            else {
                $registeredStyle = $this->styleManager->registerStyle($rowStyle);
                $isMatchingRowStyle = true;
            }
        }
        else {
            $mergedCellAndRowStyle = $this->styleMerger->merge($cell->getStyle(), $rowStyle);
            $cell->setStyle($mergedCellAndRowStyle);

            $possiblyUpdatedStyle = $this->styleManager->applyExtraStylesIfNeeded($cell);

            if ($possiblyUpdatedStyle->isUpdated()) {
                $newCellStyle = $possiblyUpdatedStyle->getStyle();
            }
            else {
                $newCellStyle = $mergedCellAndRowStyle;
            }

            $registeredStyle = $this->styleManager->registerStyle($newCellStyle);
        }

        return new RegisteredStyle($registeredStyle, $isMatchingRowStyle);
    }
    
    /**
     * @throws Exception
     */
    public function close(Worksheet $worksheet): void
    {
        $mergeRanges = $worksheet->getExternalSheet()->getMergeRanges();
        
        if(!empty($mergeRanges)) {
            $mergeCellsXML = $worksheet->getXML()->mergeCells ?? $worksheet->getXML()->addChild("mergeCells");
            
            foreach ($mergeRanges as $mergeRange) {
                $mergeRangePairs    = explode(":", $mergeRange);
                $mergeRangePairs[0] = preg_replace("/[^a-zA-Z]+/", "", $mergeRangePairs[0]) . ((int)preg_replace("/[^0-9]+/", "", $mergeRangePairs[0]) + $worksheet->getFirstRowIndex());
                $mergeRangePairs[1] = preg_replace("/[^a-zA-Z]+/", "", $mergeRangePairs[1]) . ((int)preg_replace("/[^0-9]+/", "", $mergeRangePairs[1]) + $worksheet->getFirstRowIndex());
                $mergeRange         = implode(":", $mergeRangePairs);
                
                $alreadyAddedMergeCell = false;
                
                foreach ($mergeCellsXML->children() as $xmlMergeCellElement) {
                    if (isset($xmlMergeCellElement["ref"]) && (string)$xmlMergeCellElement["ref"] === $mergeRange) {
                        $alreadyAddedMergeCell = true;
                        break;
                    }
                }
                
                if (!$alreadyAddedMergeCell) {
                    $mergeCellXML = $mergeCellsXML->addChild("mergeCell");
                    $mergeCellXML->addAttribute("ref", $mergeRange);
                    unset($mergeCellXML);
                }
                unset($alreadyAddedMergeCell);
            }
            unset($mergeCellsXML);
        }
        
        $mergeRanges = [];
        foreach ($worksheet->getXML()->xpath("xmlns:mergeCells/xmlns:mergeCell") as $xmlMergeCellElement) {
            if (isset($xmlMergeCellElement["ref"])) {
                $mergeRanges[] = (string)$xmlMergeCellElement["ref"];
            }
        }
        $worksheet->getExternalSheet()->setMergeRanges($mergeRanges);
        
        $columnsWidth    = $worksheet->getExternalSheet()->getColumnsWidth();
        $autoSizeColumns = $worksheet->getExternalSheet()->getAutoSizeColumns();
        
        if (!empty($columnsWidth) || !empty($autoSizeColumns)) {
            $colsXML           = $worksheet->getXML()->cols ?? $worksheet->getXML()->prependChild("cols");
            $mergedCellsRanges = WorksheetHelper::getExtractCellsReferences($worksheet->getExternalSheet());
            
            if (!empty($columnsWidth)) {
                foreach ($columnsWidth as $column => $width) {
                    $columnIndex = CellHelper::getColumnIndexFromString($column);
                    
                    if ($columnIndex <= $worksheet->getMaxNumColumns()) {
                        $this->addColIfNotExistsOrGreaterWidth($colsXML, $columnIndex, $width);
                    }
                }
            }
            if (!empty($autoSizeColumns)) {
                foreach ($autoSizeColumns as $autoSizeColumn) {
                    $columnIndex = CellHelper::getColumnIndexFromString($autoSizeColumn);
                    
                    if ($columnIndex <= $worksheet->getMaxNumColumns()) {
                        $maxWidth = ColumnHelper::getDefaultColumnWidth($worksheet->getFontStyle()->getFontName(), $worksheet->getFontStyle()->getFontSize());
                        
                        for ($rowIndex = 1; $rowIndex <= $worksheet->getLastWrittenRowIndex(); $rowIndex++) {
                            if ($worksheet->getRowAtIndex($rowIndex)
                                && $worksheet->getRowAtIndex($rowIndex)->getCellAtIndex($columnIndex - 1)
                                && !$worksheet->getRowAtIndex($rowIndex)->getCellAtIndex($columnIndex - 1)->isEmpty()
                                && !isset($mergedCellsRanges[CellHelper::getColumnLettersFromColumnIndex($columnIndex - 1) . $rowIndex]) // if cell is not in a range calculate his width
                                && $maxWidth < ($cellWidth = ColumnHelper::calculateCellWidth(
                                    $worksheet->getFontStyle()->getFontName(),
                                    $worksheet->getFontStyle()->getFontSize(),
                                    $worksheet->getRowAtIndex($rowIndex)->getCellAtIndex($columnIndex - 1)->getValue()
                                )))
                            {
                                $maxWidth = $cellWidth;
                            }
                        }
                        
                        $this->addColIfNotExistsOrGreaterWidth($colsXML, $columnIndex, $maxWidth);
                    }
                }
            }
            unset($colsXML);
        }
        $worksheet->getXML()->saveXML($worksheet->getFilePath());
    }
    
    private function addColIfNotExistsOrGreaterWidth(ExtendedSimpleXMLElement $colsXML, int $columnIndex, $width): void
    {
        $alreadyAddedCol = false;
        
        foreach ($colsXML->children() as $xmlColElement) {
            if (isset($xmlColElement["min"]) && (int)$xmlColElement["min"] === $columnIndex
                && isset($xmlColElement["max"]) && (int)$xmlColElement["max"] === $columnIndex
                && isset($xmlColElement["bestFit"]) && (string)$xmlColElement["bestFit"] === "true"
                && isset($xmlColElement["customWidth"]) && (string)$xmlColElement["customWidth"] === "true"
                && isset($xmlColElement["style"]) && (int)$xmlColElement["style"] === 0)
            {
                $alreadyAddedCol = true;
                
                if (isset($xmlColElement["width"]) && (float)$xmlColElement["width"] < (float)$width) {
                    $xmlColElement["width"] = $width;
                }
                break;
            }
        }
        
        if (!$alreadyAddedCol) {
            $colXML = $colsXML->addChild("col");
            $colXML->addAttribute("min", $columnIndex);
            $colXML->addAttribute("max", $columnIndex);
            $colXML->addAttribute("width", $width);
            $colXML->addAttribute("bestFit", "true");
            $colXML->addAttribute("customWidth", "true");
            $colXML->addAttribute("style", "0");
            unset($colXML);
        }
    }
}