<?php

namespace Threshold\PhpExcel\Writer\Manager;

use Exception;
use Threshold\PhpExcel\Writer\Entity\{Options, Row, Sheet, Workbook, Worksheet};
use Threshold\PhpExcel\Writer\Exception\{InvalidArgumentException,
    InvalidSheetNameException,
    IOException,
    SheetNotFoundException};
use Threshold\PhpExcel\Writer\Helper\{FileSystemHelper, StringHelper};
use Threshold\PhpExcel\Writer\Manager\Style\{StyleManager, StyleMerger};

class WorkbookManager
{
    /*
     * Maximum number of rows a XLSX sheet can contain
     * @see http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
     */
    private static int $maxRowsPerWorksheet = 1048576;
    private Workbook $workbook;
    private OptionsManager $optionsManager;
    private WorksheetManager $worksheetManager;
    private StyleManager $styleManager;
    private StyleMerger $styleMerger;
    private FileSystemHelper $fileSystemHelper;
    private Worksheet $currentWorksheet;
    
    public function __construct(Workbook $workbook, OptionsManager $optionsManager, WorksheetManager $worksheetManager,
                                StyleManager $styleManager, StyleMerger $styleMerger, FileSystemHelper $fileSystemHelper)
    {
        $this->workbook         = $workbook;
        $this->optionsManager   = $optionsManager;
        $this->worksheetManager = $worksheetManager;
        $this->styleManager     = $styleManager;
        $this->styleMerger      = $styleMerger;
        $this->fileSystemHelper = $fileSystemHelper;
    }
    
    protected function getMaxRowsPerWorksheet(): int
    {
        return self::$maxRowsPerWorksheet;
    }

    public function getWorksheetFilePath(Sheet $sheet): string
    {
        $worksheetFilesFolder = $this->fileSystemHelper->getXlWorksheetsFolder();

        return $worksheetFilesFolder . '/' . strtolower($sheet->getName()) . '.xml';
    }
    
    public function getWorkbook(): Workbook
    {
        return $this->workbook;
    }
    
    /**
     * @throws InvalidSheetNameException
     */
    public function addNewSheetAndMakeItCurrent(): Worksheet
    {
        $worksheet = $this->addNewSheet();
        $this->setCurrentWorksheet($worksheet);
        
        return $worksheet;
    }
    
    /**
     * @throws Exception|InvalidSheetNameException
     */
    private function addNewSheet(): Worksheet
    {
        $worksheets        = $this->workbook->getWorksheets();
        $newSheetIndex     = count($worksheets);
        $sheet             = new Sheet($newSheetIndex, $this->workbook->getInternalId(), new SheetManager(new StringHelper()));
        $worksheetFilePath = $this->getWorksheetFilePath($sheet);
        $worksheet         = new Worksheet($worksheetFilePath, $sheet);
        
        $this->worksheetManager->startSheet($worksheet);
        
        $worksheets[] = $worksheet;
        $this->workbook->setWorksheets($worksheets);
        
        return $worksheet;
    }
    
    public function getWorksheets(): array
    {
        return $this->workbook->getWorksheets();
    }
    
    public function getCurrentWorksheet(): Worksheet
    {
        return $this->currentWorksheet;
    }
    
    /**
     * @throws SheetNotFoundException
     */
    public function setCurrentSheet(Sheet $sheet): void
    {
        $worksheet = $this->getWorksheetFromExternalSheet($sheet);
        
        if ($worksheet !== null) {
            $this->currentWorksheet = $worksheet;
        }
        else {
            throw new SheetNotFoundException('The given sheet does not exist in the workbook.');
        }
    }
    
    /**
     * @throws InvalidArgumentException|InvalidSheetNameException
     */
    public function addRowToCurrentWorksheet(Row $row, ?int $rowIndex = null): void
    {
        $currentWorksheet  = $this->getCurrentWorksheet();
        $hasReachedMaxRows = $this->hasCurrentWorksheetReachedMaxRows();
        
        // if we reached the maximum number of rows for the current sheet...
        if ($hasReachedMaxRows) {
            // ... continue writing in a new sheet if option set
            if ($this->optionsManager->getOption(Options::SHOULD_CREATE_NEW_SHEETS_AUTOMATICALLY)) {
                $currentWorksheet = $this->addNewSheetAndMakeItCurrent();
                
                $this->addRowToWorksheet($currentWorksheet, $row, $rowIndex);
            }
//            else {
//                 otherwise, do nothing as the data won't be written anyways
//            }
        }
        else {
            $this->addRowToWorksheet($currentWorksheet, $row, $rowIndex);
        }
    }
    
    /**
     * @var resource $filePointer
     * @throws Exception|IOException
     */
    public function close($filePointer): void
    {
        $this->worksheetManager->getSharedStringsManager()->close();
        $this->fileSystemHelper
            ->createContentTypesFile($this->getWorksheets())
            ->createWorkbookFile($this->getWorksheets())
            ->createWorkbookRelsFile($this->getWorksheets())
            ->createStylesFile($this->styleManager);
        
        foreach ($this->getWorksheets() as $worksheet) {
            $this->worksheetManager->close($worksheet);
        }
        
        $this->fileSystemHelper
            ->zipRootFolderAndCopyToStream($filePointer)
            ->deleteFolderRecursively($this->fileSystemHelper->getRootFolder());
    }
    
    /**
     * @throws Exception|IOException
     */
    public function continue(): void
    {
        $this->worksheetManager->getSharedStringsManager()->close();
        $this->fileSystemHelper
            ->createContentTypesFile($this->getWorksheets())
            ->createWorkbookFile($this->getWorksheets())
            ->createWorkbookRelsFile($this->getWorksheets())
            ->createStylesFile($this->styleManager);
        
        foreach ($this->getWorksheets() as $worksheet) {
            $this->worksheetManager->close($worksheet);
        }
    }
    
    private function setCurrentWorksheet(Worksheet $worksheet): void
    {
        $this->currentWorksheet = $worksheet;
    }
    
    private function getWorksheetFromExternalSheet(Sheet $sheet): ?Worksheet
    {
        $worksheetFound = null;
        
        foreach ($this->getWorksheets() as $worksheet) {
            if ($worksheet->getExternalSheet() === $sheet) {
                $worksheetFound = $worksheet;
                break;
            }
        }
        
        return $worksheetFound;
    }
    
    private function hasCurrentWorksheetReachedMaxRows(): bool
    {
        return $this->getCurrentWorksheet()->getLastWrittenRowIndex() >= $this->getMaxRowsPerWorksheet();
    }
    
    /**
     * @throws InvalidArgumentException
     */
    private function addRowToWorksheet(Worksheet $worksheet, Row $row, ?int $rowIndex = null): void
    {
        $this->applyDefaultRowStyle($row);
        $this->worksheetManager->addRow($worksheet, $row, $rowIndex);
        
        // update max num columns for the worksheet
        $currentMaxNumColumns = $worksheet->getMaxNumColumns();
        $cellsCount           = $row->getNumCells();
        $worksheet->setMaxNumColumns(max($currentMaxNumColumns, $cellsCount));
    }
    
    /**
     * @throws InvalidArgumentException
     */
    private function applyDefaultRowStyle(Row $row)
    {
        $defaultRowStyle = $this->optionsManager->getOption(Options::DEFAULT_ROW_STYLE);
        
        if ($defaultRowStyle !== null) {
            $mergedStyle = $this->styleMerger->merge($row->getStyle(), $defaultRowStyle);
            $row->setStyle($mergedStyle);
        }
    }
}