<?php

namespace Threshold\PhpExcel\Writer;

use Exception;
use Threshold\PhpExcel\Writer\Entity\{Options, Row, Sheet, Style\Style, Workbook};
use Threshold\PhpExcel\Writer\Exception\{InvalidArgumentException,
    InvalidSheetNameException,
    InvalidStyleException,
    InvalidWidthException,
    IOException,
    SheetNotFoundException,
    WriterAlreadyOpenedException,
    WriterNotOpenedException};
use Threshold\PhpExcel\Writer\Helper\{Escaper\XLSX, FileSystemHelper, StringHelper, ZipHelper};
use Threshold\PhpExcel\Writer\Manager\{OptionsManager,
    RowManager,
    SharedStringsManager,
    Style\StyleManager,
    Style\StyleMerger,
    Style\StyleRegistry,
    WorkbookManager,
    WorksheetManager};

class Writer
{
    private string $outputFilePath;
    /** @var resource */
    private $filePointer;
    private bool $writerOpened = false;
    private OptionsManager $optionsManager;
    private ?WorkbookManager $workbookManager = null;
    private static string $headerContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    
    public function __construct()
    {
        $this->optionsManager = new OptionsManager();
    }
    
    /**
     * @throws WriterAlreadyOpenedException
     */
    public function setShouldCreateNewSheetsAutomatically(bool $shouldCreateNewSheetsAutomatically): self
    {
        $this->throwIfWriterAlreadyOpened('Writer must be configured before opening it.');
        
        $this->optionsManager->setOption(Options::SHOULD_CREATE_NEW_SHEETS_AUTOMATICALLY, $shouldCreateNewSheetsAutomatically);
        
        return $this;
    }
    
    /**
     * @throws InvalidArgumentException|InvalidWidthException|InvalidSheetNameException|InvalidStyleException|IOException
     */
    protected function openWriterChooser(bool $withoutDefaultWorksheet = false)
    {
        if ($withoutDefaultWorksheet) {
            $this->openWriterWithoutDefaultWorksheet();
        }
        else {
            $this->openWriter();
        }
    }
    
    /**
     * @throws InvalidArgumentException|InvalidSheetNameException|InvalidStyleException|InvalidWidthException|IOException
     */
    protected function openWriter(): void
    {
        if (!$this->workbookManager) {
            $fileSystemHelper = new FileSystemHelper($this->optionsManager, new ZipHelper(), new XLSX());
            $fileSystemHelper->createBaseFilesAndFolders();
            
            $xlFolder     = $fileSystemHelper->getXlFolder();
            $styleMerger  = new StyleMerger();
            $styleManager = new StyleManager(new StyleRegistry($this->optionsManager->getOption(Options::DEFAULT_ROW_STYLE), ($xlFolder . "/" . $fileSystemHelper::STYLES_XML_FILE_NAME)));
            
            $this->workbookManager = new WorkbookManager(
                new Workbook(),
                $this->optionsManager,
                new WorksheetManager($this->optionsManager, new RowManager(), $styleManager, $styleMerger, new SharedStringsManager($xlFolder, new XLSX()), new XLSX(), new StringHelper()),
                $styleManager,
                $styleMerger,
                $fileSystemHelper
            );
            
            $this->workbookManager->reloadWorksheets();
            
            $worksheet = $this->workbookManager->addNewSheetAndMakeItCurrent();
            $worksheet->setFontStyle($this->optionsManager->getFontStyle());
        }
    }
    
    /**
     * @throws InvalidArgumentException|InvalidStyleException|InvalidWidthException|IOException
     */
    protected function openWriterWithoutDefaultWorksheet(): void
    {
        if (!$this->workbookManager) {
            $fileSystemHelper = new FileSystemHelper($this->optionsManager, new ZipHelper(), new XLSX());
            $fileSystemHelper->createBaseFilesAndFolders();
            
            $xlFolder     = $fileSystemHelper->getXlFolder();
            $styleMerger  = new StyleMerger();
            $styleManager = new StyleManager(new StyleRegistry($this->optionsManager->getOption(Options::DEFAULT_ROW_STYLE), ($xlFolder . "/" . $fileSystemHelper::STYLES_XML_FILE_NAME)));
            
            $this->workbookManager = new WorkbookManager(
                new Workbook(),
                $this->optionsManager,
                new WorksheetManager($this->optionsManager, new RowManager(), $styleManager, $styleMerger, new SharedStringsManager($xlFolder, new XLSX()), new XLSX(), new StringHelper()),
                $styleManager,
                $styleMerger,
                $fileSystemHelper
            );
            
            $this->workbookManager->reloadWorksheets();
        }
    }
    
    /**
     * @throws WriterNotOpenedException
     */
    public function getSheets(): array
    {
        $this->throwIfWorkbookIsNotAvailable();
        
        $externalSheets = [];
        $worksheets = $this->workbookManager->getWorksheets();
        
        foreach ($worksheets as $worksheet) {
            $externalSheets[] = $worksheet->getExternalSheet();
        }
        
        return $externalSheets;
    }
    
    /**
     * @throws InvalidSheetNameException|WriterNotOpenedException
     */
    public function addNewSheetAndMakeItCurrent(): Sheet
    {
        $this->throwIfWorkbookIsNotAvailable();
        $worksheet = $this->workbookManager->addNewSheetAndMakeItCurrent();
        $worksheet->setFontStyle($this->optionsManager->getFontStyle());
        
        return $worksheet->getExternalSheet();
    }
    
    /**
     * @throws WriterNotOpenedException
     */
    public function getCurrentSheet(): Sheet
    {
        $this->throwIfWorkbookIsNotAvailable();
        
        return $this->workbookManager->getCurrentWorksheet()->getExternalSheet();
    }
    
    /**
     * @throws InvalidSheetNameException|WriterNotOpenedException|SheetNotFoundException
     */
    public function getSheetByName(string $sheetName): Sheet
    {
        $this->throwIfWorkbookIsNotAvailable();
        
        foreach ($this->workbookManager->getWorksheets() as $worksheet) {
            if ($worksheet->getExternalSheet()->getName() === $sheetName) {
                $this->workbookManager->setCurrentSheet($worksheet->getExternalSheet());
                return $worksheet->getExternalSheet()->setNew(false);
            }
        }
        return $this->addNewSheetAndMakeItCurrent()->setName($sheetName);
    }
    
    /**
     * @throws SheetNotFoundException|WriterNotOpenedException
     */
    public function setCurrentSheet(Sheet $sheet): void
    {
        $this->throwIfWorkbookIsNotAvailable();
        $this->workbookManager->setCurrentSheet($sheet);
    }
    
    /**
     * @throws WriterNotOpenedException
     */
    protected function throwIfWorkbookIsNotAvailable(): void
    {
        if (!$this->workbookManager->getWorkbook()) {
            throw new WriterNotOpenedException('The writer must be opened before performing this action.');
        }
    }
    
    /**
     * @throws InvalidArgumentException|InvalidSheetNameException|WriterNotOpenedException
     */
    protected function addRowToWriter(Row $row, ?int $rowIndex = null): void
    {
        $this->throwIfWorkbookIsNotAvailable();
        $this->workbookManager->addRowToCurrentWorksheet($row, $rowIndex);
    }
    
    public function setDefaultRowStyle(Style $defaultStyle): self
    {
        $defaultStyle
            ->setFontSize($defaultStyle->getFontSize())
            ->setFontColor($defaultStyle->getFontColor())
            ->setFontName($defaultStyle->getFontName());
        $this->optionsManager->setOption(Options::DEFAULT_ROW_STYLE, $defaultStyle);
        
        return $this;
    }
    
    /**
     * @throws InvalidArgumentException|InvalidSheetNameException|InvalidStyleException|InvalidWidthException|IOException
     */
    public function openToFile(string $outputFilePath, bool $withoutDefaultWorksheet = false): self
    {
        $this->outputFilePath = $outputFilePath;
        $loopCounter          = $this->optionsManager->getOption(Options::LOOP_COUNTER);
        $loopMax              = $this->optionsManager->getOption(Options::LOOP_MAX);
        
        if (!($loopCounter && $loopMax) || $loopCounter === $loopMax) {
            $this->filePointer = fopen($this->outputFilePath, 'wb+');
            $this->throwIfFilePointerIsNotAvailable();
        }
        
        $this->openWriterChooser($withoutDefaultWorksheet);
        $this->writerOpened = true;
        
        return $this;
    }
    
    /**
     * @throws InvalidArgumentException|InvalidSheetNameException|InvalidStyleException|InvalidWidthException|IOException
     */
    public function openToBrowser(string $outputFileName, bool $withoutDefaultWorksheet = false): self
    {
        $this->outputFilePath = basename($outputFileName);
        $this->filePointer    = fopen('php://output', 'w');
        $this->throwIfFilePointerIsNotAvailable();
        
        // Clear any previous output (otherwise the generated file will be corrupted)
        // @see https://github.com/box/spout/issues/241
        if (ob_get_length() > 0) {
            ob_end_clean();
        }
        
        /*
         * Set headers
         *
         * For newer browsers such as Firefox, Chrome, Opera, Safari, etc., they all support and use `filename*`
         * specified by the new standard, even if they do not automatically decode filename; it does not matter;
         * and for older versions of Internet Explorer, they are not recognized `filename*`, will automatically
         * ignore it and use the old `filename` (the only minor flaw is that there must be an English suffix name).
         * In this way, the multi-browser multi-language compatibility problem is perfectly solved, which does not
         * require UA judgment and is more in line with the standard.
         *
         * @see https://github.com/box/spout/issues/745
         * @see https://tools.ietf.org/html/rfc6266
         * @see https://developer.mozilla.org/en-US/docs/Web/HTTP/Headers/Content-Disposition
         */
        header('Content-Type: ' . static::$headerContentType);
        header(
            'Content-Disposition: attachment; ' .
            'filename="' . rawurldecode($this->outputFilePath) . '"; ' .
            'filename*=UTF-8\'\'' . rawurldecode($this->outputFilePath)
        );
        
        /*
         * When forcing the download of a file over SSL,IE8 and lower browsers fail
         * if the Cache-Control and Pragma headers are not set.
         *
         * @see http://support.microsoft.com/KB/323308
         * @see https://github.com/liuggio/ExcelBundle/issues/45
         */
        header('Cache-Control: max-age=0');
        header('Pragma: public');
        
        $this->openWriterChooser($withoutDefaultWorksheet);
        $this->writerOpened = true;
        
        return $this;
    }
    
    /**
     * @throws IOException
     */
    private function throwIfFilePointerIsNotAvailable(): void
    {
        if (!$this->filePointer) {
            throw new IOException('File pointer has not be opened');
        }
    }
    
    /**
     * @throws WriterAlreadyOpenedException
     */
    protected function throwIfWriterAlreadyOpened(string $message): void
    {
        if ($this->writerOpened) {
            throw new WriterAlreadyOpenedException($message);
        }
    }
    
    /**
     * @throws Exception|IOException|WriterNotOpenedException
     */
    public function addRow(Row $row, ?int $rowIndex = null): self
    {
        if ($this->writerOpened) {
            try {
                if (!$row->getStyle()) {
                    $row->setStyle($this->optionsManager->getFontStyle());
                }
                $this->addRowToWriter($row, $rowIndex);
            }
            catch (Exception $e) {
                // if an exception occurs while writing data,
                // close the writer and remove all files created so far.
                $this->closeAndAttemptToCleanupAllFiles();
                
                // re-throw the exception to alert developers of the error
                throw $e;
            }
        }
        else {
            throw new WriterNotOpenedException('The writer needs to be opened before adding row.');
        }
        
        return $this;
    }
    
    /**
     * @throws Exception|IOException|WriterNotOpenedException
     */
    public function setRow(int $rowIndex, Row $row): self
    {
        if ($this->writerOpened) {
            try {
                if (!$row->getStyle()) {
                    $row->setStyle($this->optionsManager->getFontStyle());
                }
                $this->addRowToWriter($row, $rowIndex);
            }
            catch (Exception $e) {
                // if an exception occurs while writing data,
                // close the writer and remove all files created so far.
                $this->closeAndAttemptToCleanupAllFiles();
                
                // re-throw the exception to alert developers of the error
                throw $e;
            }
        }
        else {
            throw new WriterNotOpenedException('The writer needs to be opened before adding row.');
        }
        
        return $this;
    }
    
    /**
     * @throws InvalidArgumentException|IOException|WriterNotOpenedException
     */
    public function addRows(array $rows): self
    {
        foreach ($rows as $row) {
            if (!$row instanceof Row) {
                $this->closeAndAttemptToCleanupAllFiles();
                throw new InvalidArgumentException('The input should be an array of Row');
            }
            
            $this->addRow($row);
        }
        
        return $this;
    }
    
    /**
     * @return null|bool|string
     * @throws IOException
     */
    public function close()
    {
        if (!$this->writerOpened) {
            return null;
        }
        
        if (!($loopCounter = $this->optionsManager->getOption(Options::LOOP_COUNTER))
            || (($loopMax = $this->optionsManager->getOption(Options::LOOP_MAX)) && (int)$loopCounter >= (int)$loopMax))
        {
            if ($this->workbookManager) {
                $this->workbookManager->close($this->filePointer);
            }
            
            $this->writerOpened = false;
            
            return true;
        }
        else {
            if ($this->workbookManager) {
                $this->workbookManager->continue();
            }
            
            $this->writerOpened = false;
            
            return $this->optionsManager->getOption(Options::TEMP_FOLDER_NAME);
        }
    }
    
    /**
     * @throws IOException
     */
    private function closeAndAttemptToCleanupAllFiles(): void
    {
        // close the writer, which should remove all temp files
        $this->close();
        
        // remove output file if it was created
        if (file_exists($this->outputFilePath) && is_file($this->outputFilePath)) {
//            $fileSystemHelper = new FileSystemHelper($this->optionsManager, new ZipHelper(), new XLSX());
            unlink($this->outputFilePath);
        }
    }
    
    /**
     * @throws WriterAlreadyOpenedException
     */
    public function setTempFolder(string $tempFolder): self
    {
        $this->throwIfWriterAlreadyOpened('Writer must be configured before opening it.');

        $this->optionsManager->setOption(Options::TEMP_FOLDER, $tempFolder);

        return $this;
    }
    
    /**
     * @throws WriterAlreadyOpenedException
     */
    public function setTempFolderName(string $tempFolderName): self
    {
        $this->throwIfWriterAlreadyOpened('Writer must be configured before opening it.');

        $this->optionsManager->setOption(Options::TEMP_FOLDER_NAME, $tempFolderName);

        return $this;
    }

    /**
     * @throws WriterAlreadyOpenedException
     */
    public function setShouldUseInlineStrings(bool $shouldUseInlineStrings): self
    {
        $this->throwIfWriterAlreadyOpened('Writer must be configured before opening it.');

        $this->optionsManager->setOption(Options::SHOULD_USE_INLINE_STRINGS, $shouldUseInlineStrings);

        return $this;
    }
    
    /**
     * @throws WriterAlreadyOpenedException
     */
    public function setLoop(int $loopMax, int $loopCounter): self
    {
        $this->throwIfWriterAlreadyOpened('Writer must be configured before opening it.');
        
        $this->optionsManager->setOption(Options::LOOP_MAX, $loopMax);
        $this->optionsManager->setOption(Options::LOOP_COUNTER, $loopCounter);
        
        return $this;
    }
    
    public function getLastWrittenRowIndex(): int
    {
        return $this->workbookManager->getCurrentWorksheet()->getLastWrittenRowIndex();
    }
    
    public function sortSheetsByName(bool $reverse = false): self
    {
        $this->workbookManager->sortWorksheetsBySheetName($reverse);
        return $this;
    }
}