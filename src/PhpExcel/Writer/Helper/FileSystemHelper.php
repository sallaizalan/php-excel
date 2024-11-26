<?php

namespace Threshold\PhpExcel\Writer\Helper;

use DateTime;
use DateTimeInterface;
use FilesystemIterator;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use SimpleXMLElement;
use Threshold\PhpExcel\Writer\Entity\Options;
use Threshold\PhpExcel\Writer\Exception\IOException;
use Threshold\PhpExcel\Writer\Helper\Escaper\XLSX;
use Threshold\PhpExcel\Writer\Manager\{OptionsManager, Style\StyleManager};

class FileSystemHelper
{
    const APP_NAME = 'Spout';

    const RELS_FOLDER_NAME = '_rels';
    const DOC_PROPS_FOLDER_NAME = 'docProps';
    const XL_FOLDER_NAME = 'xl';
    const WORKSHEETS_FOLDER_NAME = 'worksheets';

    const RELS_FILE_NAME = '.rels';
    const APP_XML_FILE_NAME = 'app.xml';
    const CORE_XML_FILE_NAME = 'core.xml';
    const CONTENT_TYPES_XML_FILE_NAME = '[Content_Types].xml';
    const WORKBOOK_XML_FILE_NAME = 'workbook.xml';
    const WORKBOOK_RELS_XML_FILE_NAME = 'workbook.xml.rels';
    const STYLES_XML_FILE_NAME = 'styles.xml';

    private OptionsManager $optionsManager;
    private ZipHelper $zipHelper;
    private XLSX $escaper;
    private string $rootFolder;
    private bool $continueWriting = false;
    private string $relsFolder;
    private string $docPropsFolder;
    private string $xlFolder;
    private string $xlRelsFolder;
    private string $xlWorksheetsFolder;
    private string $baseFolderRealPath;

    public function __construct(OptionsManager $optionsManager, ZipHelper $zipHelper, XLSX $escaper)
    {
        $this->optionsManager     = $optionsManager;
        $this->baseFolderRealPath = realpath($this->optionsManager->getOption(Options::TEMP_FOLDER));
        $this->zipHelper          = $zipHelper;
        $this->escaper            = $escaper;
    }

    public function getRootFolder(): string
    {
        return $this->rootFolder;
    }

    public function getXlFolder(): string
    {
        return $this->xlFolder;
    }

    public function getXlWorksheetsFolder(): string
    {
        return $this->xlWorksheetsFolder;
    }
    
    /**
     * @throws IOException
     */
    public function createBaseFilesAndFolders(): void
    {
        $this
            ->createRootFolder()
            ->createRelsFolderAndFile()
            ->createDocPropsFolderAndFiles()
            ->createXlFolderAndSubFolders();
    }
    
    /**
     * @throws IOException
     */
    private function createRootFolder(): self
    {
        if (($loopMax = $this->optionsManager->getOption(Options::LOOP_MAX)) && ($loopCounter = $this->optionsManager->getOption(Options::LOOP_COUNTER))) {
            $this->setContinueWriting((int)$loopCounter !== 1 && (int)$loopCounter <= (int)$loopMax);
        }
        $tempFolderName   = $this->optionsManager->getOption(Options::TEMP_FOLDER_NAME) ?? uniqid('xlsx', true);
        $this->rootFolder = $this->createFolder($this->baseFolderRealPath, $tempFolderName);
        $this->optionsManager->setOption(Options::TEMP_FOLDER_NAME, $tempFolderName);
        
        return $this;
    }
    
    /**
     * @throws IOException
     */
    private function createRelsFolderAndFile(): self
    {
        $this->relsFolder = $this->createFolder($this->rootFolder, self::RELS_FOLDER_NAME);
        
        $this->createRelsFile();
        
        return $this;
    }
    
    /**
     * @throws IOException
     */
    private function createRelsFile(): self
    {
        if (!$this->continueWriting) {
            $relsFileContents = <<<'EOD'
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rIdWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    <Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rIdApp" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
EOD;
            
            $this->createFileWithContents($this->relsFolder, self::RELS_FILE_NAME, $relsFileContents);
        }
        
        return $this;
    }
    
    /**
     * @throws IOException
     */
    private function createDocPropsFolderAndFiles(): self
    {
        $this->docPropsFolder = $this->createFolder($this->rootFolder, self::DOC_PROPS_FOLDER_NAME);
        
        $this
            ->createAppXmlFile()
            ->createCoreXmlFile();
        
        return $this;
    }
    
    /**
     * @throws IOException
     */
    private function createAppXmlFile(): self
    {
        if (!$this->continueWriting) {
            $appName = self::APP_NAME;
            $appXmlFileContents = <<<EOD
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
    <Application>$appName</Application>
    <TotalTime>0</TotalTime>
</Properties>
EOD;
            
            $this->createFileWithContents($this->docPropsFolder, self::APP_XML_FILE_NAME, $appXmlFileContents);
        }
        
        return $this;
    }
    
    /**
     * @throws IOException
     */
    private function createCoreXmlFile(): self
    {
        if (!$this->continueWriting) {
            $createdDate = (new DateTime())->format(DateTimeInterface::W3C);
            $coreXmlFileContents = <<<EOD
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dcterms:created xsi:type="dcterms:W3CDTF">$createdDate</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">$createdDate</dcterms:modified>
    <cp:revision>0</cp:revision>
</cp:coreProperties>
EOD;
            
            $this->createFileWithContents($this->docPropsFolder, self::CORE_XML_FILE_NAME, $coreXmlFileContents);
        }
        
        return $this;
    }
    
    /**
     * @throws IOException
     */
    private function createXlFolderAndSubFolders(): self
    {
        $this->xlFolder           = $this->createFolder($this->rootFolder, self::XL_FOLDER_NAME);
        $this->xlRelsFolder       = $this->createFolder($this->xlFolder, self::RELS_FOLDER_NAME);
        $this->xlWorksheetsFolder = $this->createFolder($this->xlFolder, self::WORKSHEETS_FOLDER_NAME);
        
        return $this;
    }
    
    /**
     * @throws IOException
     */
    public function createContentTypesFile(array $worksheets): self
    {
        $contentTypesXmlFilePath = $this->rootFolder . '/' . self::CONTENT_TYPES_XML_FILE_NAME;
        
        if ($this->continueWriting && file_exists($contentTypesXmlFilePath)) {
            $contentTypesXmlFileContents = simplexml_load_file($contentTypesXmlFilePath);
        }
        else {
            $contentTypesXmlFileContents = new SimpleXMLElement('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>');
            $child = $contentTypesXmlFileContents->addChild("Default");
            $child->addAttribute("ContentType", "application/xml");
            $child->addAttribute("Extension", "xml");
            unset($child);
            
            $child = $contentTypesXmlFileContents->addChild("Default");
            $child->addAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
            $child->addAttribute("Extension", "rels");
            unset($child);
            
            $child = $contentTypesXmlFileContents->addChild("Override");
            $child->addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            $child->addAttribute("PartName", "/xl/workbook.xml");
            unset($child);
            
            $child = $contentTypesXmlFileContents->addChild("Override");
            $child->addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            $child->addAttribute("PartName", "/xl/styles.xml");
            unset($child);
            
            $child = $contentTypesXmlFileContents->addChild("Override");
            $child->addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
            $child->addAttribute("PartName", "/xl/sharedStrings.xml");
            unset($child);
            
            $child = $contentTypesXmlFileContents->addChild("Override");
            $child->addAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml");
            $child->addAttribute("PartName", "/docProps/core.xml");
            unset($child);
            
            $child = $contentTypesXmlFileContents->addChild("Override");
            $child->addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
            $child->addAttribute("PartName", "/docProps/app.xml");
            unset($child);
        }
        
        foreach ($worksheets as $worksheet) {
            $alreadyAddedWorksheet = false;
            
            foreach ($contentTypesXmlFileContents->children() as $xmlElement) {
                $partNameAttribute = $xmlElement->attributes()["PartName"] ?? null;
                
                if ($partNameAttribute && preg_match("/sheet" . $worksheet->getId() . "\.xml$/", $partNameAttribute)) {
                    $alreadyAddedWorksheet = true;
                    break;
                }
            }
            
            if (!$alreadyAddedWorksheet) {
                $child = $contentTypesXmlFileContents->addChild("Override");
                $child->addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                $child->addAttribute("PartName", "/xl/worksheets/sheet" . $worksheet->getId() . ".xml");
                unset($child);
            }
        }
        
        $this->throwIfOperationNotInBaseFolder($this->rootFolder);
        $contentTypesXmlFileContents->saveXML($contentTypesXmlFilePath);
        
        return $this;
    }
    
    /**
     * @throws IOException
     */
    public function createFolder(string $parentFolderPath, string $folderName): string
    {
        $this->throwIfOperationNotInBaseFolder($parentFolderPath);
        
        $folderPath = $parentFolderPath . '/' . $folderName;
        
        if (!file_exists($folderPath) && !mkdir($folderPath, 0777, true)) {
            throw new IOException("Unable to create folder: $folderPath");
        }
        
        return $folderPath;
    }
    
    /**
     * @throws IOException
     */
    public function createFileWithContents(string $parentFolderPath, string $fileName, string $fileContents): string
    {
        $this->throwIfOperationNotInBaseFolder($parentFolderPath);
        
        $filePath              = $parentFolderPath . '/' . $fileName;
        $wasCreationSuccessful = file_put_contents($filePath, $fileContents);
        
        if ($wasCreationSuccessful === false) {
            throw new IOException("Unable to create file: $filePath");
        }
        
        return $filePath;
    }
    
    /**
     * @throws IOException
     */
    public function deleteFile(string $filePath): void
    {
        $this->throwIfOperationNotInBaseFolder($filePath);
        
        if (file_exists($filePath) && is_file($filePath)) {
            unlink($filePath);
        }
    }
    
    /**
     * @throws IOException
     */
    public function deleteFolderRecursively(string $folderPath): void
    {
        $this->throwIfOperationNotInBaseFolder($folderPath);
        
        $itemIterator = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($folderPath, FilesystemIterator::SKIP_DOTS),
            RecursiveIteratorIterator::CHILD_FIRST
        );
        
        foreach ($itemIterator as $item) {
            if ($item->isDir()) {
                rmdir($item->getPathname());
            } else {
                unlink($item->getPathname());
            }
        }
        
        rmdir($folderPath);
    }
    
    /**
     * @throws IOException
     */
    private function throwIfOperationNotInBaseFolder(string $operationFolderPath): void
    {
        $operationFolderRealPath = realpath($operationFolderPath);
        
        if (!$this->baseFolderRealPath) {
            throw new IOException("The base folder path is invalid: $this->baseFolderRealPath");
        }
        
        if (strpos($operationFolderRealPath, $this->baseFolderRealPath) !== 0) {
            throw new IOException("Cannot perform I/O operation outside of the base folder: $this->baseFolderRealPath");
        }
    }
    
    /**
     * @throws IOException
     */
    public function createWorkbookFile(array $worksheets): self
    {
        $workbookXmlFilePath = $this->xlFolder . "/" . self::WORKBOOK_XML_FILE_NAME;
        
        if ($this->continueWriting && file_exists($workbookXmlFilePath)) {
            $workbookXmlFileContents = simplexml_load_file($workbookXmlFilePath);
        }
        else {
            $workbookXmlFileContents = new SimpleXMLElement('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></workbook>');
        }
        
        $workbookXmlFileContents->registerXPathNamespace("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        $workbookXmlFileContents->registerXPathNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        
        if (!$sheets = $workbookXmlFileContents->sheets) {
            $sheets = $workbookXmlFileContents->addChild("sheets");
        }
        
        unset($workbookXmlFileContents->sheets->sheet); // unset all sheet, because of the sorting
        
        foreach ($worksheets as $worksheet) {
            $sheet = $sheets->addChild("sheet");
            $sheet->addAttribute("name", $this->escaper->escape($worksheet->getExternalSheet()->getName()));
            $sheet->addAttribute("sheetId", $worksheet->getId());
            $sheet->addAttribute("r:id", ("rIdSheet" . $worksheet->getId()), "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            $sheet->addAttribute("state", ($worksheet->getExternalSheet()->isVisible() ? 'visible' : 'hidden'));
            unset($sheet);
        }
        
        $this->throwIfOperationNotInBaseFolder($this->rootFolder);
        $workbookXmlFileContents->saveXML($workbookXmlFilePath);
        
        return $this;
    }
    
    /**
     * @throws IOException
     */
    public function createWorkbookRelsFile(array $worksheets): self
    {
        $workbookRelsXmlFilePath = $this->xlRelsFolder . "/" . self::WORKBOOK_RELS_XML_FILE_NAME;
        
        if ($this->continueWriting && file_exists($workbookRelsXmlFilePath)) {
            $workbookRelsXmlFileContents = simplexml_load_file($workbookRelsXmlFilePath);
        }
        else {
            $workbookRelsXmlFileContents = new SimpleXMLElement('<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
            $relationship = $workbookRelsXmlFileContents->addChild("Relationship");
            $relationship->addAttribute("Id", "rIdStyles");
            $relationship->addAttribute("Target", "styles.xml");
            $relationship->addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            unset($relationship);
            
            $relationship = $workbookRelsXmlFileContents->addChild("Relationship");
            $relationship->addAttribute("Id", "rIdSharedStrings");
            $relationship->addAttribute("Target", "sharedStrings.xml");
            $relationship->addAttribute("Type", "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings");
            unset($relationship);
        }
        
        foreach ($worksheets as $worksheet) {
            $alreadyAddedRelationship = false;
            
            foreach ($workbookRelsXmlFileContents->children() as $xmlElement) {
                $idAttribute = $xmlElement->attributes()["Id"] ?? null;
                
                if ($idAttribute && (string)$idAttribute === ("rIdSheet" . $worksheet->getId())) {
                    $alreadyAddedRelationship = true;
                    break;
                }
            }
            
            if (!$alreadyAddedRelationship) {
                $relationship = $workbookRelsXmlFileContents->addChild("Relationship");
                $relationship->addAttribute("Id", ("rIdSheet" . $worksheet->getId()));
                $relationship->addAttribute("Target", ("worksheets/sheet" . $worksheet->getId() . ".xml"));
                $relationship->addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                unset($relationship);
            }
        }
        
        $this->throwIfOperationNotInBaseFolder($this->rootFolder);
        $workbookRelsXmlFileContents->saveXML($workbookRelsXmlFilePath);
        
        return $this;
    }
    
    public function createStylesFile(StyleManager $styleManager): self
    {
        $styleManager->createStylesXMLFile($this->xlFolder . "/" . self::STYLES_XML_FILE_NAME);

        return $this;
    }
    
    /**
     * @var resource $filePointer
     * @throws IOException
     */
    public function zipRootFolderAndCopyToStream($filePointer): self
    {
        $zip = $this->zipHelper->createZip($this->rootFolder);

        $zipFilePath = $this->zipHelper->getZipFilePath($zip);

        // In order to have the file's mime type detected properly, files need to be added
        // to the zip file in a particular order.
        // "[Content_Types].xml" then at least 2 files located in "xl" folder should be zipped first.
        $this->zipHelper->addFileToArchive($zip, $this->rootFolder, self::CONTENT_TYPES_XML_FILE_NAME);
        $this->zipHelper->addFileToArchive($zip, $this->rootFolder, self::XL_FOLDER_NAME . '/' . self::WORKBOOK_XML_FILE_NAME);
        $this->zipHelper->addFileToArchive($zip, $this->rootFolder, self::XL_FOLDER_NAME . '/' . self::STYLES_XML_FILE_NAME);

        $this->zipHelper->addFolderToArchive($zip, $this->rootFolder, ZipHelper::EXISTING_FILES_SKIP);
        $this->zipHelper->closeArchiveAndCopyToStream($zip, $filePointer);

        // once the zip is copied, remove it
        $this->deleteFile($zipFilePath);
        
        return $this;
    }
    
    public function setContinueWriting(bool $continueWriting): self
    {
        $this->continueWriting = $continueWriting;
        return $this;
    }
}