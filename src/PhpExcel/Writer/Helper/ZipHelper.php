<?php

namespace Threshold\PhpExcel\Writer\Helper;

use FilesystemIterator;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use ZipArchive;

class ZipHelper
{
    const ZIP_EXTENSION = '.zip';

    /* Controls what to do when trying to add an existing file */
    const EXISTING_FILES_SKIP = 'skip';
    const EXISTING_FILES_OVERWRITE = 'overwrite';

    public function createZip(string $tmpFolderPath): ZipArchive
    {
        $zip         = new ZipArchive();
        $zipFilePath = $tmpFolderPath . self::ZIP_EXTENSION;

        $zip->open($zipFilePath, ZipArchive::CREATE | ZipArchive::OVERWRITE);

        return $zip;
    }

    public function getZipFilePath(ZipArchive $zip): string
    {
        return $zip->filename;
    }

    public function addFileToArchive(ZipArchive $zip, string $rootFolderPath, string $localFilePath,
                                     string $existingFileMode = self::EXISTING_FILES_OVERWRITE): void
    {
        $this->addFileToArchiveWithCompressionMethod(
            $zip,
            $rootFolderPath,
            $localFilePath,
            $existingFileMode,
            ZipArchive::CM_DEFAULT
        );
    }

    public function addUncompressedFileToArchive(ZipArchive $zip, string $rootFolderPath, string $localFilePath,
                                                 string $existingFileMode = self::EXISTING_FILES_OVERWRITE): void
    {
        $this->addFileToArchiveWithCompressionMethod(
            $zip,
            $rootFolderPath,
            $localFilePath,
            $existingFileMode,
            ZipArchive::CM_STORE
        );
    }

    protected function addFileToArchiveWithCompressionMethod(ZipArchive $zip, string $rootFolderPath,
                                                             string $localFilePath, string $existingFileMode,
                                                             int $compressionMethod): void
    {
        if (!$this->shouldSkipFile($zip, $localFilePath, $existingFileMode)) {
            $normalizedFullFilePath = $this->getNormalizedRealPath($rootFolderPath . '/' . $localFilePath);
            $zip->addFile($normalizedFullFilePath, $localFilePath);

            if (self::canChooseCompressionMethod()) {
                $zip->setCompressionName($localFilePath, $compressionMethod);
            }
        }
    }

    public static function canChooseCompressionMethod(): bool
    {
        // setCompressionName() is a PHP7+ method...
        return (method_exists(new ZipArchive(), 'setCompressionName'));
    }

    public function addFolderToArchive(ZipArchive $zip, string $folderPath,
                                       string $existingFileMode = self::EXISTING_FILES_OVERWRITE): void
    {
        $folderRealPath = $this->getNormalizedRealPath($folderPath) . '/';
        $itemIterator = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($folderPath, FilesystemIterator::SKIP_DOTS), RecursiveIteratorIterator::SELF_FIRST);

        foreach ($itemIterator as $itemInfo) {
            $itemRealPath = $this->getNormalizedRealPath($itemInfo->getPathname());
            $itemLocalPath = str_replace($folderRealPath, '', $itemRealPath);

            if ($itemInfo->isFile() && !$this->shouldSkipFile($zip, $itemLocalPath, $existingFileMode)) {
                $zip->addFile($itemRealPath, $itemLocalPath);
            }
        }
    }

    protected function shouldSkipFile(ZipArchive $zip, string $itemLocalPath, string $existingFileMode): bool
    {
        // Skip files if:
        //   - EXISTING_FILES_SKIP mode chosen
        //   - File already exists in the archive
        return $existingFileMode === self::EXISTING_FILES_SKIP && $zip->locateName($itemLocalPath) !== false;
    }

    protected function getNormalizedRealPath(string $path): string
    {
        return str_replace(DIRECTORY_SEPARATOR, '/', realpath($path));
    }
    
    /**
     * @var ZipArchive $zip
     * @var resource $streamPointer
     */
    public function closeArchiveAndCopyToStream(ZipArchive $zip, $streamPointer): void
    {
        $zipFilePath = $zip->filename;
        $zip->close();
        $this->copyZipToStream($zipFilePath, $streamPointer);
    }

    /**
     * @param resource $pointer Pointer to the stream to copy the zip
     */
    protected function copyZipToStream(string $zipFilePath, $pointer): void
    {
        $zipFilePointer = fopen($zipFilePath, 'r');
        stream_copy_to_stream($zipFilePointer, $pointer);
        fclose($zipFilePointer);
    }
}
