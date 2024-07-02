<?php

namespace Threshold\PhpExcel\Writer\Manager;

use Threshold\PhpExcel\Writer\Entity\Sheet;
use Threshold\PhpExcel\Writer\Exception\InvalidSheetNameException;
use Threshold\PhpExcel\Writer\Helper\StringHelper;

class SheetManager
{
    /* Sheet name should not exceed 31 characters */
    const MAX_LENGTH_SHEET_NAME = 31;

    private static array $INVALID_CHARACTERS_IN_SHEET_NAME = ['\\', '/', '?', '*', ':', '[', ']'];
    private static array $SHEETS_NAME_USED = [];
    private StringHelper $stringHelper;

    public function __construct(StringHelper $stringHelper)
    {
        $this->stringHelper = $stringHelper;
    }

    /**
     * @throws InvalidSheetNameException
     */
    public function throwIfNameIsInvalid($name, Sheet $sheet): void
    {
        if (!is_string($name)) {
            $actualType   = gettype($name);
            $errorMessage = "The sheet's name is invalid. It must be a string ($actualType given).";
            throw new InvalidSheetNameException($errorMessage);
        }

        $failedRequirements = [];
        $nameLength         = $this->stringHelper->getStringLength($name);

        if (!$this->isNameUnique($name, $sheet)) {
            $failedRequirements[] = 'It should be unique';
        }
        else {
            if ($nameLength === 0) {
                $failedRequirements[] = 'It should not be blank';
            }
            else {
                if ($nameLength > self::MAX_LENGTH_SHEET_NAME) {
                    $failedRequirements[] = 'It should not exceed 31 characters';
                }

                if ($this->doesContainInvalidCharacters($name)) {
                    $failedRequirements[] = 'It should not contain these characters: \\ / ? * : [ or ]';
                }

                if ($this->doesStartOrEndWithSingleQuote($name)) {
                    $failedRequirements[] = 'It should not start or end with a single quote';
                }
            }
        }

        if (count($failedRequirements) !== 0) {
            $errorMessage  = "The sheet's name (\"$name\") is invalid. It did not respect these rules:\n - ";
            $errorMessage .= implode("\n - ", $failedRequirements);
            throw new InvalidSheetNameException($errorMessage);
        }
    }

    private function doesContainInvalidCharacters(string $name): bool
    {
        return str_replace(self::$INVALID_CHARACTERS_IN_SHEET_NAME, '', $name) !== $name;
    }

    private function doesStartOrEndWithSingleQuote(string $name): bool
    {
        $startsWithSingleQuote = $this->stringHelper->getCharFirstOccurrencePosition('\'', $name) === 0;
        $endsWithSingleQuote   = $this->stringHelper->getCharLastOccurrencePosition('\'', $name) === ($this->stringHelper->getStringLength($name) - 1);

        return $startsWithSingleQuote || $endsWithSingleQuote;
    }

    private function isNameUnique(string $name, Sheet $sheet): bool
    {
        foreach (self::$SHEETS_NAME_USED[$sheet->getAssociatedWorkbookId()] as $sheetIndex => $sheetName) {
            if ($sheetIndex !== $sheet->getIndex() && $sheetName === $name) {
                return false;
            }
        }

        return true;
    }

    public function markWorkbookIdAsUsed($workbookId): void
    {
        if (!isset(self::$SHEETS_NAME_USED[$workbookId])) {
            self::$SHEETS_NAME_USED[$workbookId] = [];
        }
    }

    public function markSheetNameAsUsed(Sheet $sheet): void
    {
        self::$SHEETS_NAME_USED[$sheet->getAssociatedWorkbookId()][$sheet->getIndex()] = $sheet->getName();
    }
}
