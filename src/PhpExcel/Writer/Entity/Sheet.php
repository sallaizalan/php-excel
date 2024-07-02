<?php

namespace Threshold\PhpExcel\Writer\Entity;

use Threshold\PhpExcel\Writer\Exception\InvalidSheetNameException;
use Threshold\PhpExcel\Writer\Manager\SheetManager;

class Sheet
{
    const DEFAULT_SHEET_NAME_PREFIX = 'Sheet';

    private int $index;
    private string $associatedWorkbookId;
    private string $name;
    private bool $visible = true;
    private SheetManager $sheetManager;
    private array $mergeRanges;
    private array $autoSizeColumns;
    private array $columnsWidth;
    
    /**
     * @throws InvalidSheetNameException
     */
    public function __construct(int $sheetIndex, string $associatedWorkbookId, SheetManager $sheetManager)
    {
        $this->index                = $sheetIndex;
        $this->associatedWorkbookId = $associatedWorkbookId;
        $this->sheetManager         = $sheetManager;
        $this->sheetManager->markWorkbookIdAsUsed($associatedWorkbookId);

        $this->setName(self::DEFAULT_SHEET_NAME_PREFIX . ($sheetIndex + 1));
        $this->mergeRanges     = [];
        $this->autoSizeColumns = [];
        $this->columnsWidth    = [];
    }

    public function getIndex(): int
    {
        return $this->index;
    }

    public function getAssociatedWorkbookId(): string
    {
        return $this->associatedWorkbookId;
    }

    public function getName(): string
    {
        return $this->name;
    }

    /**
     * @throws InvalidSheetNameException
     */
    public function setName(string $name): self
    {
        $this->sheetManager->throwIfNameIsInvalid($name, $this);
        $this->name = $name;
        $this->sheetManager->markSheetNameAsUsed($this);
        return $this;
    }

    public function isVisible(): bool
    {
        return $this->visible;
    }

    public function setVisible(bool $visible): self
    {
        $this->visible = $visible;
        return $this;
    }
    
    /**
     * @return array
     */
    public function getMergeRanges(): array
    {
        return $this->mergeRanges;
    }
    
    public function setMergeRanges(array $mergeRanges): self
    {
        $this->mergeRanges = $mergeRanges;
        return $this;
    }
    
    public function getAutoSizeColumns(): array
    {
        return $this->autoSizeColumns;
    }
    
    public function setAutoSizeColumns(array $autoSizeColumns): self
    {
        $this->autoSizeColumns = $autoSizeColumns;
        return $this;
    }
    
    public function addAutoSizeColumns(string $autoSizeColumn): self
    {
        if (!in_array($autoSizeColumn, $this->autoSizeColumns)) {
            $this->autoSizeColumns[] = $autoSizeColumn;
        }
        return $this;
    }
    
    public function getColumnsWidth(): array
    {
        return $this->columnsWidth;
    }
    
    public function setColumnsWidth(array $columnsWidth): self
    {
        $this->columnsWidth = $columnsWidth;
        return $this;
    }
    
    public function addColumnWidth(string $column, int $width): self
    {
        if (!in_array($column, $this->autoSizeColumns)) {
            $this->columnsWidth[$column] = $width;
        }
        return $this;
    }
}