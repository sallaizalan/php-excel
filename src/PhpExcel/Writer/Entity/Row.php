<?php

namespace Threshold\PhpExcel\Writer\Entity;

use Threshold\PhpExcel\Writer\Entity\Style\Style;

class Row
{
    private array $cells = [];
    private ?Style $style;

    public function __construct(array $cells, ?Style $style)
    {
        $this
            ->setCells($cells)
            ->setStyle($style);
    }

    public function getCells(): array
    {
        return $this->cells;
    }

    public function setCells(array $cells): self
    {
        $this->cells = [];
        foreach ($cells as $cell) {
            $this->addCell($cell);
        }
        return $this;
    }

    public function setCellAtIndex(Cell $cell, int $cellIndex): self
    {
        $this->cells[$cellIndex] = $cell;

        return $this;
    }

    public function getCellAtIndex(int $cellIndex): ?Cell
    {
        return $this->cells[$cellIndex] ?? null;
    }

    public function addCell(Cell $cell): self
    {
        $this->cells[] = $cell;

        return $this;
    }

    public function getNumCells(): int
    {
        // When using "setCellAtIndex", it's possible to
        // have "$this->cells" contain holes.
        if (empty($this->cells)) {
            return 0;
        }

        return max(array_keys($this->cells)) + 1;
    }

    public function getStyle(): ?Style
    {
        return $this->style;
    }

    public function setStyle(?Style $style): self
    {
        $this->style = $style;

        return $this;
    }

    public function toArray(): array
    {
        return array_map(function (Cell $cell) {
            return $cell->getValue();
        }, $this->cells);
    }
}
