<?php

namespace Threshold\PhpExcel\Writer\Entity;

class Workbook
{
    private array $worksheets = [];
    private string $internalId;

    public function __construct()
    {
        $this->internalId = uniqid();
    }

    public function getWorksheets(): array
    {
        return $this->worksheets;
    }

    public function setWorksheets(array $worksheets): self
    {
        $this->worksheets = $worksheets;
        return $this;
    }

    public function getInternalId(): string
    {
        return $this->internalId;
    }
}
