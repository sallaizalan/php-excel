<?php

namespace Threshold\PhpExcel\Writer\Manager;

use Threshold\PhpExcel\Writer\Entity\Row;

class RowManager
{
    public function isEmpty(Row $row): bool
    {
        foreach ($row->getCells() as $cell) {
            if (!$cell->isEmpty()) {
                return false;
            }
        }

        return true;
    }
}
