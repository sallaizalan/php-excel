<?php

namespace Threshold\PhpExcel\Writer\Factory;

use Threshold\PhpExcel\Writer\Entity\{Cell, Row, Style\Style};

class WriterEntityFactory
{
    public static function createRow(array $cells = [], ?Style $rowStyle = null): Row
    {
        return new Row($cells, $rowStyle);
    }

    public static function createRowFromArray(array $cellValues = [], ?Style $rowStyle = null): Row
    {
        $cells = array_map(function ($cellValue) {
            return new Cell($cellValue);
        }, $cellValues);

        return new Row($cells, $rowStyle);
    }

    public static function createCell($cellValue, ?Style $cellStyle = null): Cell
    {
        return new Cell($cellValue, $cellStyle);
    }
}
