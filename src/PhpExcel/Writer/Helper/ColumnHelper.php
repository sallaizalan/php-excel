<?php

namespace Threshold\PhpExcel\Writer\Helper;

class ColumnHelper
{
    private const DEFAULT_COLUMN_WIDTHS = [
        'Arial' => [
            1 => ['px' => 24, 'width' => 12.00000000, 'height' => 5.25],
            2 => ['px' => 24, 'width' => 12.00000000, 'height' => 5.25],
            3 => ['px' => 32, 'width' => 10.66406250, 'height' => 6.0],
            
            4 => ['px' => 32, 'width' => 10.66406250, 'height' => 6.75],
            5 => ['px' => 40, 'width' => 10.00000000, 'height' => 8.25],
            6 => ['px' => 48, 'width' => 9.59765625, 'height' => 8.25],
            7 => ['px' => 48, 'width' => 9.59765625, 'height' => 9.0],
            8 => ['px' => 56, 'width' => 9.33203125, 'height' => 11.25],
            9 => ['px' => 64, 'width' => 9.14062500, 'height' => 12.0],
            10 => ['px' => 64, 'width' => 9.14062500, 'height' => 12.75],
        ],
        'Calibri' => [
            1 => ['px' => 24, 'width' => 12.00000000, 'height' => 5.25],
            2 => ['px' => 24, 'width' => 12.00000000, 'height' => 5.25],
            3 => ['px' => 32, 'width' => 10.66406250, 'height' => 6.00],
            4 => ['px' => 32, 'width' => 10.66406250, 'height' => 6.75],
            5 => ['px' => 40, 'width' => 10.00000000, 'height' => 8.25],
            6 => ['px' => 48, 'width' => 9.59765625, 'height' => 8.25],
            7 => ['px' => 48, 'width' => 9.59765625, 'height' => 9.0],
            8 => ['px' => 56, 'width' => 9.33203125, 'height' => 11.25],
            9 => ['px' => 56, 'width' => 9.33203125, 'height' => 12.0],
            10 => ['px' => 64, 'width' => 9.14062500, 'height' => 12.75],
            11 => ['px' => 64, 'width' => 9.14062500, 'height' => 15.0],
        ],
        'Verdana' => [
            1 => ['px' => 24, 'width' => 12.00000000, 'height' => 5.25],
            2 => ['px' => 24, 'width' => 12.00000000, 'height' => 5.25],
            3 => ['px' => 32, 'width' => 10.66406250, 'height' => 6.0],
            4 => ['px' => 32, 'width' => 10.66406250, 'height' => 6.75],
            5 => ['px' => 40, 'width' => 10.00000000, 'height' => 8.25],
            6 => ['px' => 48, 'width' => 9.59765625, 'height' => 8.25],
            7 => ['px' => 48, 'width' => 9.59765625, 'height' => 9.0],
            8 => ['px' => 64, 'width' => 9.14062500, 'height' => 10.5],
            9 => ['px' => 72, 'width' => 9.00000000, 'height' => 11.25],
            10 => ['px' => 72, 'width' => 9.00000000, 'height' => 12.75],
        ],
    ];
    
    public static function getDefaultColumnWidth(string $font, int $fontSize)
    {
        return self::DEFAULT_COLUMN_WIDTHS[$font][$fontSize]["width"] ?? (self::DEFAULT_COLUMN_WIDTHS[$font] ? self::DEFAULT_COLUMN_WIDTHS[$font][max(array_keys(self::DEFAULT_COLUMN_WIDTHS[$font]))]["width"] : self::DEFAULT_COLUMN_WIDTHS["Calibri"][11]["width"]);
    }
    
    public static function calculateCellWidth(string $font, int $fontSize, $cellText): float
    {
        $cellText = (string) $cellText;
        if (strpos($cellText, "\n") !== false) {
            $lineTexts = explode("\n", $cellText);
            $lineWidths = [];
            foreach ($lineTexts as $lineText) {
                $lineWidths[] = self::calculateCellWidth($font, $fontSize, $lineText);
            }
            
            return max($lineWidths); // width of longest line in cell
        }
        
        $stringLength = (new StringHelper())->getStringLength($cellText);
        switch ($font) {
            case 'Arial':
                // value 8 was set because of experience in different exports at Arial 10 font.
                $columnWidth = (int) (8 * $stringLength);
                $columnWidth = $columnWidth * $fontSize / 10; // extrapolate from font size
                
                break;
            case 'Verdana':
                // value 8 was found via interpolation by inspecting real Excel files with Verdana 10 font.
                $columnWidth = (int) (8 * $stringLength);
                $columnWidth = $columnWidth * $fontSize / 10; // extrapolate from font size
                
                break;
            default:
                // just assume Calibri
                // value 8.26 was found via interpolation by inspecting real Excel files with Calibri 11 font.
                $columnWidth = (int) (8.26 * $stringLength);
                $columnWidth = $columnWidth * $fontSize / 11; // extrapolate from font size
                
                break;
        }
        return round(self::pixelsToCellDimension($font, $fontSize, (int)$columnWidth), 4);
    }
    
    private static function pixelsToCellDimension(string $font, int $fontSize, $pixelValue)
    {
        return isset(self::DEFAULT_COLUMN_WIDTHS[$font][$fontSize]) ? ($pixelValue * self::DEFAULT_COLUMN_WIDTHS[$font][$fontSize]["width"] / self::DEFAULT_COLUMN_WIDTHS[$font][$fontSize]["px"]) : ($pixelValue * 11 * self::DEFAULT_COLUMN_WIDTHS["Calibri"][11]["width"] / self::DEFAULT_COLUMN_WIDTHS["Calibri"][11]["px"] / $fontSize);
    }
}