<?php

namespace Threshold\PhpExcel\Writer\Helper;

use Threshold\PhpExcel\Writer\Entity\Style\{Border, BorderPart};
use Threshold\PhpExcel\Writer\Exception\{InvalidStyleException, InvalidWidthException};

class BorderHelper
{
    private static array $xlsxStyleMap = [
        Border::STYLE_SOLID => [
            Border::WIDTH_THIN   => 'thin',
            Border::WIDTH_MEDIUM => 'medium',
            Border::WIDTH_THICK  => 'thick',
        ],
        Border::STYLE_DOTTED => [
            Border::WIDTH_THIN   => 'dotted',
            Border::WIDTH_MEDIUM => 'dotted',
            Border::WIDTH_THICK  => 'dotted',
        ],
        Border::STYLE_DASHED => [
            Border::WIDTH_THIN   => 'dashed',
            Border::WIDTH_MEDIUM => 'mediumDashed',
            Border::WIDTH_THICK  => 'mediumDashed',
        ],
        Border::STYLE_DOUBLE => [
            Border::WIDTH_THIN   => 'double',
            Border::WIDTH_MEDIUM => 'double',
            Border::WIDTH_THICK  => 'double',
        ],
        Border::STYLE_NONE => [
            Border::WIDTH_THIN   => 'none',
            Border::WIDTH_MEDIUM => 'none',
            Border::WIDTH_THICK  => 'none',
        ],
    ];

    public static function serializeBorderPart(BorderPart $borderPart): string
    {
        $borderStyle = self::getBorderStyle($borderPart);

        $colorEl = $borderPart->getColor() ? sprintf('<color rgb="%s"/>', $borderPart->getColor()) : '';
        $partEl = sprintf(
            '<%s style="%s">%s</%s>',
            $borderPart->getName(),
            $borderStyle,
            $colorEl,
            $borderPart->getName()
        );

        return $partEl . PHP_EOL;
    }

    public static function getBorderStyle(BorderPart $borderPart): string
    {
        return self::$xlsxStyleMap[$borderPart->getStyle()][$borderPart->getWidth()];
    }
    
    /**
     * @throws InvalidStyleException|InvalidWidthException
     */
    public static function setBorderPartStyleAndWidthFromXMLStyle(BorderPart $borderPart, string $xmlStyle): void
    {
        foreach (self::$xlsxStyleMap as $style => $widthData) {
            $found = false;
            foreach ($widthData as $width => $xmlKey) {
                if ($xmlKey === $xmlStyle) {
                    $borderPart->setStyle($style)->setWidth($width);
                    $found = true;
                    break;
                }
            }
            if ($found) {
                break;
            }
        }
    }
}
