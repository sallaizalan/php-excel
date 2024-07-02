<?php

namespace Threshold\PhpExcel\Writer\Entity\Style;

use Threshold\PhpExcel\Writer\Exception\InvalidColorException;

abstract class Color
{
    /* Standard colors - based on Office Online */
    const BLACK = '000000';
    const WHITE = 'FFFFFF';
    const RED = 'FF0000';
    const DARK_RED = 'C00000';
    const ORANGE = 'FFC000';
    const YELLOW = 'FFFF00';
    const LIGHT_GREEN = '92D040';
    const GREEN = '00B050';
    const LIGHT_BLUE = '00B0E0';
    const BLUE = '0070C0';
    const DARK_BLUE = '002060';
    const PURPLE = '7030A0';
    
    /**
     * @throws InvalidColorException
     */
    public static function rgb(int $red, int $green, int $blue): string
    {
        self::throwIfInvalidColorComponentValue($red);
        self::throwIfInvalidColorComponentValue($green);
        self::throwIfInvalidColorComponentValue($blue);

        return strtoupper(
            self::convertColorComponentToHex($red) .
            self::convertColorComponentToHex($green) .
            self::convertColorComponentToHex($blue)
        );
    }

    /**
     * @throws InvalidColorException
     */
    private static function throwIfInvalidColorComponentValue(int $colorComponent): void
    {
        if ($colorComponent < 0 || $colorComponent > 255) {
            throw new InvalidColorException("The RGB components must be between 0 and 255. Received: $colorComponent");
        }
    }

    private static function convertColorComponentToHex(int $colorComponent): string
    {
        return str_pad(dechex($colorComponent), 2, '0', STR_PAD_LEFT);
    }

    public static function toARGB(string $rgbColor): string
    {
        return strlen($rgbColor) === 8 ? $rgbColor : 'FF' . $rgbColor;
    }

    public static function toRGB(string $argbColor): string
    {
        return strlen($argbColor) === 8 ? substr($argbColor, 2) : $argbColor;
    }
}
