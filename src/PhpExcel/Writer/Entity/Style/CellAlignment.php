<?php

namespace Threshold\PhpExcel\Writer\Entity\Style;

abstract class CellAlignment
{
    const LEFT = 'left';
    const RIGHT = 'right';
    const CENTER = 'center';
    const JUSTIFY = 'justify';
    const TOP = 'top';
    const BOTTOM = 'bottom';

    private static array $VALID_ALIGNMENTS = [
        self::LEFT    => 1,
        self::RIGHT   => 1,
        self::CENTER  => 1,
        self::JUSTIFY => 1,
        self::TOP     => 1,
        self::BOTTOM  => 1
    ];

    public static function isValid(string $cellAlignment): bool
    {
        return isset(self::$VALID_ALIGNMENTS[$cellAlignment]);
    }
}
