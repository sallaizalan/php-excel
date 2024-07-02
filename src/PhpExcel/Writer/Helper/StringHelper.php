<?php

namespace Threshold\PhpExcel\Writer\Helper;

class StringHelper
{
    protected bool $hasMbStringSupport;
    private bool $isRunningPhp7OrOlder;
    private array $localeInfo;

    public function __construct()
    {
        $this->hasMbStringSupport   = extension_loaded('mbstring');
        $this->isRunningPhp7OrOlder = version_compare(PHP_VERSION, '8.0.0') < 0;
        $this->localeInfo           = localeconv();
    }

    public function getStringLength(string $string): int
    {
        return $this->hasMbStringSupport ? mb_strwidth($string) : strlen($string);
    }

    public function getCharFirstOccurrencePosition(string $char, string $string): int
    {
        $position = $this->hasMbStringSupport ? mb_strpos($string, $char) : strpos($string, $char);

        return ($position !== false) ? $position : -1;
    }

    public function getCharLastOccurrencePosition(string $char, string $string): int
    {
        $position = $this->hasMbStringSupport ? mb_strrpos($string, $char) : strrpos($string, $char);

        return ($position !== false) ? $position : -1;
    }

    /**
     * @param int|float $numericValue
     */
    public function formatNumericValue($numericValue): string
    {
        if ($this->isRunningPhp7OrOlder && is_float($numericValue)) {
            return str_replace(
                [$this->localeInfo['thousands_sep'], $this->localeInfo['decimal_point']],
                ['', '.'],
                $numericValue
            );
        }

        return $numericValue;
    }
}
