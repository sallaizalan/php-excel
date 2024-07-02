<?php

namespace Threshold\PhpExcel\Writer\Manager;

use Threshold\PhpExcel\Writer\Entity\Style\Style;

class RegisteredStyle
{
    private Style $style;

    private bool $matchingRowStyle;

    public function __construct(Style $style, bool $matchingRowStyle)
    {
        $this->style            = $style;
        $this->matchingRowStyle = $matchingRowStyle;
    }

    public function getStyle(): Style
    {
        return $this->style;
    }

    public function isMatchingRowStyle(): bool
    {
        return $this->matchingRowStyle;
    }
}
