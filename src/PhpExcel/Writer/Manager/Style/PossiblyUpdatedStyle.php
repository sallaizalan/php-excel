<?php

namespace Threshold\PhpExcel\Writer\Manager\Style;

use Threshold\PhpExcel\Writer\Entity\Style\Style;

class PossiblyUpdatedStyle
{
    private Style $style;
    private bool $updated;

    public function __construct(Style $style, bool $updated)
    {
        $this->style   = $style;
        $this->updated = $updated;
    }

    public function getStyle() : Style
    {
        return $this->style;
    }

    public function isUpdated() : bool
    {
        return $this->updated;
    }
}
