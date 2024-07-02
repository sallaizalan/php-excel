<?php

namespace Threshold\PhpExcel\Writer\Manager\Style;

use Threshold\PhpExcel\Writer\Entity\Style\Style;
use Threshold\PhpExcel\Writer\Exception\InvalidArgumentException;

class StyleMerger
{
    /**
     * @throws InvalidArgumentException
     */
    public function merge(Style $style, Style $baseStyle): Style
    {
        $mergedStyle = clone $style;

        $this->mergeFontStyles($mergedStyle, $style, $baseStyle);
        $this->mergeOtherFontProperties($mergedStyle, $baseStyle);
        $this->mergeCellProperties($mergedStyle, $style, $baseStyle);

        return $mergedStyle;
    }

    private function mergeFontStyles(Style $styleToUpdate, Style $style, Style $baseStyle): void
    {
        if (!$style->hasSetFontBold() && $baseStyle->isFontBold()) {
            $styleToUpdate->setFontBold();
        }
        if (!$style->hasSetFontItalic() && $baseStyle->isFontItalic()) {
            $styleToUpdate->setFontItalic();
        }
        if (!$style->hasSetFontUnderline() && $baseStyle->isFontUnderline()) {
            $styleToUpdate->setFontUnderline();
        }
        if (!$style->hasSetFontStrikethrough() && $baseStyle->isFontStrikethrough()) {
            $styleToUpdate->setFontStrikethrough();
        }
    }

    private function mergeOtherFontProperties(Style $styleToUpdate, Style $baseStyle): void
    {
        if ($baseStyle->getFontSize() !== Style::DEFAULT_FONT_SIZE) {
            $styleToUpdate->setFontSize($baseStyle->getFontSize());
        }
        if ($baseStyle->getFontColor() !== Style::DEFAULT_FONT_COLOR) {
            $styleToUpdate->setFontColor($baseStyle->getFontColor());
        }
        if ($baseStyle->getFontName() !== Style::DEFAULT_FONT_NAME) {
            $styleToUpdate->setFontName($baseStyle->getFontName());
        }
    }
    
    /**
     * @throws InvalidArgumentException
     */
    private function mergeCellProperties(Style $styleToUpdate, Style $style, Style $baseStyle): void
    {
        if (!$style->hasSetWrapText() && $baseStyle->shouldWrapText()) {
            $styleToUpdate->setShouldWrapText();
        }
        if (!$style->hasSetCellAlignment() && $baseStyle->shouldApplyCellAlignment()) {
            $styleToUpdate->setCellAlignment($baseStyle->getCellAlignment());
        }
        if (!$style->getBorder() && $baseStyle->shouldApplyBorder()) {
            $styleToUpdate->setBorder($baseStyle->getBorder());
        }
        if (!$style->getFormat() && $baseStyle->shouldApplyFormat()) {
            $styleToUpdate->setFormat($baseStyle->getFormat());
        }
        if (!$style->shouldApplyBackgroundColor() && $baseStyle->shouldApplyBackgroundColor()) {
            $styleToUpdate->setBackgroundColor($baseStyle->getBackgroundColor());
        }
    }
}
