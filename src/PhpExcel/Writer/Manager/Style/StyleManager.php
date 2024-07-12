<?php

namespace Threshold\PhpExcel\Writer\Manager\Style;

use SimpleXMLElement;
use Threshold\PhpExcel\Writer\Entity\{Cell, Style\Border, Style\Color, Style\Style};
use Threshold\PhpExcel\Writer\Helper\BorderHelper;

class StyleManager
{
    const STYLES_XML_FILE_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>';
    
    protected StyleRegistry $styleRegistry;
    private ?SimpleXMLElement $xml = null;
    
    public function __construct(StyleRegistry $styleRegistry)
    {
        $this->styleRegistry = $styleRegistry;
    }
    
    protected function getDefaultStyle(): Style
    {
        // By construction, the default style has ID 0
        return $this->styleRegistry->getRegisteredStyles()[0];
    }
    
    public function registerStyle(Style $style): Style
    {
        return $this->styleRegistry->registerStyle($style);
    }
    
    public function applyExtraStylesIfNeeded(Cell $cell): PossiblyUpdatedStyle
    {
        return $this->applyWrapTextIfCellContainsNewLine($cell);
    }
    
    protected function applyWrapTextIfCellContainsNewLine(Cell $cell): PossiblyUpdatedStyle
    {
        $cellStyle = $cell->getStyle();
        
        // if the "wrap text" option is already set, no-op
        if (!$cellStyle->hasSetWrapText() && $cell->isString() && strpos($cell->getValue(), "\n") !== false) {
            $cellStyle->setShouldWrapText();
            
            return new PossiblyUpdatedStyle($cellStyle, true);
        }
        
        return new PossiblyUpdatedStyle($cellStyle, false);
    }

    public function shouldApplyStyleOnEmptyCell(int $styleId): bool
    {
        $associatedFillId   = $this->styleRegistry->getFillIdForStyleId($styleId);
        $hasStyleCustomFill = ($associatedFillId !== null && $associatedFillId !== 0);

        $associatedBorderId    = $this->styleRegistry->getBorderIdForStyleId($styleId);
        $hasStyleCustomBorders = ($associatedBorderId !== null && $associatedBorderId !== 0);

        $associatedFormatId    = $this->styleRegistry->getFormatIdForStyleId($styleId);
        $hasStyleCustomFormats = ($associatedFormatId !== null && $associatedFormatId !== 0);

        return ($hasStyleCustomFill || $hasStyleCustomBorders || $hasStyleCustomFormats);
    }
    
    public function createStylesXMLFile(string $stylesXMLFilePath): void
    {
        if (file_exists($stylesXMLFilePath)) {
            $this->xml = simplexml_load_file($stylesXMLFilePath);
        }
        else {
            $this->xml = new SimpleXMLElement(self::STYLES_XML_FILE_HEADER);
        }
        $this->xml->registerXPathNamespace("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        
        $this->addFormatSectionToXML();
        $this->addFontsSectionToXML();
        $this->addFillsSectionToXML();
        $this->addBordersSectionToXML();
        $this->addCellStyleXfsSectionToXML();
        $this->addCellXfsSectionToXML();
        $this->addCellStylesSectionToXML();
        
        $this->xml->saveXML($stylesXMLFilePath);
    }
    
    private function addFormatSectionToXML(): void
    {
        $numFmts = $this->xml->numFmts ?? $this->xml->addChild("numFmts");
        
        foreach ($this->styleRegistry->getRegisteredFormats() as $styleId) {
            $numFmt   = null;
            $numFmtId = $this->styleRegistry->getFormatIdForStyleId($styleId);
            
            //Built-in formats do not need to be declared, skip them
            if ($numFmtId < 164) {
                continue;
            }
            
            foreach ($numFmts->children() as $xmlElement) {
                $numFmtIdAttr = $xmlElement->attributes()["numFmtId"] ?? null;
                
                if ($numFmtIdAttr && (int)$numFmtIdAttr === $numFmtId) {
                    $numFmt = $xmlElement;
                    unset($numFmtIdAttr);
                    break;
                }
                unset($numFmtIdAttr);
            }
            
            if (!$numFmt) {
                $numFmt = $numFmts->addChild("numFmt");
            }
            
            $style  = $this->styleRegistry->getStyleFromStyleId($styleId);
            $numFmt->addAttribute("numFmtId", $numFmtId);
            $numFmt->addAttribute("formatCode", $style->getFormat());
            unset($numFmtId);
            unset($style);
            unset($numFmt);
        }
        
        if (isset($numFmts["count"])) {
            $numFmts["count"] = $numFmts->children()->count();
        }
        else {
            $numFmts->addAttribute("count", $numFmts->children()->count());
        }
        unset($numFmts);
    }
    
    private function addFontsSectionToXML(): void
    {
        $fonts = $this->xml->fonts ?? $this->xml->addChild("fonts");
        
        foreach ($this->styleRegistry->getRegisteredStyles() as $style) {
            $alreadyAddedFont = false;
            
            foreach ($this->xml->xpath("xmlns:fonts/xmlns:font") as $xmlFontIndex => $xmlFontElement) {
                if (isset($xmlFontElement->sz["val"]) && (int)$xmlFontElement->sz["val"] === $style->getFontSize()
                    && isset($xmlFontElement->color["rgb"]) && (string)$xmlFontElement->color["rgb"] === Color::toARGB($style->getFontColor())
                    && isset($xmlFontElement->name["val"]) && (string)$xmlFontElement->name["val"] === $style->getFontName()
                    && (((isset($xmlFontElement->b) && $style->isFontBold())
                            || (isset($xmlFontElement->i) && $style->isFontItalic())
                            || (isset($xmlFontElement->u) && $style->isFontUnderline())
                            || (isset($xmlFontElement->strike) && $style->isFontStrikethrough()))
                        || (!isset($xmlFontElement->b) && !isset($xmlFontElement->i) && !isset($xmlFontElement->u) && !isset($xmlFontElement->strike)
                            && !$style->isFontBold() && !$style->isFontItalic() && !$style->isFontUnderline() && !$style->isFontStrikethrough())))
                {
                    $alreadyAddedFont = true;
                    $style->setFontId($xmlFontIndex);
                    break;
                }
            }
            
            if (!$alreadyAddedFont) {
                $style->setFontId($fonts->children()->count());
                
                $font     = $fonts->addChild("font");
                $fontSize = $font->addChild("sz");
                $fontSize->addAttribute("val", $style->getFontSize());
                
                $fontColor = $font->addChild("color");
                $fontColor->addAttribute("rgb", Color::toARGB($style->getFontColor()));
                
                $fontName = $font->addChild("name");
                $fontName->addAttribute("val", $style->getFontName());
                
                if ($style->isFontBold()) {
                    $font->addChild("b");
                } elseif ($style->isFontItalic()) {
                    $font->addChild("i");
                } elseif ($style->isFontUnderline()) {
                    $font->addChild("u");
                } elseif ($style->isFontStrikethrough()) {
                    $font->addChild("strike");
                }
                
                unset($font);
                unset($fontSize);
                unset($fontColor);
                unset($fontName);
            }
            unset($alreadyAddedFont);
        }
        
        if (isset($fonts["count"])) {
            $fonts["count"] = $fonts->children()->count();
        }
        else {
            $fonts->addAttribute("count", $fonts->children()->count());
        }
        
        unset($fonts);
    }
    
    private function addFillsSectionToXML(): void
    {
        $fills = $this->xml->fills ?? $this->xml->addChild("fills");
        
        $alreadyAddedFillNoneType    = false;
        $alreadyAddedFillGray125Type = false;
        
        foreach ($this->xml->xpath("xmlns:fills/xmlns:fill/xmlns:patternFill") as $xmlPatternFillElement) {
            $patternTypeAttr = $xmlPatternFillElement["patternType"] ?? null;
            
            if ($patternTypeAttr && (string)$patternTypeAttr === "none") {
                $alreadyAddedFillNoneType = true;
            }
            elseif ($patternTypeAttr && (string)$patternTypeAttr === "gray125") {
                $alreadyAddedFillGray125Type = true;
            }
            
            if ($alreadyAddedFillNoneType && $alreadyAddedFillGray125Type) {
                break;
            }
        }
        
        if (!$alreadyAddedFillNoneType) {
            $fill        = $fills->addChild("fill");
            $patternFill = $fill->addChild("patternFill");
            $patternFill->addAttribute("patternType", "none");
            unset($fill);
            unset($patternFill);
        }
        
        
        if (!$alreadyAddedFillGray125Type) {
            $fill        = $fills->addChild("fill");
            $patternFill = $fill->addChild("patternFill");
            $patternFill->addAttribute("patternType", "gray125");
            unset($fill);
            unset($patternFill);
        }
        unset($alreadyAddedFillNoneType);
        unset($alreadyAddedFillGray125Type);
        
        foreach ($this->styleRegistry->getRegisteredFills() as $styleId) {
            $alreadyAddedFill = false;
            $style            = $this->styleRegistry->getStyleFromStyleId($styleId);
            
            foreach ($this->xml->xpath("xmlns:fills/xmlns:fill/xmlns:patternFill") as $xmlIndex => $xmlPatternFillElement) {
                $patternTypeAttr = $xmlPatternFillElement["patternType"] ?? null;

                if ($patternTypeAttr && (string)$patternTypeAttr === "solid"
                    && isset($xmlPatternFillElement->fgColor["rgb"]) && (string)$xmlPatternFillElement->fgColor["rgb"] === Color::toARGB($style->getBackgroundColor()))
                {
                    $alreadyAddedFill = true;

                    if ($xmlIndex !== $styleId) {
                        $this->styleRegistry->modifyFillMappingTable($styleId, $xmlIndex);
                    }
                    break;
                }
            }
            
            if (!$alreadyAddedFill) {
                $this->styleRegistry->modifyFillMappingTable($styleId, $fills->children()->count());
                
                $fill        = $fills->addChild("fill");
                $patternFill = $fill->addChild("patternFill");
                $fgColor     = $patternFill->addChild("fgColor");
                $patternFill->addAttribute("patternType", "solid");
                $fgColor->addAttribute("rgb", Color::toARGB($style->getBackgroundColor()));
                unset($fill);
                unset($patternFill);
                unset($fgColor);
            }
            unset($alreadyAddedFill);
            unset($style);
        }
        
        if (isset($fills["count"])) {
            $fills["count"] = $fills->children()->count();
        }
        else {
            $fills->addAttribute("count", $fills->children()->count());
        }
        
        unset($fills);
    }
    
    private function addBordersSectionToXML(): void
    {
        $borders = $this->xml->borders ?? $this->xml->addChild("borders");
        
        if ($borders->count() === 0) {
            $border = $borders->addChild("border");
            $border->addChild("left");
            $border->addChild("right");
            $border->addChild("top");
            $border->addChild("bottom");
            unset($border);
        }
        
        foreach ($this->styleRegistry->getRegisteredBorders() as $styleId) {
            $alreadyAddedBorder = false;
            $style              = $this->styleRegistry->getStyleFromStyleId($styleId);
            
            foreach ($this->xml->xpath("xmlns:borders/xmlns:border") as $xmlIndex => $xmlBorderElement) {
                if ($xmlBorderElement->children()->count() === count($style->getBorder()->getParts())
                    && ((!isset($xmlBorderElement->top) && !$style->getBorder()->hasPart(Border::TOP))
                        || (isset($xmlBorderElement->top) && $style->getBorder()->hasPart(Border::TOP)
                            && isset($xmlBorderElement->top["style"]) && (string)$xmlBorderElement->top["style"] === BorderHelper::getBorderStyle($style->getBorder()->getPart(Border::TOP))))
                    && ((!isset($xmlBorderElement->right) && !$style->getBorder()->hasPart(Border::RIGHT))
                        || (isset($xmlBorderElement->right) && $style->getBorder()->hasPart(Border::RIGHT)
                            && isset($xmlBorderElement->right["style"]) && (string)$xmlBorderElement->right["style"] === BorderHelper::getBorderStyle($style->getBorder()->getPart(Border::RIGHT))))
                    && ((!isset($xmlBorderElement->bottom) && !$style->getBorder()->hasPart(Border::BOTTOM))
                        || (isset($xmlBorderElement->bottom) && $style->getBorder()->hasPart(Border::BOTTOM)
                            && isset($xmlBorderElement->bottom["style"]) && (string)$xmlBorderElement->bottom["style"] === BorderHelper::getBorderStyle($style->getBorder()->getPart(Border::BOTTOM))))
                    && ((!isset($xmlBorderElement->left) && !$style->getBorder()->hasPart(Border::LEFT))
                        || (isset($xmlBorderElement->left) && $style->getBorder()->hasPart(Border::LEFT)
                            && isset($xmlBorderElement->left["style"]) && (string)$xmlBorderElement->left["style"] === BorderHelper::getBorderStyle($style->getBorder()->getPart(Border::LEFT)))))
                {
                    $alreadyAddedBorder = true;
                    
                    if ($xmlIndex !== $styleId) {
                        $this->styleRegistry->modifyBorderMappingTable($styleId, $xmlIndex);
                    }
                    break;
                }
            }
            
            if (!$alreadyAddedBorder) {
                $this->styleRegistry->modifyBorderMappingTable($styleId, $borders->children()->count());
                
                $border = $borders->addChild("border");
                
                foreach (['left', 'right', 'top', 'bottom'] as $partName) {
                    if ($part = $style->getBorder()->getPart($partName)) {
                        $partXML = $border->addChild($partName);
                        $partXML->addAttribute("style", BorderHelper::getBorderStyle($part));
                        
                        if ($part->getColor()) {
                            $color = $partXML->addChild("color");
                            $color->addAttribute("rgb", $part->getColor());
                            unset($color);
                        }
                        unset($partXML);
                    }
                    unset($part);
                }
                unset($border);
            }
            unset($alreadyAddedBorder);
            unset($style);
        }
        
        if (isset($borders["count"])) {
            $borders["count"] = $borders->children()->count();
        }
        else {
            $borders->addAttribute("count", $borders->children()->count());
        }
        
        unset($borders);
    }
    
    private function addCellStyleXfsSectionToXML(): void
    {
        if (!$this->xml->cellStyleXfs) {
            $cellStyleXfs = $this->xml->addChild("cellStyleXfs");
            $cellStyleXfs->addAttribute("count", "1");
            $xf = $cellStyleXfs->addChild("xf");
            $xf->addAttribute("borderId", "0");
            $xf->addAttribute("fillId", "0");
            $xf->addAttribute("fontId", "0");
            $xf->addAttribute("numFmtId", "0");
            unset($cellStyleXfs);
            unset($xf);
        }
    }
    
    private function addCellXfsSectionToXML(): void
    {
        $cellXfs = $this->xml->cellXfs ?? $this->xml->addChild("cellXfs");
        
        foreach ($this->styleRegistry->getRegisteredStyles() as $style) {
            $alreadyAddedXf = false;
            $numFmtId       = $style->getId() !== 0 ? ($this->styleRegistry->getFormatIdForStyleId($style->getId()) ?: 0) : $style->getId();
            $fontId         = $style->getFontId();
            $fillId         = $style->getId() !== 0 ? ($this->styleRegistry->getFillIdForStyleId($style->getId()) ?: 0) : $style->getId();
            $borderId       = $style->getId() !== 0 ? ($this->styleRegistry->getBorderIdForStyleId($style->getId()) ?: 0) : $style->getId();
            
            foreach ($this->xml->xpath("xmlns:cellXfs/xmlns:xf") as $xmlXfElement) {
                if (isset($xmlXfElement["numFmtId"]) && (int)$xmlXfElement["numFmtId"] === (int)$numFmtId
                    && isset($xmlXfElement["fontId"]) && (int)$xmlXfElement["fontId"] === (int)$fontId
                    && isset($xmlXfElement["fillId"]) && (int)$xmlXfElement["fillId"] === (int)$fillId
                    && isset($xmlXfElement["borderId"]) && (int)$xmlXfElement["borderId"] === (int)$borderId
//                    && (((!isset($xmlXfElement["applyFont"]) || (string)$xmlXfElement["applyFont"] === "0") && !$style->shouldApplyFont())
//                        || ((string)$xmlXfElement["applyFont"] === "1" && $style->shouldApplyFont()))
                    && (((!isset($xmlXfElement["applyBorder"]) || (string)$xmlXfElement["applyBorder"] === "0") && !$style->shouldApplyBorder())
                        || ((string)$xmlXfElement["applyBorder"] === "1" && $style->shouldApplyBorder()))
                    && ((!isset($xmlXfElement["applyAlignment"]) && !isset($xmlXfElement->alignment) && !$style->shouldApplyCellAlignment() && !$style->shouldWrapText())
                        || (isset($xmlXfElement["applyAlignment"]) && (string)$xmlXfElement["applyAlignment"] === "1" && $style->shouldApplyCellAlignment() && !$style->shouldWrapText()
                            && (($style->getCellAlignment() !== null && isset($xmlXfElement->alignment["horizontal"]) && (string)$xmlXfElement->alignment["horizontal"] === $style->getCellAlignment() && isset($xmlXfElement->alignment["vertical"]) && (string)$xmlXfElement->alignment["vertical"] === $style->getCellAlignment())
                                || ($style->getCellAlignment() === null && (!isset($xmlXfElement->alignment["vertical"]) || !isset($xmlXfElement->alignment["horizontal"])) && (
                                        ($style->getCellVerticalAlignment() !== null && isset($xmlXfElement->alignment["vertical"]) && (string)$xmlXfElement->alignment["vertical"] === $style->getCellVerticalAlignment())
                                        || ($style->getCellHorizontalAlignment() !== null && isset($xmlXfElement->alignment["horizontal"]) && (string)$xmlXfElement->alignment["horizontal"] === $style->getCellHorizontalAlignment()))))
                        )
                        || (isset($xmlXfElement["applyAlignment"]) && (string)$xmlXfElement["applyAlignment"] === "1" && $style->shouldWrapText() && !$style->shouldApplyCellAlignment()
                            && isset($xmlXfElement->alignment["wrapText"]) && (string)$xmlXfElement->alignment["wrapText"] === "1")
                        || (isset($xmlXfElement["applyAlignment"]) && (string)$xmlXfElement["applyAlignment"] === "1" && $style->shouldApplyCellAlignment()
                            && (($style->getCellAlignment() !== null && isset($xmlXfElement->alignment["horizontal"]) && (string)$xmlXfElement->alignment["horizontal"] === $style->getCellAlignment() && isset($xmlXfElement->alignment["vertical"]) && (string)$xmlXfElement->alignment["vertical"] === $style->getCellAlignment())
                                || ($style->getCellAlignment() === null && (!isset($xmlXfElement->alignment["vertical"]) || !isset($xmlXfElement->alignment["horizontal"])) && (
                                        ($style->getCellVerticalAlignment() !== null && isset($xmlXfElement->alignment["vertical"]) && (string)$xmlXfElement->alignment["vertical"] === $style->getCellVerticalAlignment())
                                        || ($style->getCellHorizontalAlignment() !== null && isset($xmlXfElement->alignment["horizontal"]) && (string)$xmlXfElement->alignment["horizontal"] === $style->getCellHorizontalAlignment()))))
                            && isset($xmlXfElement["applyAlignment"]) && (string)$xmlXfElement["applyAlignment"] === "1" && $style->shouldWrapText() && isset($xmlXfElement->alignment["wrapText"]) && (string)$xmlXfElement->alignment["wrapText"] === "1"))
                ) {
                    $alreadyAddedXf = true;
                    break;
                }
            }
            
            if (!$alreadyAddedXf) {
                $xf = $cellXfs->addChild("xf");
                $xf->addAttribute("numFmtId", $numFmtId);
                $xf->addAttribute("fontId", $fontId);
                $xf->addAttribute("fillId", $fillId);
                $xf->addAttribute("borderId", $borderId);
                $xf->addAttribute("xfId", "0");
                $xf->addAttribute("applyFont", "1");
                $xf->addAttribute("applyBorder", $style->shouldApplyBorder() ? "1" : "0");
                
                if ($style->shouldApplyCellAlignment() || $style->shouldWrapText()) {
                    $xf->addAttribute("applyAlignment", "1");
                    $alignment = $xf->addChild("alignment");
                    
                    if ($style->shouldApplyCellAlignment()) {
                        if ($style->getCellAlignment() !== null) {
                            $alignment->addAttribute("horizontal", $style->getCellAlignment());
                            $alignment->addAttribute("vertical", $style->getCellAlignment());
                        }
                        else {
                            if ($style->getCellVerticalAlignment() !== null) {
                                $alignment->addAttribute("vertical", $style->getCellVerticalAlignment());
                            }
                            if ($style->getCellHorizontalAlignment() !== null) {
                                $alignment->addAttribute("horizontal", $style->getCellHorizontalAlignment());
                            }
                        }
                    }
                    if ($style->shouldWrapText()) {
                        $alignment->addAttribute("wrapText", "1");
                    }
                    unset($alignment);
                }
                
                unset($xf);
            }
            unset($alreadyAddedXf);
            unset($numFmtId);
            unset($fontId);
            unset($fillId);
            unset($borderId);
        }
        
        if (isset($cellXfs["count"])) {
            $cellXfs["count"] = $cellXfs->children()->count();
        }
        else {
            $cellXfs->addAttribute("count", $cellXfs->children()->count());
        }
        
        unset($cellXfs);
    }
    
    private function addCellStylesSectionToXML(): void
    {
        if (!$this->xml->cellStyles) {
            $cellStyles = $this->xml->addChild("cellStyles");
            $cellStyles->addAttribute("count", 1);
            $cellStyle = $cellStyles->addChild("cellStyle");
            $cellStyle->addAttribute("builtinId", 0);
            $cellStyle->addAttribute("name", "normal");
            $cellStyle->addAttribute("xfId", 0);
            unset($cellStyle);
            unset($cellStyles);
        }
    }
}
