<?php

namespace Threshold\PhpExcel\Writer\Manager\Style;

use Threshold\PhpExcel\Writer\Entity\Style\{Border, BorderPart, Color, Style};
use Threshold\PhpExcel\Writer\Exception\{InvalidArgumentException, InvalidStyleException, InvalidWidthException};
use Threshold\PhpExcel\Writer\Helper\BorderHelper;

/**
 * Class StyleRegistry
 * Registry for all used styles
 */
class StyleRegistry
{
    private array $serializedStyleToStyleIdMappingTable = [];
    private array $styleIdToStyleMappingTable = [];
    /*
     * @see https://msdn.microsoft.com/en-us/library/ff529597(v=office.12).aspx
     */
    private static array $builtinNumFormatToIdMapping = [
        'General' => 0,
        '0' => 1,
        '0.00' => 2,
        '#,##0' => 3,
        '#,##0.00' => 4,
        '$#,##0,\-$#,##0' => 5,
        '$#,##0,[Red]\-$#,##0' => 6,
        '$#,##0.00,\-$#,##0.00' => 7,
        '$#,##0.00,[Red]\-$#,##0.00' => 8,
        '0%' => 9,
        '0.00%' => 10,
        '0.00E+00' => 11,
        '# ?/?' => 12,
        '# ??/??' => 13,
        'mm-dd-yy' => 14,
        'd-mmm-yy' => 15,
        'd-mmm' => 16,
        'mmm-yy' => 17,
        'h:mm AM/PM' => 18,
        'h:mm:ss AM/PM' => 19,
        'h:mm' => 20,
        'h:mm:ss' => 21,
        'm/d/yy h:mm' => 22,

        '#,##0 ,(#,##0)' => 37,
        '#,##0 ,[Red](#,##0)' => 38,
        '#,##0.00,(#,##0.00)' => 39,
        '#,##0.00,[Red](#,##0.00)' => 40,

        '_("$"* #,##0.00_),_("$"* \(#,##0.00\),_("$"* "-"??_),_(@_)' => 44,
        'mm:ss' => 45,
        '[h]:mm:ss' => 46,
        'mm:ss.0' => 47,

        '##0.0E+0' => 48,
        '@' => 49,

        '[$-404]e/m/d' => 27,
        'm/d/yy' => 30,
        't0' => 59,
        't0.00' => 60,
        't#,##0' => 61,
        't#,##0.00' => 62,
        't0%' => 67,
        't0.00%' => 68,
        't# ?/?' => 69,
        't# ??/??' => 70,
    ];
    private array $registeredFormats = [];
    private array $styleIdToFormatMappingTable = [];
    private int $formatIndex = 164;
    private array $registeredFills = [];
    private array $styleIdToFillMappingTable = [];
    private int $fillIndex = 2;
    private array $registeredBorders = [];
    private array $styleIdToBorderMappingTable = [];
    
    /**
     * @throws InvalidArgumentException|InvalidStyleException|InvalidWidthException
     */
    public function __construct(Style $defaultStyle, string $stylesXMLFilePath)
    {
        // This ensures that the default style is the first one to be registered
        if (file_exists($stylesXMLFilePath)) {
            $this->loadStylesAndRegisterFromXML($stylesXMLFilePath);
        }
        else {
            $this->registerStyle($defaultStyle);
        }
    }

    public function registerStyle(Style $style): Style
    {
        if ($style->isRegistered()) {
            return $style;
        }
        
        $serializedStyle = $this->serialize($style);
        if (!isset($this->serializedStyleToStyleIdMappingTable[$serializedStyle])) {
            $nextStyleId = count($this->serializedStyleToStyleIdMappingTable);
            $style->markAsRegistered($nextStyleId);
            
            $this->serializedStyleToStyleIdMappingTable[$serializedStyle] = $nextStyleId;
            $this->styleIdToStyleMappingTable[$nextStyleId]               = $style;
            
            $registeredStyle = $this->getStyleFromSerializedStyle($serializedStyle);
            $this->registerFill($registeredStyle);
            $this->registerFormat($registeredStyle);
            $this->registerBorder($registeredStyle);
        }
        else {
            $registeredStyle = $this->getStyleFromSerializedStyle($serializedStyle);
        }

        return $registeredStyle;
    }
    
    public function getRegisteredStyles(): array
    {
        return array_values($this->styleIdToStyleMappingTable);
    }
    
    public function getBorderIdForStyleId(int $styleId): ?int
    {
        return $this->styleIdToBorderMappingTable[$styleId] ?? null;
    }
    
    public function getRegisteredFills(): array
    {
        return $this->registeredFills;
    }
    
    public function getRegisteredBorders(): array
    {
        return $this->registeredBorders;
    }
    
    public function getRegisteredFormats(): array
    {
        return $this->registeredFormats;
    }
    
    public function getStyleFromStyleId(int $styleId): Style
    {
        return $this->styleIdToStyleMappingTable[$styleId];
    }
    
    public function getFormatIdForStyleId(int $styleId): ?int
    {
        return $this->styleIdToFormatMappingTable[$styleId] ?? null;
    }
    
    public function serialize(Style $style): string
    {
        // In order to be able to properly compare style, set static ID value and reset registration
        $currentId        = $style->getId();
        $currentFontId    = $style->getFontId();
        $currentIsFromXML = $style->isFromXML();
        
        $style->setFontId(0);
        $style->setFromXML(false);
        $style->unmarkAsRegistered();
        
        $serializedStyle = serialize($style);
        
        $style->markAsRegistered($currentId);
        $style->setFontId($currentFontId);
        $style->setFromXML($currentIsFromXML);
        
        return $serializedStyle;
    }
    
    public function getFillIdForStyleId(int $styleId): ?int
    {
        return $this->styleIdToFillMappingTable[$styleId] ?? null;
    }
    
    public function registerFormat(Style $style): void
    {
        $styleId = $style->getId();
        $format  = $style->getFormat();
        
        if ($format) {
            $isFormatRegistered = isset($this->registeredFormats[$format]);
            
            // We need to track the already registered format definitions
            if ($isFormatRegistered) {
                $registeredStyleId                           = $this->registeredFormats[$format];
                $registeredFormatId                          = $this->styleIdToFormatMappingTable[$registeredStyleId];
                $this->styleIdToFormatMappingTable[$styleId] = $registeredFormatId;
            }
            else {
                $this->registeredFormats[$format]            = $styleId;
                $this->styleIdToFormatMappingTable[$styleId] = self::$builtinNumFormatToIdMapping[$format] ?? $this->formatIndex++;
            }
        }
        else {
            // The formatId maps a style to a format declaration
            // When there is no format definition - we default to 0 ( General )
            $this->styleIdToFormatMappingTable[$styleId] = 0;
        }
    }
    
    private function getStyleFromSerializedStyle(string $serializedStyle): Style
    {
        $styleId = $this->serializedStyleToStyleIdMappingTable[$serializedStyle];
        
        return $this->styleIdToStyleMappingTable[$styleId];
    }

    private function registerFill(Style $style): void
    {
        $styleId = $style->getId();

        // Currently - only solid backgrounds are supported
        // so $backgroundColor is a scalar value (RGB Color)
        $backgroundColor = $style->getBackgroundColor();

        if ($backgroundColor) {
            $isBackgroundColorRegistered = isset($this->registeredFills[$backgroundColor]);

            // We need to track the already registered background definitions
            if ($isBackgroundColorRegistered) {
                $registeredStyleId                         = $this->registeredFills[$backgroundColor];
                $registeredFillId                          = $this->styleIdToFillMappingTable[$registeredStyleId];
                $this->styleIdToFillMappingTable[$styleId] = $registeredFillId;
            }
            else {
                $this->registeredFills[$backgroundColor]   = $styleId;
                $this->styleIdToFillMappingTable[$styleId] = $this->fillIndex++;
            }
        }
        else {
            // The fillId maps a style to a fill declaration
            // When there is no background color definition - we default to 0
            $this->styleIdToFillMappingTable[$styleId] = 0;
        }
    }
    
    public function modifyFillMappingTable(int $styleId, int $index): void
    {
        $this->styleIdToFillMappingTable[$styleId] = $index;
    }

    private function registerBorder(Style $style): void
    {
        $styleId = $style->getId();

        if ($style->shouldApplyBorder()) {
            $border           = $style->getBorder();
            $serializedBorder = serialize($border);

            $isBorderAlreadyRegistered = isset($this->registeredBorders[$serializedBorder]);

            if ($isBorderAlreadyRegistered) {
                $registeredStyleId                           = $this->registeredBorders[$serializedBorder];
                $registeredBorderId                          = $this->styleIdToBorderMappingTable[$registeredStyleId];
                $this->styleIdToBorderMappingTable[$styleId] = $registeredBorderId;
            }
            else {
                $this->registeredBorders[$serializedBorder]  = $styleId;
                $this->styleIdToBorderMappingTable[$styleId] = count($this->registeredBorders);
            }
        }
        else {
            // If no border should be applied - the mapping is the default border: 0
            $this->styleIdToBorderMappingTable[$styleId] = 0;
        }
    }
    
    public function modifyBorderMappingTable(int $styleId, int $index): void
    {
        $this->styleIdToBorderMappingTable[$styleId] = $index;
    }
    
    /**
     * @throws InvalidArgumentException|InvalidStyleException|InvalidWidthException
     */
    private function loadStylesAndRegisterFromXML(string $stylesXMLFilePath): void
    {
        $xml = simplexml_load_file($stylesXMLFilePath);
        $xml->registerXPathNamespace("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        
        foreach ($xml->xpath("xmlns:cellXfs/xmlns:xf") as $xmlIndex => $xmlXfElement) {
            $style = (new Style())->setId($xmlIndex)->setFromXml(true);
            
            if (isset($xmlXfElement["fontId"])) {
                foreach ($xml->xpath("xmlns:fonts/xmlns:font") as $xmlFontIndex => $xmlFontElement) {
                    if ((int)$xmlFontIndex === (int)$xmlXfElement["fontId"]) {
                        $style->setFontId($xmlFontIndex);
                        
                        if (isset($xmlFontElement->sz["val"])) {
                            $style->setFontSize((int)$xmlFontElement->sz["val"]);
                        }
                        if (isset($xmlFontElement->color["rgb"])) {
                            $style->setFontColor(Color::toRGB((string)$xmlFontElement->color["rgb"]));
                        }
                        if (isset($xmlFontElement->name["val"])) {
                            $style->setFontName((string)$xmlFontElement->name["val"]);
                        }
                        if (isset($xmlFontElement->b)) {
                            $style->setFontBold();
                        }
                        if (isset($xmlFontElement->i)) {
                            $style->setFontItalic();
                        }
                        if (isset($xmlFontElement->u)) {
                            $style->setFontUnderline();
                        }
                        if (isset($xmlFontElement->strike)) {
                            $style->setFontStrikethrough();
                        }
                    }
                }
            }
            if (isset($xmlXfElement["fillId"]) && (int)$xmlXfElement["fillId"] > 1) {
                foreach ($xml->xpath("xmlns:fills/xmlns:fill") as $xmlFillIndex => $xmlFillElement) {
                    if ((int)$xmlFillIndex === (int)$xmlXfElement["fillId"]) {
                        if (isset($xmlFillElement->patternFill->fgColor["rgb"])) {
                            $style->setBackgroundColor(Color::toRGB((string)$xmlFillElement->patternFill->fgColor["rgb"]));
                        }
                    }
                }
            }
            if (isset($xmlXfElement["borderId"]) && (int)$xmlXfElement["borderId"] > 0) {
                foreach ($xml->xpath("xmlns:borders/xmlns:border") as $xmlBorderIndex => $xmlBorderElement) {
                    if ((int)$xmlBorderIndex === (int)$xmlXfElement["borderId"]) {
                        $borderParts = [];
                        if (isset($xmlBorderElement->left["style"])) {
                            $borderParts["left"] = (new BorderPart(Border::LEFT));
                            BorderHelper::setBorderPartStyleAndWidthFromXMLStyle($borderParts["left"], $xmlBorderElement->left["style"]);
                            
                            if (isset($xmlBorderElement->left->color["rgb"])) {
                                $borderParts["left"]->setColor(Color::toRGB((string)$xmlBorderElement->left->color["rgb"]));
                            }
                        }
                        if (isset($xmlBorderElement->right["style"])) {
                            $borderParts["right"] = (new BorderPart(Border::RIGHT));
                            BorderHelper::setBorderPartStyleAndWidthFromXMLStyle($borderParts["right"], $xmlBorderElement->right["style"]);
                            
                            if (isset($xmlBorderElement->right->color["rgb"])) {
                                $borderParts["right"]->setColor(Color::toRGB((string)$xmlBorderElement->right->color["rgb"]));
                            }
                        }
                        if (isset($xmlBorderElement->top["style"])) {
                            $borderParts["top"] = (new BorderPart(Border::TOP));
                            BorderHelper::setBorderPartStyleAndWidthFromXMLStyle($borderParts["top"], $xmlBorderElement->top["style"]);
                            
                            if (isset($xmlBorderElement->top->color["rgb"])) {
                                $borderParts["top"]->setColor(Color::toRGB((string)$xmlBorderElement->top->color["rgb"]));
                            }
                        }
                        if (isset($xmlBorderElement->bottom["style"])) {
                            $borderParts["bottom"] = (new BorderPart(Border::BOTTOM));
                            BorderHelper::setBorderPartStyleAndWidthFromXMLStyle($borderParts["bottom"], $xmlBorderElement->bottom["style"]);
                            
                            if (isset($xmlBorderElement->bottom->color["rgb"])) {
                                $borderParts["bottom"]->setColor(Color::toRGB((string)$xmlBorderElement->bottom->color["rgb"]));
                            }
                        }
                        $style->setBorder(new Border($borderParts));
                    }
                }
            }
            if (isset($xmlXfElement["applyAlignment"]) && (int)$xmlXfElement["applyAlignment"] === 1 && isset($xmlXfElement->alignment)) {
                if (isset($xmlXfElement->alignment["wrapText"]) && (int)$xmlXfElement->alignment["wrapText"] === 1) {
                    $style->setShouldWrapText();
                }
                
                $horizontal = isset($xmlXfElement->alignment["horizontal"]) ? (string)$xmlXfElement->alignment["horizontal"] : null;
                $vertical   = isset($xmlXfElement->alignment["vertical"]) ? (string)$xmlXfElement->alignment["vertical"] : null;
                
                if ($horizontal && $vertical && $horizontal === $vertical) {
                    $style->setCellAlignment($horizontal);
                }
                elseif ($horizontal) {
                    $style->setCellHorizontalAlignment($horizontal);
                }
                elseif ($vertical) {
                    $style->setCellVerticalAlignment($vertical);
                }
            }
            
            $this->registerStyle($style);
        }
    }
}