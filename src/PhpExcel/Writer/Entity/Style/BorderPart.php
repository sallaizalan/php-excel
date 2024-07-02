<?php

namespace Threshold\PhpExcel\Writer\Entity\Style;

use Threshold\PhpExcel\Writer\Exception\{InvalidNameException, InvalidStyleException, InvalidWidthException};

class BorderPart
{
    private string $style;
    private string $name;
    private string $color;
    private string $width;
    private static array $allowedStyles = [
        'none',
        'solid',
        'dashed',
        'dotted',
        'double',
    ];
    private static array $allowedNames = [
        'left',
        'right',
        'top',
        'bottom',
    ];
    private static array $allowedWidths = [
        'thin',
        'medium',
        'thick',
    ];

    /**
     * @throws InvalidNameException|InvalidStyleException|InvalidWidthException
     */
    public function __construct(string $name, string $color = Color::BLACK, string $width = Border::WIDTH_MEDIUM,
                                string $style = Border::STYLE_SOLID)
    {
        $this->setName($name);
        $this->setColor($color);
        $this->setWidth($width);
        $this->setStyle($style);
    }

    public function getName(): string
    {
        return $this->name;
    }

    /**
     * @throws InvalidNameException
     */
    public function setName(string $name): self
    {
        if (!in_array($name, self::$allowedNames)) {
            throw new InvalidNameException($name);
        }
        $this->name = $name;
        return $this;
    }

    public function getStyle(): string
    {
        return $this->style;
    }

    /**
     * @throws InvalidStyleException
     */
    public function setStyle(string $style): self
    {
        if (!in_array($style, self::$allowedStyles)) {
            throw new InvalidStyleException($style);
        }
        $this->style = $style;
        return $this;
    }

    public function getColor(): string
    {
        return $this->color;
    }

    public function setColor(string $color): self
    {
        $this->color = $color;
        return $this;
    }

    public function getWidth(): string
    {
        return $this->width;
    }

    /**
     * @throws InvalidWidthException
     */
    public function setWidth(string $width): self
    {
        if (!in_array($width, self::$allowedWidths)) {
            throw new InvalidWidthException($width);
        }
        $this->width = $width;
        return $this;
    }

    public static function getAllowedStyles(): array
    {
        return self::$allowedStyles;
    }

    public static function getAllowedNames(): array
    {
        return self::$allowedNames;
    }

    public static function getAllowedWidths(): array
    {
        return self::$allowedWidths;
    }
}
