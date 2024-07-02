<?php

namespace Threshold\PhpExcel\Writer\Entity\Style;

use Threshold\PhpExcel\Writer\Exception\{InvalidNameException, InvalidWidthException, InvalidStyleException};

class Border
{
    const LEFT = 'left';
    const RIGHT = 'right';
    const TOP = 'top';
    const BOTTOM = 'bottom';

    const STYLE_NONE = 'none';
    const STYLE_SOLID = 'solid';
    const STYLE_DASHED = 'dashed';
    const STYLE_DOTTED = 'dotted';
    const STYLE_DOUBLE = 'double';

    const WIDTH_THIN = 'thin';
    const WIDTH_MEDIUM = 'medium';
    const WIDTH_THICK = 'thick';
    
    private array $parts = [];

    public function __construct(array $borderParts = [])
    {
        $this->setParts($borderParts);
    }

    public function getPart(string $name): ?BorderPart
    {
        return $this->parts[$name] ?? null;
    }

    public function hasPart(string $name): bool
    {
        return isset($this->parts[$name]);
    }

    public function getParts(): array
    {
        return $this->parts;
    }

    public function setParts(array $parts): self
    {
        unset($this->parts);
        foreach ($parts as $part) {
            $this->addPart($part);
        }
        return $this;
    }
    
    public function addPart(BorderPart $borderPart): self
    {
        $this->parts[$borderPart->getName()] = $borderPart;
        ksort($this->parts);

        return $this;
    }
    
    /**
     * @throws InvalidNameException|InvalidWidthException|InvalidStyleException
     */
    public function setBorderTop(string $color = Color::BLACK, string $width = self::WIDTH_MEDIUM,
                                 string $style = self::STYLE_SOLID): self
    {
        $this->addPart(new BorderPart(self::TOP, $color, $width, $style));
        return $this;
    }
    
    /**
     * @throws InvalidNameException|InvalidWidthException|InvalidStyleException
     */
    public function setBorderRight(string $color = Color::BLACK, string $width = self::WIDTH_MEDIUM,
                                   string $style = self::STYLE_SOLID): self
    {
        $this->addPart(new BorderPart(self::RIGHT, $color, $width, $style));
        return $this;
    }
    
    /**
     * @throws InvalidNameException|InvalidWidthException|InvalidStyleException
     */
    public function setBorderBottom(string $color = Color::BLACK, string $width = self::WIDTH_MEDIUM,
                                    string $style = self::STYLE_SOLID): self
    {
        $this->addPart(new BorderPart(self::BOTTOM, $color, $width, $style));
        return $this;
    }
    
    /**
     * @throws InvalidNameException|InvalidWidthException|InvalidStyleException
     */
    public function setBorderLeft(string $color = Color::BLACK, string $width = self::WIDTH_MEDIUM,
                                  string $style = self::STYLE_SOLID): self
    {
        $this->addPart(new BorderPart(self::LEFT, $color, $width, $style));
        return $this;
    }
}
