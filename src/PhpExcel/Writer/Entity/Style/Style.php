<?php

namespace Threshold\PhpExcel\Writer\Entity\Style;

use Threshold\PhpExcel\Writer\Exception\InvalidArgumentException;

class Style
{
    const DEFAULT_FONT_SIZE = 11;
    const DEFAULT_FONT_COLOR = Color::BLACK;
    const DEFAULT_FONT_NAME = 'Calibri';

    private ?int $id = null;
    private ?int $fontId = null;
    private bool $fontBold = false;
    private bool $hasSetFontBold = false;
    
    private bool $fontItalic = false;
    private bool $hasSetFontItalic = false;
    
    private bool $fontUnderline = false;
    private bool $hasSetFontUnderline = false;
    
    private bool $fontStrikethrough = false;
    private bool $hasSetFontStrikethrough = false;
    
    private int $fontSize = self::DEFAULT_FONT_SIZE;
    
    private string $fontColor = self::DEFAULT_FONT_COLOR;
    
    private string $fontName = self::DEFAULT_FONT_NAME;

    private bool $shouldApplyCellAlignment = false;
    private ?string $cellVerticalAlignment = null;
    private ?string $cellHorizontalAlignment = null;
    private ?string $cellAlignment = null;
    private bool $hasSetCellAlignment = false;

    private bool $shouldWrapText = false;
    private bool $hasSetWrapText = false;

    private ?Border $border = null;
    private bool $shouldApplyBorder = false;

    private ?string $backgroundColor = null;
    private bool $hasSetBackgroundColor = false;

    private ?string $format = null;
    private bool $hasSetFormat = false;

    private bool $registered = false;

    private bool $empty = true;
    private bool $fromXML = false;

    public function getId(): ?int
    {
        return $this->id;
    }
    
    public function setId(?int $id): self
    {
        $this->id = $id;
        return $this;
    }
    
    public function getFontId(): ?int
    {
        return $this->fontId;
    }
    
    public function setFontId(?int $fontId): self
    {
        $this->fontId = $fontId;
        return $this;
    }

    public function getBorder(): ?Border
    {
        return $this->border;
    }

    public function setBorder(Border $border): self
    {
        $this->shouldApplyBorder = true;
        $this->border            = $border;
        $this->empty             = false;
        return $this;
    }

    public function shouldApplyBorder(): bool
    {
        return $this->shouldApplyBorder;
    }

    public function isFontBold(): bool
    {
        return $this->fontBold;
    }

    public function setFontBold(): self
    {
        $this->fontBold        = true;
        $this->hasSetFontBold  = true;
        $this->empty           = false;
        return $this;
    }

    public function hasSetFontBold(): bool
    {
        return $this->hasSetFontBold;
    }

    public function isFontItalic(): bool
    {
        return $this->fontItalic;
    }

    public function setFontItalic(): self
    {
        $this->fontItalic       = true;
        $this->hasSetFontItalic = true;
        $this->empty            = false;
        return $this;
    }

    public function hasSetFontItalic(): bool
    {
        return $this->hasSetFontItalic;
    }

    public function isFontUnderline(): bool
    {
        return $this->fontUnderline;
    }

    public function setFontUnderline(): self
    {
        $this->fontUnderline       = true;
        $this->hasSetFontUnderline = true;
        $this->empty               = false;
        return $this;
    }

    public function hasSetFontUnderline(): bool
    {
        return $this->hasSetFontUnderline;
    }

    public function isFontStrikethrough(): bool
    {
        return $this->fontStrikethrough;
    }

    public function setFontStrikethrough(): self
    {
        $this->fontStrikethrough       = true;
        $this->hasSetFontStrikethrough = true;
        $this->empty                   = false;
        return $this;
    }

    public function hasSetFontStrikethrough(): bool
    {
        return $this->hasSetFontStrikethrough;
    }

    public function getFontSize(): int
    {
        return $this->fontSize;
    }

    public function setFontSize(int $fontSize): self
    {
        $this->fontSize = $fontSize;
        $this->empty    = false;
        return $this;
    }

    public function getFontColor(): string
    {
        return $this->fontColor;
    }

    public function setFontColor(string $fontColor): self
    {
        $this->fontColor = $fontColor;
        $this->empty     = false;
        return $this;
    }

    public function getFontName(): string
    {
        return $this->fontName;
    }

    public function setFontName(string $fontName): self
    {
        $this->fontName = $fontName;
        $this->empty    = false;
        return $this;
    }
    
    public function getCellVerticalAlignment(): ?string
    {
        return $this->cellVerticalAlignment;
    }
    
    /**
     * @throws InvalidArgumentException
     */
    public function setCellVerticalAlignment(string $cellVerticalAlignment): self
    {
        if (!CellAlignment::isValid($cellVerticalAlignment)) {
            throw new InvalidArgumentException('Invalid cell alignment value');
        }
        
        $this->cellVerticalAlignment    = $cellVerticalAlignment;
        $this->hasSetCellAlignment      = true;
        $this->shouldApplyCellAlignment = true;
        $this->empty                    = false;
        return $this;
    }
    
    public function getCellHorizontalAlignment(): ?string
    {
        return $this->cellHorizontalAlignment;
    }
    
    /**
     * @throws InvalidArgumentException
     */
    public function setCellHorizontalAlignment(string $cellHorizontalAlignment): self
    {
        if (!CellAlignment::isValid($cellHorizontalAlignment)) {
            throw new InvalidArgumentException('Invalid cell alignment value');
        }
        
        $this->cellHorizontalAlignment  = $cellHorizontalAlignment;
        $this->hasSetCellAlignment      = true;
        $this->shouldApplyCellAlignment = true;
        $this->empty                    = false;
        return $this;
    }

    public function getCellAlignment(): ?string
    {
        return $this->cellAlignment;
    }
    
    /**
     * @throws InvalidArgumentException
     */
    public function setCellAlignment(string $cellAlignment): self
    {
        if (!CellAlignment::isValid($cellAlignment)) {
            throw new InvalidArgumentException('Invalid cell alignment value');
        }
        
        $this->cellAlignment            = $cellAlignment;
        $this->hasSetCellAlignment      = true;
        $this->shouldApplyCellAlignment = true;
        $this->empty                    = false;
        return $this;
    }

    public function hasSetCellAlignment(): bool
    {
        return $this->hasSetCellAlignment;
    }

    public function shouldApplyCellAlignment(): bool
    {
        return $this->shouldApplyCellAlignment;
    }

    public function shouldWrapText(): bool
    {
        return $this->shouldWrapText;
    }

    public function setShouldWrapText(bool $shouldWrap = true): self
    {
        $this->shouldWrapText = $shouldWrap;
        $this->hasSetWrapText = true;
        $this->empty          = false;
        return $this;
    }

    public function hasSetWrapText(): bool
    {
        return $this->hasSetWrapText;
    }

    public function setBackgroundColor(string $color): self
    {
        $this->hasSetBackgroundColor = true;
        $this->backgroundColor       = $color;
        $this->empty                 = false;
        return $this;
    }

    public function getBackgroundColor(): ?string
    {
        return $this->backgroundColor;
    }

    public function shouldApplyBackgroundColor(): bool
    {
        return $this->hasSetBackgroundColor;
    }

    public function setFormat(string $format): self
    {
        $this->hasSetFormat = true;
        $this->format       = $format;
        $this->empty        = false;
        return $this;
    }

    public function getFormat(): ?string
    {
        return $this->format;
    }

    public function shouldApplyFormat(): bool
    {
        return $this->hasSetFormat;
    }

    public function isRegistered(): bool
    {
        return $this->registered;
    }

    public function markAsRegistered(?int $id): void
    {
        $this->setId($id);
        $this->registered = true;
        
        if ($this->fontId === null) {
            $this->fontId = $id;
        }
    }

    public function unmarkAsRegistered(): void
    {
        $this->setId(0);
        $this->registered = false;
        
        if ($this->fontId === null) {
            $this->fontId = 0;
        }
    }

    public function isEmpty(): bool
    {
        return $this->empty;
    }
    
    public function isFromXML(): bool
    {
        return $this->fromXML;
    }
    
    public function setFromXML(bool $fromXML): self
    {
        $this->fromXML = $fromXML;
        return $this;
    }
}