<?php

namespace Threshold\PhpExcel\Writer\Entity;

use DateInterval;
use DateTime;
use Threshold\PhpExcel\Writer\Entity\Style\Style;

class Cell
{
    const TYPE_NUMERIC = 0;
    const TYPE_STRING = 1;
    const TYPE_FORMULA = 2;
    const TYPE_EMPTY = 3;
    const TYPE_BOOLEAN = 4;
    const TYPE_DATE = 5;
    const TYPE_ERROR = 6;

    private $value;
    private ?int $type;
    private Style $style;

    public function __construct($value, ?Style $style = null)
    {
        $this
            ->setValue($value)
            ->setStyle($style);
    }

    public function setValue($value): self
    {
        $this->value = $value;
        $this->type  = $this->detectType($value);
        return $this;
    }
    
    /**
     * @return mixed
     */
    public function getValue()
    {
        return !$this->isError() ? $this->value : null;
    }

    public function getValueEvenIfError()
    {
        return $this->value;
    }

    public function setStyle(?Style $style): self
    {
        $this->style = $style ?: new Style();
        return $this;
    }

    public function getStyle(): Style
    {
        return $this->style;
    }

    public function getType(): ?int
    {
        return $this->type;
    }

    public function setType(int $type): self
    {
        $this->type = $type;
        return $this;
    }

    public function isBoolean(): bool
    {
        return $this->type === self::TYPE_BOOLEAN;
    }

    public function isEmpty(): bool
    {
        return $this->type === self::TYPE_EMPTY;
    }

    public function isNumeric(): bool
    {
        return $this->type === self::TYPE_NUMERIC;
    }

    public function isString(): bool
    {
        return $this->type === self::TYPE_STRING;
    }

    public function isDate(): bool
    {
        return $this->type === self::TYPE_DATE;
    }

    public function isError(): bool
    {
        return $this->type === self::TYPE_ERROR;
    }

    public function __toString(): string
    {
        return (string) $this->getValue();
    }
    
    private function detectType($value): int
    {
        if (gettype($value) === 'boolean') {
            return self::TYPE_BOOLEAN;
        }
        if ($value === null || $value === '') {
            return self::TYPE_EMPTY;
        }
        if (gettype($value) === 'integer' || gettype($value) === 'double') {
            return self::TYPE_NUMERIC;
        }
        if ($value instanceof DateTime || $value instanceof DateInterval) {
            return self::TYPE_DATE;
        }
        if (gettype($value) === 'string') {
            return self::TYPE_STRING;
        }
        
        return self::TYPE_ERROR;
    }
}
