<?php

namespace Threshold\PhpExcel\Writer\Manager;

use Threshold\PhpExcel\Writer\Entity\{Options, Style\Style};

class OptionsManager
{
    private Style $style;
    private array $supportedOptions = [
        Options::TEMP_FOLDER,
        Options::TEMP_FOLDER_NAME,
        Options::LOOP_MAX,
        Options::LOOP_COUNTER,
        Options::DEFAULT_ROW_STYLE,
        Options::SHOULD_CREATE_NEW_SHEETS_AUTOMATICALLY,
        Options::SHOULD_USE_INLINE_STRINGS,
    ];
    private array $options = [];

    public function __construct()
    {
        $this->style = new Style();
        $this->setOption(Options::TEMP_FOLDER, sys_get_temp_dir());
        $this->setOption(Options::DEFAULT_ROW_STYLE, $this->style
            ->setFontSize(Style::DEFAULT_FONT_SIZE)
            ->setFontName(Style::DEFAULT_FONT_NAME)
            ->setFontColor(Style::DEFAULT_FONT_COLOR));
        $this->setOption(Options::SHOULD_CREATE_NEW_SHEETS_AUTOMATICALLY, true);
        $this->setOption(Options::SHOULD_USE_INLINE_STRINGS, true);
    }
    
    public function getFontStyle(): Style
    {
        return $this->getOption(Options::DEFAULT_ROW_STYLE) ?? $this->style
            ->setFontSize(Style::DEFAULT_FONT_SIZE)
            ->setFontName(Style::DEFAULT_FONT_NAME)
            ->setFontColor(Style::DEFAULT_FONT_COLOR);
    }
    
    public function setOption(string $optionName, $optionValue): void
    {
        if (in_array($optionName, $this->supportedOptions)) {
            $this->options[$optionName] = $optionValue;
        }
    }
    
    public function getOption(string $optionName)
    {
        $optionValue = null;
        
        if (isset($this->options[$optionName])) {
            $optionValue = $this->options[$optionName];
        }
        
        return $optionValue;
    }
}
