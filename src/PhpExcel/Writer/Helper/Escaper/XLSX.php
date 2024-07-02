<?php

namespace Threshold\PhpExcel\Writer\Helper\Escaper;

class XLSX
{
    private bool $isAlreadyInitialized = false;
    private string $escapableControlCharactersPattern;
    private array $controlCharactersEscapingMap;
    private array $controlCharactersEscapingReverseMap;

    protected function initIfNeeded()
    {
        if (!$this->isAlreadyInitialized) {
            $this->escapableControlCharactersPattern   = $this->getEscapableControlCharactersPattern();
            $this->controlCharactersEscapingMap        = $this->getControlCharactersEscapingMap();
            $this->controlCharactersEscapingReverseMap = array_flip($this->controlCharactersEscapingMap);
            $this->isAlreadyInitialized                = true;
        }
    }

    public function escape(string $string): string
    {
        $this->initIfNeeded();

        $escapedString = $this->escapeControlCharacters($string);
        // @NOTE: Using ENT_QUOTES as XML entities ('<', '>', '&') as well as
        //        single/double quotes (for XML attributes) need to be encoded.
        return htmlspecialchars($escapedString, ENT_QUOTES, 'UTF-8');
    }

    public function unescape(string $string): string
    {
        $this->initIfNeeded();

        // ==============
        // =   WARNING  =
        // ==============
        // It is assumed that the given string has already had its XML entities decoded.
        // This is true if the string is coming from a DOMNode (as DOMNode already decode XML entities on creation).
        // Therefore, there is no need to call "htmlspecialchars_decode()".
        return $this->unescapeControlCharacters($string);
    }

    private function getEscapableControlCharactersPattern(): string
    {
        // control characters values are from 0 to 1F (hex values) in the ASCII table
        // some characters should not be escaped though: "\t", "\r" and "\n".
        return '[\x00-\x08' .
                // skipping "\t" (0x9) and "\n" (0xA)
                '\x0B-\x0C' .
                // skipping "\r" (0xD)
                '\x0E-\x1F]';
    }

    private function getControlCharactersEscapingMap(): array
    {
        $controlCharactersEscapingMap = [];

        // control characters values are from 0 to 1F (hex values) in the ASCII table
        for ($charValue = 0x00; $charValue <= 0x1F; $charValue++) {
            $character = chr($charValue);
            if (preg_match("/$this->escapableControlCharactersPattern/", $character)) {
                $charHexValue = dechex($charValue);
                $escapedChar = '_x' . sprintf('%04s', strtoupper($charHexValue)) . '_';
                $controlCharactersEscapingMap[$escapedChar] = $character;
            }
        }

        return $controlCharactersEscapingMap;
    }
    
    private function escapeControlCharacters(string $string): string
    {
        $escapedString = $this->escapeEscapeCharacter($string);

        // if no control characters
        if (!preg_match("/$this->escapableControlCharactersPattern/", $escapedString)) {
            return $escapedString;
        }

        return preg_replace_callback("/($this->escapableControlCharactersPattern)/", function ($matches) {
            return $this->controlCharactersEscapingReverseMap[$matches[0]];
        }, $escapedString);
    }
    
    private function escapeEscapeCharacter(string $string): string
    {
        return preg_replace('/_(x[\dA-F]{4})_/', '_x005F_$1_', $string);
    }
    
    private function unescapeControlCharacters(string $string): string
    {
        $unescapedString = $string;

        foreach ($this->controlCharactersEscapingMap as $escapedCharValue => $charValue) {
            // only unescape characters that don't contain the escaped character for now
            $unescapedString = preg_replace("/(?<!_x005F)($escapedCharValue)/", $charValue, $unescapedString);
        }

        return $this->unescapeEscapeCharacter($unescapedString);
    }
    
    private function unescapeEscapeCharacter(string $string): string
    {
        return preg_replace('/_x005F(_x[\dA-F]{4}_)/', '$1', $string);
    }
}
