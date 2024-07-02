<?php

namespace Threshold\PhpExcel\Writer\Helper;

class CellHelper
{
    private static array $columnIndexToColumnLettersCache = [];
    private static array $indexCache = [];

    public static function getColumnLettersFromColumnIndex(int $columnIndexZeroBased): string
    {
        $originalColumnIndex = $columnIndexZeroBased;

        // Using isset here because it is way faster than array_key_exists...
        if (!isset(self::$columnIndexToColumnLettersCache[$originalColumnIndex])) {
            $columnLetters = '';
            $capitalAAsciiValue = \ord('A');

            do {
                $modulus = $columnIndexZeroBased % 26;
                $columnLetters = \chr($capitalAAsciiValue + $modulus) . $columnLetters;

                // substracting 1 because it's zero-based
                $columnIndexZeroBased = (int) ($columnIndexZeroBased / 26) - 1;
            } while ($columnIndexZeroBased >= 0);

            self::$columnIndexToColumnLettersCache[$originalColumnIndex] = $columnLetters;
        }

        return self::$columnIndexToColumnLettersCache[$originalColumnIndex];
    }
    
    public static function getColumnIndexFromString($columnAddress)
    {
        //    Using a lookup cache adds a slight memory overhead, but boosts speed
        //    caching using a static within the method is faster than a class static,
        //        though it's additional memory overhead
        $columnAddress = $columnAddress ?? '';
        
        if (isset($indexCache[$columnAddress])) {
            return $indexCache[$columnAddress];
        }
        //    It's surprising how costly the strtoupper() and ord() calls actually are, so we use a lookup array
        //        rather than use ord() and make it case insensitive to get rid of the strtoupper() as well.
        //        Because it's a static, there's no significant memory overhead either.
        static $columnLookup = [
            'A' => 1, 'B' => 2, 'C' => 3, 'D' => 4, 'E' => 5, 'F' => 6, 'G' => 7, 'H' => 8, 'I' => 9, 'J' => 10,
            'K' => 11, 'L' => 12, 'M' => 13, 'N' => 14, 'O' => 15, 'P' => 16, 'Q' => 17, 'R' => 18, 'S' => 19,
            'T' => 20, 'U' => 21, 'V' => 22, 'W' => 23, 'X' => 24, 'Y' => 25, 'Z' => 26,
            'a' => 1, 'b' => 2, 'c' => 3, 'd' => 4, 'e' => 5, 'f' => 6, 'g' => 7, 'h' => 8, 'i' => 9, 'j' => 10,
            'k' => 11, 'l' => 12, 'm' => 13, 'n' => 14, 'o' => 15, 'p' => 16, 'q' => 17, 'r' => 18, 's' => 19,
            't' => 20, 'u' => 21, 'v' => 22, 'w' => 23, 'x' => 24, 'y' => 25, 'z' => 26,
        ];
        
        //    We also use the language construct isset() rather than the more costly strlen() function to match the
        //       length of $columnAddress for improved performance
        if (isset($columnAddress[0])) {
            if (!isset($columnAddress[1])) {
                $indexCache[$columnAddress] = $columnLookup[$columnAddress];
                
                return $indexCache[$columnAddress];
            }
            elseif (!isset($columnAddress[2])) {
                $indexCache[$columnAddress] = $columnLookup[$columnAddress[0]] * 26
                    + $columnLookup[$columnAddress[1]];
                
                return $indexCache[$columnAddress];
            }
            elseif (!isset($columnAddress[3])) {
                $indexCache[$columnAddress] = $columnLookup[$columnAddress[0]] * 676
                    + $columnLookup[$columnAddress[1]] * 26
                    + $columnLookup[$columnAddress[2]];
                
                return $indexCache[$columnAddress];
            }
        }
        
        return self::$indexCache[$columnAddress];
    }
}
