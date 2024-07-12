<?php

namespace Threshold\PhpExcel\Writer\Helper;

use Exception;
use Threshold\PhpExcel\Writer\Entity\Sheet;

class WorksheetHelper
{
    private const SHEET_NAME_REQUIRES_NO_QUOTES = '/^[_\p{L}][_\p{L}\p{N}]*$/mui';
    private const A1_COORDINATE_REGEX = '/^(?<col>\$?[A-Z]{1,3})(?<row>\$?\d{1,7})$/i';
    const DEFAULT_RANGE = 'A1:A1';
    
    /**
     * @throws Exception
     */
    public static function getExtractCellsReferences(Sheet $sheet): array
    {
        $mergeCells = [];
        foreach ($sheet->getMergeRanges() as $cells) {
            foreach (WorksheetHelper::extractAllCellReferencesInRange($cells) as $cellReference) {
                $mergeCells[$cellReference] = true;
            }
        }
        return $mergeCells;
    }
    
    /**
     * @throws Exception
     */
    public static function extractAllCellReferencesInRange($cellRange): array
    {
        if (substr_count($cellRange, '!') > 1) {
            throw new Exception('3-D Range References are not supported');
        }
        
        [$worksheet, $cellRange] = self::extractSheetTitle($cellRange, true);
        $quoted                  = '';
        
        if ($worksheet > '') {
            $quoted = self::nameRequiresQuotes($worksheet) ? "'" : '';
            if (substr($worksheet, 0, 1) === "'" && substr($worksheet, -1) === "'") {
                $worksheet = substr($worksheet, 1, -1);
            }
            $worksheet = str_replace("'", "''", $worksheet);
        }
        
        [$ranges, $operators] = self::getCellBlocksFromRangeString($cellRange);
        
        $cells = [];
        foreach ($ranges as $range) {
            $cells[] = self::getReferencesForCellBlock($range);
        }
        
        $cells = self::processRangeSetOperators($operators, $cells);
        
        if (empty($cells)) {
            return [];
        }
        
        $cellList = array_merge(...$cells);
        
        return array_map(
            function ($cellAddress) use ($worksheet, $quoted) {
                return ($worksheet !== '') ? "$quoted$worksheet$quoted!$cellAddress" : $cellAddress;
            },
            self::sortCellReferenceArray($cellList)
        );
    }
    
    private static function processRangeSetOperators(array $operators, array $cells): array
    {
        $operatorCount = count($operators);
        for ($offset = 0; $offset < $operatorCount; ++$offset) {
            $operator = $operators[$offset];
            if ($operator !== ' ') {
                continue;
            }
            
            $cells[$offset] = array_intersect($cells[$offset], $cells[$offset + 1]);
            unset($operators[$offset], $cells[$offset + 1]);
            $operators = array_values($operators);
            $cells = array_values($cells);
            --$offset;
            --$operatorCount;
        }
        
        return $cells;
    }
    
    private static function sortCellReferenceArray(array $cellList): array
    {
        //    Sort the result by column and row
        $sortKeys = [];
        foreach ($cellList as $coordinate) {
            $column = '';
            $row    = 0;
            sscanf($coordinate, '%[A-Z]%d', $column, $row);
            $key = (--$row * 16384) + CellHelper::getColumnIndexFromString($column);
            $sortKeys[$key] = $coordinate;
        }
        ksort($sortKeys);
        
        return array_values($sortKeys);
    }
    
    /**
     * @throws Exception
     */
    private static function getReferencesForCellBlock($cellBlock): array
    {
        $returnValue = [];
        
        // Single cell?
        if (!self::coordinateIsRange($cellBlock)) {
            return (array) $cellBlock;
        }
        
        // Range...
        $ranges = self::splitRange($cellBlock);
        foreach ($ranges as $range) {
            // Single cell?
            if (!isset($range[1])) {
                $returnValue[] = $range[0];
                
                continue;
            }
            
            // Range...
            [$rangeStart, $rangeEnd] = $range;
            [$startColumn, $startRow] = self::coordinateFromString($rangeStart);
            [$endColumn, $endRow] = self::coordinateFromString($rangeEnd);
            $startColumnIndex = CellHelper::getColumnIndexFromString($startColumn);
            $endColumnIndex = CellHelper::getColumnIndexFromString($endColumn);
            ++$endColumnIndex;
            
            if ($startColumn === $endColumn && $startRow < $endRow) { // vertical merging does not count in the cell width calculation
                continue;
            }
            
            // Current data
            $currentColumnIndex = $startColumnIndex;
            $currentRow = $startRow;
            
            self::validateRange($cellBlock, $startColumnIndex, $endColumnIndex, (int) $currentRow, (int) $endRow);
            
            // Loop cells
            while ($currentColumnIndex < $endColumnIndex) {
                while ($currentRow <= $endRow) {
                    $returnValue[] = CellHelper::getColumnLettersFromColumnIndex($currentColumnIndex - 1) . $currentRow;
                    ++$currentRow;
                }
                ++$currentColumnIndex;
                $currentRow = $startRow;
            }
        }
        
        return $returnValue;
    }
    
    private static function getCellBlocksFromRangeString($rangeString): array
    {
        $rangeString = str_replace('$', '', strtoupper($rangeString));
        
        // split range sets on intersection (space) or union (,) operators
        $tokens = preg_split('/([ ,])/', $rangeString, -1, PREG_SPLIT_DELIM_CAPTURE);
        /** @phpstan-ignore-next-line */
        $split = array_chunk($tokens, 2);
        $ranges = array_column($split, 0);
        $operators = array_column($split, 1);
        
        return [$ranges, $operators];
    }
    
    /**
     * @throws Exception
     */
    private static function validateRange($cellBlock, $startColumnIndex, $endColumnIndex, $currentRow, $endRow): void
    {
        if ($startColumnIndex >= $endColumnIndex || $currentRow > $endRow) {
            throw new Exception('Invalid range: "' . $cellBlock . '"');
        }
    }
    
    /**
     * @throws Exception
     */
    public static function coordinateFromString($cellAddress): array
    {
        if (preg_match(self::A1_COORDINATE_REGEX, $cellAddress, $matches)) {
            return [$matches['col'], $matches['row']];
        } elseif (self::coordinateIsRange($cellAddress)) {
            throw new Exception('Cell coordinate string can not be a range of cells');
        } elseif ($cellAddress == '') {
            throw new Exception('Cell coordinate can not be zero-length string');
        }
        
        throw new Exception('Invalid cell coordinate ' . $cellAddress);
    }
    
    public static function coordinateIsRange($cellAddress): bool
    {
        return strpos($cellAddress, ":") !== false || strpos($cellAddress, ',') !== false;
    }
    
    public static function splitRange($range): array
    {
        // Ensure $pRange is a valid range
        if (empty($range)) {
            $range = self::DEFAULT_RANGE;
        }
        
        $exploded = explode(',', $range);
        $counter  = count($exploded);
        
        for ($i = 0; $i < $counter; ++$i) {
            // @phpstan-ignore-next-line
            $exploded[$i] = explode(':', $exploded[$i]);
        }
        
        return $exploded;
    }
    
    /**
     * @param $range
     * @param bool $returnInRange
     * @return array|string|null
     */
    public static function extractSheetTitle($range, bool $returnInRange = false)
    {
        if (empty($range)) {
            return $returnInRange ? [null, null] : null;
        }
        
        // Sheet title included?
        if (($sep = strrpos($range, '!')) === false) {
            return $returnInRange ? ['', $range] : '';
        }
        
        if ($returnInRange) {
            return [substr($range, 0, $sep), substr($range, $sep + 1)];
        }
        
        return substr($range, $sep + 1);
    }
    
    public static function nameRequiresQuotes(string $sheetName): bool
    {
        return preg_match(self::SHEET_NAME_REQUIRES_NO_QUOTES, $sheetName) !== 1;
    }
}