<?php

namespace Threshold\PhpExcel\Writer\Entity;

abstract class Options
{
    // Multisheets options
    const TEMP_FOLDER = 'tempFolder';
    const TEMP_FOLDER_NAME = 'tempFolderName';
    const LOOP_MAX = 'loopMax';
    const LOOP_COUNTER = 'loopCounter';
    const DEFAULT_ROW_STYLE = 'defaultRowStyle';
    const SHOULD_CREATE_NEW_SHEETS_AUTOMATICALLY = 'shouldCreateNewSheetsAutomatically';

    // XLSX specific options
    const SHOULD_USE_INLINE_STRINGS = 'shouldUseInlineStrings';
}
