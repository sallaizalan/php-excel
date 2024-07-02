# php-excel
PHP Library to only write and only XLSX files, in a fast and scalable way, box/spout package served as a basis.

## Installation

Open a command console, enter your project directory and execute the
following command to download the latest stable version of this bundle:

```console
$ composer require threshold/php-excel
```

This command requires you to have Composer installed globally, as explained
in the [installation chapter](https://getcomposer.org/doc/00-intro.md)
of the Composer documentation.

## Usage

Simple XLSX file generation with instant download:

```php
$writer = new \Threshold\PhpExcel\Writer\Writer();
$writer
    ->setShouldCreateNewSheetsAutomatically(false)
    ->setShouldUseInlineStrings(true)
    ->openToBrowser("test.xlsx");

$sheet = $writer->getCurrentSheet();
$sheet
    ->setName("Statistics")
    ->setMergeRanges(["A1:B1"]);

$writer->addRow(\Threshold\PhpExcel\Writer\Factory\WriterEntityFactory::createRow([
    \Threshold\PhpExcel\Writer\Factory\WriterEntityFactory::createCell(
        "Statistics",
        (new \Threshold\PhpExcel\Writer\Entity\Style\Style())
            ->setCellAlignment(\Threshold\PhpExcel\Writer\Entity\Style\CellAlignment::CENTER)
            ->setFontBold())
]));

$writer->close();
```

Generating and saving a simple XLSX file to a given location:

```php
$writer = new \Threshold\PhpExcel\Writer\Writer();
$writer
    ->setShouldCreateNewSheetsAutomatically(false)
    ->setShouldUseInlineStrings(true)
    ->openToFile("path/to/save/file/test.xlsx");

$sheet = $writer->getCurrentSheet();
$sheet
    ->setName("Statistics")
    ->setMergeRanges(["A1:B1"]);

$writer->addRow(\Threshold\PhpExcel\Writer\Factory\WriterEntityFactory::createRow([
    \Threshold\PhpExcel\Writer\Factory\WriterEntityFactory::createCell(
        "Statistics",
        (new \Threshold\PhpExcel\Writer\Entity\Style\Style())
            ->setCellAlignment(\Threshold\PhpExcel\Writer\Entity\Style\CellAlignment::CENTER)
            ->setFontBold())
]));

$writer->close();
```

Generate XLSX recursively, with only one file returned at the end:

```php
// In each round, you must state that you are in the last round and how many rounds there will be in total.
$loopMax     = $request->query->get("loopMax", 1);
$loopCounter = $request->query->get("loopCounter", 1);
$writer      = new \Threshold\PhpExcel\Writer\Writer();
$writer
    ->setTempFolder("your/temp/folder/path/if/you/want")
    ->setShouldCreateNewSheetsAutomatically(false)
    ->setShouldUseInlineStrings(true)
    ->setDefaultRowStyle((new \Threshold\PhpExcel\Writer\Entity\Style\Style())
        ->setFontName("Calibri")
        ->setFontSize(11))
    ->setLoop($loopMax, $loopCounter); // Here you have to pass the circles,
                                       // from this it will know whether the writing will continue
                                       // or the file must be closed.
    
// $writer->close() until the last round returns the name of the temporary folder
// created in the first round in the temporary path,
// this must be specified here in order to continue generating the XLSX file.
if ($request->query->get("tempFolderName")) {
    $writer
        ->setTempFolderName($request->query->get("tempFolderName"));
}

$writer->openToFile("php://output")

$sheet = $writer->getCurrentSheet();
$sheet
    ->setName("Statistics");
    
if ($loopCounter === 1) {
    // The merged cells are counted from the last written row in each round,
    // so for example the cell merging of the address only needs to be entered in the first round.
    // In later rounds, the "A1:B1" cell merge will not apply to the first row, for example,
    // if we wrote 15 rows in the previous round, then cells A and B of the 16th row will be merged.
    // In this case, "A1:B1" will be as if we entered "A16:B16" in the first round.
    $sheet->setMergeRanges(["A1:B1"]);
    
    $writer->addRow(\Threshold\PhpExcel\Writer\Factory\WriterEntityFactory::createRow([
        \Threshold\PhpExcel\Writer\Factory\WriterEntityFactory::createCell("Statistics", (new \Threshold\PhpExcel\Writer\Entity\Style\Style())
            ->setCellAlignment(\Threshold\PhpExcel\Writer\Entity\Style\CellAlignment::CENTER)
            ->setBorder((new \Threshold\PhpExcel\Writer\Entity\Style\Border())
                ->setBorderTop(\Threshold\PhpExcel\Writer\Entity\Style\Color::BLACK, \Threshold\PhpExcel\Writer\Entity\Style\Border::WIDTH_MEDIUM, \Threshold\PhpExcel\Writer\Entity\Style\Border::STYLE_SOLID)
                ->setBorderRight(\Threshold\PhpExcel\Writer\Entity\Style\Color::BLACK, \Threshold\PhpExcel\Writer\Entity\Style\Border::WIDTH_MEDIUM, \Threshold\PhpExcel\Writer\Entity\Style\Border::STYLE_SOLID)
                ->setBorderBottom(\Threshold\PhpExcel\Writer\Entity\Style\Color::BLACK, \Threshold\PhpExcel\Writer\Entity\Style\Border::WIDTH_MEDIUM, \Threshold\PhpExcel\Writer\Entity\Style\Border::STYLE_SOLID)
                ->setBorderLeft(\Threshold\PhpExcel\Writer\Entity\Style\Color::BLACK, \Threshold\PhpExcel\Writer\Entity\Style\Border::WIDTH_MEDIUM, \Threshold\PhpExcel\Writer\Entity\Style\Border::STYLE_SOLID))
            ->setFontName("Arial")
            ->setFontBold()
            ->setFontSize(20)),
        \Threshold\PhpExcel\Writer\Factory\WriterEntityFactory::createCell("", (new \Threshold\PhpExcel\Writer\Entity\Style\Style()) // Create empty cells with same Style to have borders
            ->setCellAlignment(\Threshold\PhpExcel\Writer\Entity\Style\CellAlignment::CENTER)
            ->setBorder((new \Threshold\PhpExcel\Writer\Entity\Style\Border())
                ->setBorderTop(\Threshold\PhpExcel\Writer\Entity\Style\Color::BLACK, \Threshold\PhpExcel\Writer\Entity\Style\Border::WIDTH_MEDIUM, \Threshold\PhpExcel\Writer\Entity\Style\Border::STYLE_SOLID)
                ->setBorderRight(\Threshold\PhpExcel\Writer\Entity\Style\Color::BLACK, \Threshold\PhpExcel\Writer\Entity\Style\Border::WIDTH_MEDIUM, \Threshold\PhpExcel\Writer\Entity\Style\Border::STYLE_SOLID)
                ->setBorderBottom(\Threshold\PhpExcel\Writer\Entity\Style\Color::BLACK, \Threshold\PhpExcel\Writer\Entity\Style\Border::WIDTH_MEDIUM, \Threshold\PhpExcel\Writer\Entity\Style\Border::STYLE_SOLID)
                ->setBorderLeft(\Threshold\PhpExcel\Writer\Entity\Style\Color::BLACK, \Threshold\PhpExcel\Writer\Entity\Style\Border::WIDTH_MEDIUM, \Threshold\PhpExcel\Writer\Entity\Style\Border::STYLE_SOLID))
            ->setFontBold()->setFontSize(20))
    ]));
}

// We write this line in each round. In the first round it goes to cell "A2", in the second round to cell "A3" and so on.
$writer->addRow(\Threshold\PhpExcel\Writer\Factory\WriterEntityFactory::createRow([
    \Threshold\PhpExcel\Writer\Factory\WriterEntityFactory::createCell("Round: " . $loopCounter)
]));

if (($writerResponse = $writer->close()) === true) {
    header('Content-Type: application/zip');
    header('Content-Disposition: attachment; filename="test.xlsx"');
} else {
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(["tempFolderName" => $writerResponse]);
}
exit;
```