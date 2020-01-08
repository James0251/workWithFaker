<?php
require_once 'vendor/autoload.php';

//Используем namespace для reader
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
//Используем namespace для writer
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;

//Путь до готового шаблона, который лежит в папке excelFiles
$existingFilePath = 'createExcelFile/28.12.2019.xlsx';
//Здесь будет лежать созданный файл, когда мы его заполним и сохраним
$newFilePath = 'createExcelFile/new.xlsx';

//Создаем reader, чтобы он прочитал наш файл 1.xlsx с информацией
$reader = ReaderEntityFactory::createReaderFromFile($existingFilePath);
//Открываем наш файл 1.xlsx
$reader->open($existingFilePath);
//Эта строка делается для того, чтобы иметь возможность копировать даты
$reader->setShouldFormatDates(true);

//Пишем writer для создания нового файла
$writer = WriterEntityFactory::createWriterFromFile($newFilePath);
//Он будет открывать наш созданный файл с данными new.xlsx
$writer->openToFile($newFilePath);

//Прочитаем всю таблицу целиком
foreach ($reader->getSheetIterator() as $sheetIndex => $sheet) {
    if ($sheetIndex != 1) {
        $writer->addNewSheetAndMakeItCurrent();
    }
    foreach ($sheet->getRowIterator() as  $rowIndex => $row) {
        $songTitle = $row->getCellAtIndex(0);
        $artist = $row->getCellAtIndex(1);


        // Выбираем альбом "Yellow Submarine"
        if ($songTitle == 'Yellow Submarine') {
            $row->setCellAtIndex(WriterEntityFactory::createCell('The White Album'), 2);
        }

        //Пропускаем песни Боба Марли
        if ($artist == 'Bob Marley') {
            continue;
        }

        //Запишем отредактированную строку в файл
        $writer->addRow($row);

        //Вставим новую песню в нужное место, между 3-м и 4-м рядами
        if ($rowIndex == 7) {
            $writer->addRow(
                WriterEntityFactory::createRowFromArray(['James', 'The Best', 'PHP-programmer', 2020])
            );
        }
    }
}
$reader->close();
$writer->close();


