<?php

require_once '/var/www/html/Faker/vendor/fzaninotto/faker/src/autoload.php';
require_once 'vendor/box/spout/src/Spout/Autoloader/autoload.php';

use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Common\Entity\Style\CellAlignment;
use Box\Spout\Common\Entity\Style\Color;

$newFilePath = 'createExcelFile/28.12.2019.xlsx';
$writer = WriterEntityFactory::createXLSXWriter();
$cells = CellAlignment::CENTER;
$writer->openToFile($newFilePath);

$faker = Faker\Factory::create('ru_Ru');
for ($i = 0; $i < 1000; $i++) {
    $cells = [
        WriterEntityFactory::createCell($faker->numberBetween(0, 1005)),
        WriterEntityFactory::createCell($faker->date('Y-m-d')),
        WriterEntityFactory::createCell($faker->text),
        WriterEntityFactory::createCell($faker->userName),
        WriterEntityFactory::createCell($faker->name),
        WriterEntityFactory::createCell($faker->boolean),
        WriterEntityFactory::createCell($faker->date('2020-02-01', 'now')),
        WriterEntityFactory::createCell($faker->boolean),
        WriterEntityFactory::createCell($faker->date('2020-01-01', '2020-01-29')),
        WriterEntityFactory::createCell($faker->lastName),
    ];
    $singleRow = WriterEntityFactory::createRow($cells);
    $writer->addRow($singleRow);
}
$writer->close();
