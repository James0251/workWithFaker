<?php

require_once '/var/www/html/Faker/vendor/fzaninotto/faker/src/autoload.php';
require_once 'vendor/box/spout/src/Spout/Autoloader/autoload.php';

use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Common\Entity\Style\CellAlignment;

$newFilePath = 'createExcelFile/28.12.2019.xlsx';
$writer = WriterEntityFactory::createXLSXWriter();
$cells = CellAlignment::CENTER;
$writer->openToFile($newFilePath);
$faker = Faker\Factory::create('ru_Ru');
for ($i = 0; $i < 1000; $i++) {
    $cells = WriterEntityFactory::createRowFromArray([
        $faker->numberBetween(0, 1005),
        $faker->date('Y-m-d'),
        $faker->text,
        $faker->userName,
        $faker->name,
        $faker->boolean,
        $faker->date('2020-02-01', 'now'),
        $faker->boolean,
        $faker->date('2020-01-01', '2020-01-29'),
        $faker->lastName,
    ]);
    $writer->addRow($cells);
}
$writer->close();
