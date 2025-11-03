<?php
require_once 'vendor/autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

// Create a new PHPWord instance
$phpWord = new PhpWord();

// Add a section
$section = $phpWord->addSection();

// Add some text before the list
$section->addText('Texto antes de la primera lista:');

// Add a first list with default style
$section->addListItem('Primer elemento de la lista 1', 0);
$section->addListItem('Segundo elemento de la lista 1', 0);
$section->addListItem('Tercer elemento de la lista 1', 0);

$section->addTextBreak(2);
$section->addText('Texto entre las listas:');

// Add a second list with the same default style (should get spaceBefore)
$section->addListItem('Primer elemento de la lista 2', 0);
$section->addListItem('Segundo elemento de la lista 2', 0);
$section->addListItem('Tercer elemento de la lista 2', 0);

$section->addTextBreak(2);
$section->addText('Lista multinivel:');

// Add a multilevel list
$multilevelStyleName = 'multilevel';
$phpWord->addNumberingStyle(
    $multilevelStyleName,
    [
        'type' => 'multilevel',
        'levels' => [
            ['format' => 'decimal', 'text' => '%1.', 'left' => 360, 'hanging' => 360, 'tabPos' => 360],
            ['format' => 'lowerLetter', 'text' => '%2.', 'left' => 720, 'hanging' => 360, 'tabPos' => 720],
        ]
    ]
);

$section->addListItem('Primer elemento (nivel 0)', 0, null, $multilevelStyleName);
$section->addListItem('Elemento nivel 1', 1, null, $multilevelStyleName);
$section->addListItem('Segundo elemento (nivel 0)', 0, null, $multilevelStyleName);
$section->addListItem('Elemento nivel 1', 1, null, $multilevelStyleName);

$section->addTextBreak(2);
$section->addText('Lista con addListItemRun:');

// Add list items using ListItemRun
$listItemRun = $section->addListItemRun();
$listItemRun->addText('Primer elemento con ');
$listItemRun->addText('texto en negrita', ['bold' => true]);

$listItemRun = $section->addListItemRun();
$listItemRun->addText('Segundo elemento con ');
$listItemRun->addText('texto en cursiva', ['italic' => true]);

// Save the document
$objWriter = IOFactory::createWriter($phpWord, 'Word2007');
$filename = 'test_list_spacing_result.docx';
$objWriter->save($filename);

echo "Documento creado exitosamente: $filename\n";
echo "El primer elemento de cada lista (nivel 0) debería tener 15pt de espacio antes.\n";
echo "El último elemento de cada lista (nivel 0) debería tener 15pt de espacio después.\n";
