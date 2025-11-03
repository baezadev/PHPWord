<?php
require_once 'vendor/autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\SimpleType\JcTable;

// Create a new PHPWord instance
$phpWord = new PhpWord();

// Add a section
$section = $phpWord->addSection();

$section->addText('Tabla con encabezado (alineación vertical centrada por defecto):', ['bold' => true, 'size' => 14]);
$section->addTextBreak();

// Define table style
$tableStyle = [
    'borderSize' => 6,
    'borderColor' => '006699',
    'cellMargin' => 80,
    'alignment' => JcTable::CENTER
];

// Create table
$table = $section->addTable($tableStyle);

// Add header row (tblHeader = true)
$headerRow = $table->addRow(1200, ['tblHeader' => true]);
$headerRow->addCell(2000)->addText('Columna 1', ['bold' => true]);
$headerRow->addCell(2000)->addText('Columna 2', ['bold' => true]);
$headerRow->addCell(2000)->addText('Columna 3', ['bold' => true]);

// Add data rows
for ($i = 1; $i <= 5; $i++) {
    $row = $table->addRow();
    $row->addCell(2000)->addText("Fila $i, Celda 1\nCon múltiples\nlíneas de texto");
    $row->addCell(2000)->addText("Fila $i, Celda 2");
    $row->addCell(2000)->addText("Fila $i, Celda 3\nTexto adicional");
}

$section->addTextBreak(2);
$section->addText('Tabla sin encabezado (alineación normal):', ['bold' => true, 'size' => 14]);
$section->addTextBreak();

// Create table without header
$table2 = $section->addTable($tableStyle);

// Add normal rows
for ($i = 1; $i <= 3; $i++) {
    $row = $table2->addRow();
    $row->addCell(2000)->addText("Fila $i, Celda 1\nSin centrado vertical");
    $row->addCell(2000)->addText("Fila $i, Celda 2");
    $row->addCell(2000)->addText("Fila $i, Celda 3");
}

$section->addTextBreak(2);
$section->addText('Tabla con encabezado y estilo vAlign personalizado:', ['bold' => true, 'size' => 14]);
$section->addTextBreak();

// Create table with custom vAlign in header
$table3 = $section->addTable($tableStyle);

// Add header row with explicit vAlign (should override default)
$headerRow3 = $table3->addRow(1200, ['tblHeader' => true]);
$headerRow3->addCell(2000, ['valign' => 'top'])->addText('Top', ['bold' => true]);
$headerRow3->addCell(2000, ['valign' => 'bottom'])->addText('Bottom', ['bold' => true]);
$headerRow3->addCell(2000)->addText('Default (center)', ['bold' => true]);

// Add data row
$row = $table3->addRow();
$row->addCell(2000)->addText("Celda 1\nCon múltiples\nlíneas");
$row->addCell(2000)->addText("Celda 2\nCon múltiples\nlíneas");
$row->addCell(2000)->addText("Celda 3\nCon múltiples\nlíneas");

// Save the document
$objWriter = IOFactory::createWriter($phpWord, 'Word2007');
$filename = 'test_table_header_valign_result.docx';
$objWriter->save($filename);

echo "Documento creado exitosamente: $filename\n";
echo "Las celdas de encabezado (tblHeader=true) deberían estar centradas verticalmente por defecto.\n";
echo "Las celdas que especifican un vAlign explícito deben usar ese valor en su lugar.\n";
