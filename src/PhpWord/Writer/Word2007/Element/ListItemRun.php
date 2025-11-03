<?php

/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @see         https://github.com/PHPOffice/PHPWord
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Writer\Word2007\Element;

use PhpOffice\PhpWord\Element\ListItemRun as ListItemRunElement;
use PhpOffice\PhpWord\Writer\Word2007\Style\Paragraph as ParagraphStyleWriter;
use PhpOffice\PhpWord\Writer\Word2007\Style\Spacing as SpacingStyleWriter;

/**
 * ListItemRun element writer.
 *
 * @since 0.10.0
 */
class ListItemRun extends AbstractElement
{
    /**
     * Track list boundaries (first and last items).
     *
     * @var array
     */
    private static $listBoundaries = [];

    /**
     * Analyze container elements to find list boundaries.
     *
     * @param array $elements
     */
    private static function analyzeListBoundaries($elements): void
    {
        $listSequences = [];
        $currentList = null;

        foreach ($elements as $index => $element) {
            if (
                $element instanceof \PhpOffice\PhpWord\Element\ListItem ||
                $element instanceof \PhpOffice\PhpWord\Element\ListItemRun
            ) {

                $numId = $element->getStyle()->getNumId();
                $depth = $element->getDepth();

                // Only track depth 0 items
                if ($depth === 0) {
                    $listKey = 'list_' . $numId;

                    if ($currentList !== $listKey) {
                        // End previous list
                        if ($currentList !== null && isset($listSequences[$currentList])) {
                            $listSequences[$currentList]['last'] = $listSequences[$currentList]['current'];
                        }

                        // Start new list
                        $currentList = $listKey;
                        if (!isset($listSequences[$currentList])) {
                            $listSequences[$currentList] = [
                                'first' => $index,
                                'current' => $index,
                                'last' => null
                            ];
                        }
                    }

                    $listSequences[$currentList]['current'] = $index;
                }
            } else {
                // Non-list element breaks the sequence
                if ($currentList !== null && isset($listSequences[$currentList])) {
                    $listSequences[$currentList]['last'] = $listSequences[$currentList]['current'];
                    $currentList = null;
                }
            }
        }

        // Close last list
        if ($currentList !== null && isset($listSequences[$currentList])) {
            $listSequences[$currentList]['last'] = $listSequences[$currentList]['current'];
        }

        self::$listBoundaries = $listSequences;
    }

    /**
     * Check if analysis is needed.
     *
     * @param \PhpOffice\PhpWord\Element\ListItemRun $element
     */
    private static function ensureBoundariesAnalyzed($element): void
    {
        if (empty(self::$listBoundaries)) {
            $parent = $element->getParent();
            if ($parent !== null && $parent instanceof \PhpOffice\PhpWord\Element\AbstractContainer) {
                $elements = $parent->getElements();
                self::analyzeListBoundaries($elements);
            }
        }
    }

    /**
     * Write list item element.
     */
    public function write(): void
    {
        $element = $this->getElement();

        if (!$element instanceof ListItemRunElement) {
            return;
        }

        $this->writeParagraph($element);
    }

    private function writeParagraph(ListItemRunElement $element): void
    {
        $xmlWriter = $this->getXmlWriter();
        $xmlWriter->startElement('w:p');

        $this->writeParagraphProperties($element);

        $containerWriter = new Container($xmlWriter, $element);
        $containerWriter->write();

        $xmlWriter->endElement(); // w:p
    }

    private function writeParagraphProperties(ListItemRunElement $element): void
    {
        self::ensureBoundariesAnalyzed($element);

        $xmlWriter = $this->getXmlWriter();
        $xmlWriter->startElement('w:pPr');

        $styleWriter = new ParagraphStyleWriter($xmlWriter, $element->getParagraphStyle());
        $styleWriter->setIsInline(true);
        $styleWriter->setWithoutPPR(true);
        $styleWriter->write();

        $numId = $element->getStyle()->getNumId();
        $depth = $element->getDepth();
        $elementIndex = $element->getElementIndex() - 1; // Adjust for 0-based array
        $listKey = 'list_' . $numId;

        // Add spacing for first/last items at depth 0
        if ($depth === 0 && isset(self::$listBoundaries[$listKey])) {
            $isFirst = ($elementIndex === self::$listBoundaries[$listKey]['first']);
            $isLast = ($elementIndex === self::$listBoundaries[$listKey]['last']);

            if ($isFirst || $isLast) {
                $spacingConfig = [];
                if ($isFirst) {
                    $spacingConfig['before'] = 300; // 15pt
                }
                if ($isLast) {
                    $spacingConfig['after'] = 300; // 15pt
                }

                $spacingStyle = new \PhpOffice\PhpWord\Style\Spacing($spacingConfig);
                $spacingWriter = new SpacingStyleWriter($xmlWriter, $spacingStyle);
                $spacingWriter->write();
            }
        }

        $this->writeParagraphPropertiesNumbering($element);

        $xmlWriter->endElement(); // w:pPr
    }

    private function writeParagraphPropertiesNumbering(ListItemRunElement $element): void
    {
        $xmlWriter = $this->getXmlWriter();
        $xmlWriter->startElement('w:numPr');

        $xmlWriter->writeElementBlock('w:ilvl', [
            'w:val' => $element->getDepth(),
        ]);

        $xmlWriter->writeElementBlock('w:numId', [
            'w:val' => $element->getStyle()->getNumId(),
        ]);

        $xmlWriter->endElement(); // w:numPr
    }
}
