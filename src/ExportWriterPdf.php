<?php

/**
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2018
 * @package yii2-export
 * @version 1.3.7
 */

namespace kartik\export;

use DOMDocument;
use kartik\mpdf\Pdf;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;

/**
 * Krajee custom PDF Writer library based on MPdf
 *
 * @author Kartik Visweswaran <kartikv2@gmail.com>
 * @since 1.0
 */
class ExportWriterPdf extends Mpdf
{
    /**
     * @var string the exported output file name. Defaults to 'grid-export';
     */
    public $filename = '';

    /**
     * @var array kartik\mpdf\Pdf component configuration settings
     */
    public $pdfConfig = [];

    /**
     * @inheritdoc
     */
    protected function createExternalWriterInstance($config = [])
    {
        $config = array_replace_recursive($config, $this->pdfConfig);
        return new Pdf($config);
    }

    /**
     * Save Spreadsheet to file.
     *
     * @param string $pFilename Name of the file to save as
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws PhpSpreadsheetException
     * @throws \yii\base\InvalidConfigException
     */
    public function save($pFilename)
    {
        $fileHandle = parent::prepareForSave($pFilename);

        //  Default PDF paper size
        $paperSize = Pdf::FORMAT_A4;

        //  Check for paper size and page orientation
        if (null === $this->getSheetIndex()) {
            $orientation = ($this->spreadsheet->getSheet(0)->getPageSetup()->getOrientation()
                == PageSetup::ORIENTATION_LANDSCAPE) ? 'L' : 'P';
            $printPaperSize = $this->spreadsheet->getSheet(0)->getPageSetup()->getPaperSize();
        } else {
            $orientation = ($this->spreadsheet->getSheet($this->getSheetIndex())->getPageSetup()->getOrientation()
                == PageSetup::ORIENTATION_LANDSCAPE) ? 'L' : 'P';
            $printPaperSize = $this->spreadsheet->getSheet($this->getSheetIndex())->getPageSetup()->getPaperSize();
        }
        $this->setOrientation($orientation);

        //  Override Page Orientation
        if (null !== $this->getOrientation()) {
            $orientation = ($this->getOrientation() == PageSetup::ORIENTATION_DEFAULT)
                ? PageSetup::ORIENTATION_PORTRAIT
                : $this->getOrientation();
        }
        $orientation = strtoupper($orientation);

        //  Override Paper Size
        if (null !== $this->getPaperSize()) {
            $printPaperSize = $this->getPaperSize();
        }

        if (isset(self::$paperSizes[$printPaperSize])) {
            $paperSize = self::$paperSizes[$printPaperSize];
        }

        $properties = $this->spreadsheet->getProperties();

        //  Create PDF
        $pdf = $this->createExternalWriterInstance([
            'format' => $paperSize,
            'orientation' => $orientation,
            'methods' => [
                'SetTitle' => $properties->getTitle(),
                'SetAuthor' => $properties->getCreator(),
                'SetSubject' => $properties->getSubject(),
                'SetKeywords' => $properties->getKeywords(),
                'SetCreator' => $properties->getCreator(),
            ],
        ]);
        $content = $this->generateHTMLHeader(false) . $this->generateSheetData() . $this->generateHTMLFooter();
        //  Write to file
        fwrite($fileHandle, $pdf->output(static::cleanHTML($content), $this->filename, Pdf::DEST_STRING));
        parent::restoreStateAfterSave($fileHandle);
    }

    /**
     * Cleans HTML of embedded script, style, and link tags
     *
     * @param string $content the source HTML content
     * @return string the cleaned HTML content
     */
    protected static function cleanHTML($content)
    {
        if (empty($content)) {
            return $content;
        }
        $doc = new DOMDocument();
        $doc->loadHTML($content);
        static::removeElementsByTagName('script', $doc);
        static::removeElementsByTagName('style', $doc);
        static::removeElementsByTagName('link', $doc);
        return $doc->saveHTML();
    }

    /**
     * Remove DOM elements by tag name
     * @param string $tagName the tag name to parse and remove
     * @param DOMDocument $document the DomDocument object
     */
    protected static function removeElementsByTagName($tagName, $document)
    {
        $nodeList = $document->getElementsByTagName($tagName);
        for ($nodeIdx = $nodeList->length; --$nodeIdx >= 0;) {
            $node = $nodeList->item($nodeIdx);
            $node->parentNode->removeChild($node);
        }
    }
}