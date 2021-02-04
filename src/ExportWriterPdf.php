<?php

/**
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2020
 * @package yii2-export
 * @version 1.4.2
 */

namespace kartik\export;

use kartik\mpdf\Pdf;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf;

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
        if (isset($config['tempDir'])) {
            unset($config['tempDir']);
        }
        $config = array_replace_recursive($config, $this->pdfConfig);
        
        $pdf = new Pdf($config);
        return $pdf;
    }

    /**
     * @inheritdoc
     */
    public function save($pFilename): void
    {
        $fileHandle = parent::prepareForSave($pFilename);

        
        //  Create PDF
        $config = ['tempDir' => $this->tempDir . '/mpdf'];
        $pdf = $this->createExternalWriterInstance($config);
        
        $this->spreadsheet->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1,1);
        $pdf->cssInline = $this->generateStyles(false);
        $html = str_replace("<div style='page: page0'>", '<div>', $this->generateSheetData());
        //  Write to file
        fwrite($fileHandle, $pdf->output($html, '', 'S'));

        parent::restoreStateAfterSave();
    }
}
