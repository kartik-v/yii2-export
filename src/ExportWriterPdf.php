<?php

/**
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2023
 * @package yii2-export
 * @version 1.4.3
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

        return $pdf->getApi();
    }
}
