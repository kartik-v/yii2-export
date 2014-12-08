<?php

/**
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2014
 * @package yii2-export
 * @version 1.0.0
 */

namespace kartik\export;

use \Yii;
use \PHPExcel;
use \PHPExcel_IOFactory;
use \PHPExcel_Settings;
use \PHPExcel_Style_Fill;
use \PHPExcel_Writer_IWriter;
use \PHPExcel_Worksheet;
use \Closure;
use yii\base\InvalidConfigException;
use yii\base\InvalidValueException;
use yii\helpers\Html;
use yii\helpers\Json;
use yii\helpers\Inflector;
use yii\helpers\ArrayHelper;
use yii\data\DataProvider;
use yii\grid\GridView;
use yii\bootstrap\ButtonDropdown;
use yii\web\View;
use yii\web\JsExpression;

/**
 * Export menu widget. Export tabular data to various formats using the PHPExcel library
 * by reading data from a dataProvider - with configuration very similar to a GridView.
 *
 * @author Kartik Visweswaran <kartikv2@gmail.com>
 * @since  1.0
 */
class ExportMenu extends GridView
{
    const FORMAT_HTML = 'HTML';
    const FORMAT_CSV = 'CSV';
    const FORMAT_TEXT = 'TXT';
    const FORMAT_PDF = 'PDF';
    const FORMAT_EXCEL = 'Excel5';
    const FORMAT_EXCEL_X = 'Excel2007';

    /**
     * @var bool whether to render the export menu as bootstrap button dropdown
     * widget. Defaults to `true`. If set to `false`, this will generate a simple
     * HTML list of links.
     */
    public $asDropdown = true;

    /**
     * @var array the HTML attributes for the export button menu. Applicable only
     * if `asDropdown` is set to `true`. The following special options are
     * available:
     * - label: string, defaults to `<i class="glyphicon glyphicon-export"></i>'
     * - menuOptions: array, the HTML attributes for the dropdown menu.
     */
    public $dropdownOptions = ['class' => 'btn btn-default'];

    /**
     * @var boolean whether to use font awesome icons for rendering the icons
     * as defined in `exportConfig`. If set to `true`, you must load the FontAwesome
     * CSS separately in your application.
     */
    public $fontAwesome = false;

    /**
     * @var array the export configuration. The array keys must be the one of the `format` constants
     * (CSV, HTML, TEXT, EXCEL, PDF, JSON) and the array value is a configuration array consisiting of these settings:
     * - label: string,the label for the export format menu item displayed
     * - icon: string,the glyphicon or font-awesome name suffix to be displayed before the export menu item label. 
     *   If set to an empty string, this will not be displayed. 
     * - iconOptions: array, HTML attributes for export menu icon.
     * - linkOptions: array, HTML attributes for each export item link.
     * - filename: the base file name for the generated file. Defaults to 'grid-export'. This will be used to generate a default
     *   file name for downloading.
     * - extension: the extension for the file name
     * - alertMsg: string, the message prompt to show before saving. If this is empty or not set it will not be displayed.
     * - mime: string, the mime type (for the file format) to be set before downloading.
     * - writer: string, the PHP Excel writer type
     * - options: array, HTML attributes for the export menu item.
     */
    public $exportConfig = [];

    /**
     * @var string the request parameter ($_GET or $_POST) that
     * will be submitted during export
     */
    public $exportRequestParam = 'exportFull';

    /**
     * @var array the output style configuration options. It must be the style
     * configuration array as required by PHPExcel.
     */
    public $styleOptions = [];

    /**
     * @var bool whether to auto-size the excel output column widths.
     * Defaults to `true`.
     */
    public $autoWidth = true;
    
    /**
     * @var string encoding for the downloaded file header. Defaults to 'utf-8'.
     */
    public $encoding = 'utf-8';

    /**
     * @var string the exported output file name. Defaults to 'grid-export';
     */
    public $filename;

    /**
     * @var bool whether to stream output to the browser
     */
    public $stream = true;

    /**
     * @var array, the configuration of various messages that will be displayed at runtime:
     * - allowPopups: string, the message to be shown to disable browser popups for download. Defaults to `Disable any popup blockers in your browser to ensure proper download.`.
     * - confirmDownload: string, the message to be shown for confirming to proceed with the download. Defaults to `Ok to proceed?`.
     * - downloadProgress: string, the message to be shown in a popup dialog when download request is executed. Defaults to `Generating file. Please wait...`.
     * - downloadComplete: string, the message to be shown in a popup dialog when download request is completed. Defaults to 
     *   `All done! Click anywhere here to close this window, once you have downloaded the file.`.
     */
    public $messages = [];
    
    /**
     * @var Closure the callback function on initializing the writer. 
     * The anonymous function should have the following signature:
     * ```php
     * function ($writer, $grid)
     * ```
     * where:
     * - `$writer`: the PHPExcel_Writer_IWriter object instance
     * - `$grid`: the GridView object
     */
    public $onInitWriter = null;
    
    /**
     * @var Closure the callback function to be executed on initializing the active sheet. 
     * The anonymous function should have the following signature:
     * ```php
     * function ($sheet, $grid)
     * ```
     * where:
     * - `$sheet`: the PHPExcel_Worksheet object instance
     * - `$grid`: the GridView object
     */
    public $onInitSheet = null;
    
    /**
     * @var Closure the callback function to be executed on rendering the header cell output
     * content. The anonymous function should have the following signature:
     * ```php
     * function ($cell, $content, $grid)
     * ```
     * where:
     * - `$cell`: PHPExcel_Cell, is the current PHPExcel cell being rendered
     * - `$content`: string, is the header cell content being rendered
     * - `$grid`: GridView, is the current GridView object
     */
    public $onRenderHeaderCell = null;

    /**
     * @var Closure the callback function to be executed on rendering each body data cell
     * content. The anonymous function should have the following signature:
     * ```php
     * function ($cell, $content, $model, $key, $index, $grid)
     * ```
     * where:
     * - `$cell`: the current PHPExcel cell being rendered
     * - `$content`: the data cell content being rendered
     * - `$model`: the data model to be rendered
     * - `$key`: the key associated with the data model
     * - `$index`: the zero-based index of the data model among the model array returned by [[dataProvider]].     
     * - `$grid`: the GridView object
     */
    public $onRenderDataCell = null;

    /**
     * @var Closure the callback function to be executed on rendering the footer cell output
     * content. The anonymous function should have the following signature:
     * ```php
     * function ($cell, $content, $grid)
     * ```
     * where:
     * - `$cell`: the current PHPExcel cell being rendered
     * - `$content`: the footer cell content being rendered
     * - `$grid`: the GridView object
     */
    public $onRenderFooterCell = null;
    
    /**
     * @var Closure the callback function to be executed on rendering the sheet. The anonymous function
     * should have the following signature:
     * ```php
     * function ($sheet, $grid)
     * ```
     * where:
     * - `$sheet`: the current PHPExcel sheet being rendered
     * - `$grid`: the GridView object
     */
    public $onRenderSheet = null;

    /**
     * @var array the PHPExcel document properties
     */
    public $docProperties = [];

    /**
     * @var string the library used to render the PDF. Defaults to `'mPDF'`.
     * Must be one of:
     * - `PHPExcel_Settings::PDF_RENDERER_TCPDF` or `'tcPDF'`
     * - `PHPExcel_Settings::PDF_RENDERER_DOMPDF` or `'DomPDF'`
     * - `PHPExcel_Settings::PDF_RENDERER_MPDF` or `'mPDF'`
     */
    public $pdfLibrary = PHPExcel_Settings::PDF_RENDERER_MPDF;
    
    /**
     * @var the alias for the pdf library path to export to PDF
     */
    public $pdfLibraryPath = '@vendor/kartik-v/mpdf';

    /**
     * @var array the the internalization configuration for this module
     */
    public $i18n = [];

    /**
     * @var string the data output format type. Defaults to
     * `ExportMenu::FORMAT_EXCEL_X`.
     */
    protected $_exportType = self::FORMAT_EXCEL_X;

    /**
     * @var array the default export configuration
     */

    protected $_defaultExportConfig = [];

    /**
     * @var DataProvider the modified data provider for usage with export.
     */
    protected $_provider;

    /**
     * @var PHPExcel object instance
     */
    protected $_objPHPExcel;

    /**
     * @var PHPExcel_Writer_IWriter object instance
     */
    protected $_objPHPExcelWriter;
    
    /**
     * @var PHPExcel_Worksheet object instance
     */
    protected $_objPHPExcelSheet;

    /**
     * @var int the header beginning row
     */
    protected $_headerBeginRow = 1;

    /**
     * @var int the table beginning row
     */
    protected $_beginRow = 1;

    /**
     * @var int the current table end row
     */
    protected $_endRow = 1;

    /**
     * @var int the current table end column
     */
    protected $_endCol = 1;

    /**
     * @var array the default style configuration
     */
    protected $_defaultStyleOptions = [
        self::FORMAT_EXCEL => [
            'font' => [
                'bold' => true,
                'color' => [
                    'argb' => 'FFFFFFFF',
                ],
            ],
            'fill' => [
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => [
                    'argb' => '00000000',
                ],
            ],
        ],
        self::FORMAT_EXCEL_X => [
            'font' => [
                'bold' => true,
                'color' => [
                    'argb' => 'FFFFFFFF',
                ],
            ],
            'fill' => [
                'type' => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
                'startcolor' => [
                    'argb' => 'FFA0A0A0',
                ],
                'endcolor' => [
                    'argb' => 'FFFFFFFF',
                ],
            ],
        ],
    ];

    private $_triggerDownload = false;

    /**
     * @inherit doc
     */
    public function init()
    {
        $this->_triggerDownload = !empty($_POST) &&
            !empty($_POST[$this->exportRequestParam]) &&
            $_POST[$this->exportRequestParam];
        if ($this->_triggerDownload) {
            Yii::$app->controller->layout = false;
            $this->_exportType = $_POST['export_type'];
        }
        parent::init();
    }

    /**
     * @inherit doc
     */
    public function run()
    {
        $this->initI18N();
        $this->initExport();
        if (!$this->_triggerDownload) {
            $this->registerAssets();
            echo $this->renderExportMenu();
            return;
        }
        ob_end_clean();
        $config = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
        if (empty($config)) {
            throw new InvalidConfigException("The  '{$this->pdfLibrary}' was not found or installed at path '{$path}'.");
        }
        if (empty($config['writer'])) {
            throw new InvalidConfigException("The 'writer' setting for PHPExcel must be setup in 'exportConfig'.");
        }
        if ($this->_exportType === self::FORMAT_PDF) {
            $path = Yii::getAlias($this->pdfLibraryPath);
            if (!PHPExcel_Settings::setPdfRenderer($this->pdfLibrary, $path)) {
                throw new InvalidConfigException("The pdf rendering library '{$this->pdfLibrary}' was not found or installed at path '{$path}'.");
            }
        }
        $this->initPHPExcel();
        $this->initPHPExcelWriter($config['writer']);
        $this->initPHPExcelSheet();
        $this->generateHeader();
        $row = $this->generateBody();
        $this->generateFooter($row);
        $writer = $this->_objPHPExcelWriter;
        $sheet = $this->_objPHPExcelSheet;
        if ($this->autoWidth) {
            foreach ($this->columns as $n => $column) {
                $sheet->getColumnDimension(self::columnName($n + 1))->setAutoSize(true);
            }
        }
        $this->raiseEvent('onRenderSheet', [$sheet, $this]);
        if (!$this->stream) {
            $writer->save($this->filename . '.' . $config['extension']);
        } else {
            ob_end_clean();
            $this->setHttpHeaders();
            $writer->save('php://output');
            Yii::$app->end();
        }
    }

    /**
     * Initializes i18n settings
     */
    public function initI18N()
    {
        Yii::setAlias('@kvexport', dirname(__FILE__));
        if (empty($this->i18n)) {
            $this->i18n = [
                'class' => 'yii\i18n\PhpMessageSource',
                'basePath' => '@kvexport/messages',
                'forceTranslation' => true
            ];
        }
        Yii::$app->i18n->translations['kvexport'] = $this->i18n;
    }

    /**
     * Renders the export menu
     *
     * @return string the export menu markup
     */
    public function renderExportMenu()
    {
        $items = $this->asDropdown ? [] : '';
        foreach ($this->exportConfig as $format => $settings) {
            $icon = '';
            $label = '';
            if (!empty($settings['icon'])) {
                $css = $this->fontAwesome ? 'fa fa-' : 'glyphicon glyphicon-';
                $iconOptions = ArrayHelper::getValue($settings, 'iconOptions', []);
                Html::addCssClass($iconOptions, $css . $settings['icon']);
                $label = Html::tag('i', '', $iconOptions) . ' ';
            }
            if (!empty($settings['label'])) {
                $label .= $settings['label'];
            }
            $fmt = strtolower($format);
            $linkOptions = ArrayHelper::getValue($settings, 'linkOptions', []);
            $linkOptions['id'] = $this->options['id'] . '-' . $fmt;
            $linkOptions['data-format'] = $format;
            $options = ArrayHelper::getValue($settings, 'options', []);
            Html::addCssClass($linkOptions, "export-full-{$fmt}");
            if ($this->asDropdown) {
                $items[] = [
                    'label' => $label,
                    'url' => '#',
                    'linkOptions' => $linkOptions,
                    'options' => $options
                ];
            } else {
                $items .= Html::tag('li', Html::a($label, '#', $linkOptions), $options);
            }
        }
        $form = Html::beginForm('', 'post', [
                'class' => 'kv-export-full-form',
                'style' => 'display:none',
                'target' => 'kvDownloadDialog',
                'data-pjax' => false,
                'id' => $this->options['id'] . '-form'
            ]) .
            Html::hiddenInput('export_type', $this->_exportType) .
            Html::hiddenInput($this->exportRequestParam, 1) .
            '</form>';

        if ($this->asDropdown) {
            $title = ArrayHelper::remove($this->dropdownOptions, 'label', '<i class="glyphicon glyphicon-export"></i>');
            $menuOptions = ArrayHelper::remove($this->dropdownOptions, 'menuOptions', []);
            return ButtonDropdown::widget([
                'label' => $title,
                'dropdown' => ['items' => $items, 'encodeLabels' => false, 'options' => $menuOptions],
                'options' => $this->dropdownOptions,
                'encodeLabel' => false
            ]) . $form;
        } else {
            return $items;
        }
    }

    /**
     * Initializes export settings
     */
    public function initExport()
    {
        $this->_provider = clone($this->dataProvider);
        $this->_provider->pagination = false;
        $this->styleOptions = ArrayHelper::merge($this->_defaultStyleOptions, $this->styleOptions);
        $this->filterModel = null;
        $this->setDefaultExportConfig();
        $this->exportConfig = ArrayHelper::merge($this->_defaultExportConfig, $this->exportConfig);
        if (empty($this->filename)) {
            $this->filename = Yii::t('kvexport', 'grid-export');
        }
    }
    
    /**
     * Gets the PHP Excel object
     * @return PHPExcel the current PHPExcel object instance
     */
    public function getPHPExcel()
    {
        return $this->_objPHPExcel;
    }

    /**
     * Gets the PHP Excel writer object
     * @return PHPExcel_Writer_IWriter the current PHPExcel_Writer_IWriter object instance
     */
    public function getPHPExcelWriter()
    {
        return $this->_objPHPExcelWriter;
    }
    
    /**
     * Gets the PHP Excel sheet object
     * @return PHPExcel_Worksheet the current PHPExcel_Worksheet object instance
     */
    public function getPHPExcelSheet()
    {
        return $this->_objPHPExcelSheet;
    }

    /**
     * Initializes PHP Excel Object Instance
     */
    public function initPHPExcel()
    {
        $this->_objPHPExcel = new PHPExcel();
        $creator = '';
        $title = '';
        $subject = '';
        $description = Yii::t('kvexport', 'Grid export generated by Krajee ExportMenu widget (yii2-export)');
        $category = '';
        $keywords = '';
        $manager = '';
        $company = 'Krajee Solutions';
        $created = date("Y-m-d H:i:s");
        $lastModifiedBy = date("Y-m-d H:i:s");
        extract($this->docProperties);
        $this->_objPHPExcel->getProperties()
            ->setCreator($creator)
            ->setTitle($title)
            ->setSubject($subject)
            ->setDescription($description)
            ->setCategory($category)
            ->setKeywords($keywords)
            ->setManager($manager)
            ->setCompany($company)
            ->setCreated($created)
            ->setLastModifiedBy($lastModifiedBy);
    }
    
    /**
     * Initializes PHP Excel Writer Object Instance
     * @param string the writer type as set in export config
     */
    public function initPHPExcelWriter($writer)
    {
        $this->_objPHPExcelWriter = PHPExcel_IOFactory::createWriter(
            $this->_objPHPExcel,
            $writer
        );
        $this->raiseEvent('onInitWriter', [$this->_objPHPExcelWriter, $this]);
    }
    
    /**
     * Initializes PHP Excel Worksheet Instance
     */
    public function initPHPExcelSheet()
    {
        $this->_objPHPExcelSheet = $this->_objPHPExcel->getActiveSheet();
        $this->raiseEvent('onInitSheet', [$this->_objPHPExcelSheet, $this]);
    }

    /**
     * Destroys PHP Excel Object Instance
     */
    public function destroyPHPExcel()
    {
        $this->_objPHPExcel->disconnectWorksheets();
        unset($this->_objPHPExcel);
    }

    /**
     * Generates the output data header content.
     */
    public function generateHeader()
    {
        $cells = [];
        foreach ($this->columns as $column) {
            /* @var $column Column */
            $cells[] = $column->renderHeaderCell();
        }
        $content = Html::tag('tr', implode('', $cells), $this->headerRowOptions);
        $sheet = $this->_objPHPExcelSheet;
        $style = ArrayHelper::getValue($this->styleOptions, $this->_exportType, []);
        $colFirst = self::columnName(1);
        if (!empty($this->caption)) {
            $cell = $sheet->setCellValue($colFirst . $this->_beginRow, $this->caption, true);
            $this->_beginRow += 2;
        }
        $this->_endCol = 0;
        foreach ($this->columns as $column) {
            $this->_endCol++;
            if ($column->header === null && !empty($column->attribute)) {
                if ($this->_provider instanceof ActiveDataProvider && $this->_provider->query instanceof ActiveQueryInterface) {
                    /* @var $model Model */
                    $model = new $this->_provider->query->modelClass;
                    $head = $model->getAttributeLabel($column->attribute);
                } else {
                    $models = $this->_provider->getModels();
                    if (($model = reset($models)) instanceof Model) {
                        /* @var $model Model */
                        $head = $model->getAttributeLabel($column->attribute);
                    } else {
                        $head = Inflector::camel2words($column->attribute);
                    }
                }
            } else {
                $head = $column->header;
            }

            $cell = $sheet->setCellValue(self::columnName($this->_endCol) . $this->_beginRow, $head, true);
            // Apply formatting to header cell
            $cell = $sheet->getStyle(self::columnName($this->_endCol) . $this->_beginRow)->applyFromArray($style);
            $this->raiseEvent('onRenderHeaderCell', [$cell, $head, $this]);
        }

        for ($i = $this->_headerBeginRow; $i < ($this->_beginRow - 1); $i++) {
            $sheet->mergeCells($colFirst . $i . ":" . self::columnName($this->_endCol) . $i);
            $cell = $sheet->getStyle($colFirst . $i)->applyFromArray($style);
        }
        // Freeze the top row
        $sheet->freezePane($colFirst . ($this->_beginRow + 1));
    }

    /**
     * Generates the output data body content.
     * @return int the number of output rows.
     */
    public function generateBody()
    {
        $models = array_values($this->_provider->getModels());
        $keys = $this->_provider->getKeys();
        $this->_endRow = 0;
        foreach ($models as $index => $model) {
            $key = $keys[$index];
            $this->generateRow($model, $key, $index);
            $this->_endRow++;
        }
        // Set autofilter on
        $this->_objPHPExcelSheet->setAutoFilter(
            self::columnName(1) .
            $this->_beginRow .
            ":" .
            self::columnName($this->_endCol) .
            $this->_endRow
        );
        return ($this->_endRow > 0) ? count($models) : 0;
    }

    /**
     * Generates an output data row with the given data model and key.
     * @param mixed $model the data model to be rendered
     * @param mixed $key the key associated with the data model
     * @param integer $index the zero-based index of the data model among the model array returned by [[dataProvider]].
     */
    public function generateRow($model, $key, $index)
    {
        $cells = [];
        /* @var $column Column */
        $this->_endCol = 0;
        foreach ($this->columns as $column) {
            if ($column instanceof \yii\grid\SerialColumn || $column instanceof \kartik\grid\SerialColumn) {
                $value = $column->renderDataCell($model, $key, $index);
            } elseif (!empty($column->attribute) && $column->attribute !== null) {
                $value = empty($model[$column->attribute]) ? "" : $model[$column->attribute];
            } elseif ($column instanceof \yii\grid\ActionColumn) {
                $value = '';
            } else {
                $value = $column->renderDataCellContent();
            }
            $this->_endCol++;
            $cell = $this->_objPHPExcelSheet->setCellValue(self::columnName($this->_endCol) . ($index + $this->_beginRow + 1),
                strip_tags($value), true);
            $this->raiseEvent('onRenderDataCell', [$cell, $value, $model, $key, $index, $this]);
        }
    }

    /**
     * Generates the output footer row after a specific row number
     * @param int $row the row number after which the footer is to be generated
     */
    public function generateFooter($row)
    {
        $this->_endCol = 0;
        foreach ($this->columns as $n => $column) {
            $this->_endCol = $this->_endCol + 1;
            if ($column->footer) {
                $footer = trim($column->footer) !== '' ? $column->footer : $column->grid->blankDisplay;
                $cell = $this->_objPHPExcel->getActiveSheet()->setCellValue(self::columnName($this->_endCol) . ($row + 2),
                    $footer, true);
                $this->raiseEvent('onRenderFooterCell', [$cell, $footer, $this]);
            }
        }
    }

    /**
     * Returns an excel column name.
     *
     * @param int $index the column index number
     * @return string
     */
    public static function columnName($index)
    {
        $i = $index - 1;
        if ($i >= 0 && $i < 26) {
            return chr(ord('A') + $i);
        }
        if ($i > 25) {
            return (self::columnName($i / 26)) . (self::columnName($i % 26 + 1));
        }
        throw new InvalidValueException("Invalid Column # {$index}");
    }

    /**
     * Set HTTP headers for download
     */
    protected function setHttpHeaders()
    {
        $config = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
        $extension = ArrayHelper::getValue($config, 'extension', 'xlsx');
        $mime = ArrayHelper::getValue($config, 'mime', 'binary');
        if (strstr($_SERVER["HTTP_USER_AGENT"], "MSIE") == false) {
            header("Cache-Control: no-cache");
            header("Pragma: no-cache");
        } else {
            header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
            header("Pragma: public");
        }
        header("Expires: Sat, 26 Jul 1979 05:00:00 GMT");
        header("Content-Encoding: {$this->encoding}");
        header("Content-Type: {$mime}; charset={$this->encoding}");
        header("Content-Disposition: attachment; filename={$this->filename}.{$extension}");
        header("Cache-Control: max-age=0");
    }

    /**
     * Raises a callable event
     * @param string $event the event name
     * @param array $params the parameters to the callable function
     */
    protected function raiseEvent($event, $params)
    {
        if (isset($this->$event) && is_callable($this->$event)) {
            call_user_func_array($this->$event, $params);
        }
    }

    /**
     * Gets the default export configuration
     */
    protected function setDefaultExportConfig()
    {
        $isFa = $this->fontAwesome;
        $this->_defaultExportConfig = [
            self::FORMAT_HTML => [
                'label' => Yii::t('kvexport', 'HTML'),
                'icon' => $isFa ? 'file-text' : 'floppy-saved',
                'iconOptions' => ['class' => 'text-info'],
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Hyper Text Markup Language')],
                'alertMsg' => Yii::t('kvexport', 'The HTML export file will be generated for download.'),
                'mime' => 'text/html',
                'extension' => 'html',
                'writer' => 'HTML'
            ],
            self::FORMAT_CSV => [
                'label' => Yii::t('kvexport', 'CSV'),
                'icon' => $isFa ? 'file-code-o' : 'floppy-open',
                'iconOptions' => ['class' => 'text-primary'],
                'colDelimiter' => ",",
                'rowDelimiter' => "\r\n",
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Comma Separated Values')],
                'alertMsg' => Yii::t('kvexport', 'The CSV export file will be generated for download.'),
                'mime' => 'application/csv',
                'extension' => 'csv',
                'writer' => 'CSV'
            ],
            self::FORMAT_TEXT => [
                'label' => Yii::t('kvexport', 'Text'),
                'icon' => $isFa ? 'file-text-o' : 'floppy-save',
                'iconOptions' => ['class' => 'text-muted'],
                'colDelimiter' => ",",
                'rowDelimiter' => "\r\n",
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Tab Delimited Text')],
                'alertMsg' => Yii::t('kvexport', 'The TEXT export file will be generated for download.'),
                'mime' => 'text/plain',
                'extension' => 'csv',
                'writer' => 'CSV'
            ],
            self::FORMAT_PDF => [
                'label' => Yii::t('kvexport', 'PDF'),
                'icon' => $isFa ? 'file-pdf-o' : 'floppy-disk',
                'iconOptions' => ['class' => 'text-danger'],
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Portable Document Format')],
                'alertMsg' => Yii::t('kvexport', 'The PDF export file will be generated for download.'),
                'mime' => 'application/pdf',
                'extension' => 'pdf',
                'writer' => 'PDF'
            ],
            self::FORMAT_EXCEL => [
                'label' => Yii::t('kvexport', 'Excel 95 +'),
                'icon' => $isFa ? 'file-excel-o' : 'floppy-remove',
                'iconOptions' => ['class' => 'text-success'],
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Microsoft Excel 95+ (xls)')],
                'alertMsg' => Yii::t('kvexport', 'The EXCEL 95+ (xls) export file will be generated for download.'),
                'mime' => 'application/vnd.ms-excel',
                'extension' => 'xls',
                'writer' => 'Excel5'
            ],
            self::FORMAT_EXCEL_X => [
                'label' => Yii::t('kvexport', 'Excel 2007+'),
                'icon' => $isFa ? 'file-excel-o' : 'floppy-remove',
                'iconOptions' => ['class' => 'text-success'],
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Microsoft Excel 2007+ (xlsx)')],
                'alertMsg' => Yii::t('kvexport', 'The EXCEL 2007+ (xlsx) export file will be generated for download.'),
                'mime' => 'application/vnd.ms-excel',
                'extension' => 'xlsx',
                'writer' => 'Excel2007'
            ],
        ];
    }

    /**
     * Registers client assets needed for Export Menu widget
     */
    protected function registerAssets()
    {
        $view = $this->getView();
        ExportMenuAsset::register($view);
        $this->messages += [
            'allowPopups' => Yii::t('kvexport', 'Disable any popup blockers in your browser to ensure proper download.'),
            'confirmDownload' => Yii::t('kvexport', 'Ok to proceed?'),
            'downloadProgress' => Yii::t('kvexport', 'Generating the export file. Please wait...'),
            'downloadComplete' => Yii::t('kvexport', 'Request submitted! You may safely close this dialog after saving your downloaded file.'),
        ];
        $options = Json::encode([
            'formId' => $this->options['id'] . '-form',
            'messages' => $this->messages
        ]);
        $menu = 'kvexpmenu_' . hash('crc32', $options);
        $view->registerJs("var {$menu} = {$options};\n", View::POS_HEAD);
        foreach ($this->exportConfig as $format => $setting) {
            $id = $this->options['id'] . '-' . strtolower($format);
            $options = Json::encode([
                'settings' => new JsExpression($menu),
                'alertMsg' => $setting['alertMsg']
            ]);
            $view->registerJs("jQuery('#{$id}').exportdata({$options});");
        }
    }
}