<?php

/**
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015
 * @version   1.2.3
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
use yii\data\ActiveDataProvider;
use yii\db\ActiveQueryInterface;
use yii\web\View;
use yii\web\JsExpression;
use yii\bootstrap\ButtonDropdown;
use kartik\grid\GridView;

/**
 * Export menu widget. Export tabular data to various formats using the PHPExcel library
 * by reading data from a dataProvider - with configuration very similar to a GridView.
 *
 * @author Kartik Visweswaran <kartikv2@gmail.com>
 * @since  1.0
 */
class ExportMenu extends GridView
{
    use \kartik\base\TranslationTrait;

    /**
     * Export formats
     */
    const FORMAT_HTML = 'HTML';
    const FORMAT_CSV = 'CSV';
    const FORMAT_TEXT = 'TXT';
    const FORMAT_PDF = 'PDF';
    const FORMAT_EXCEL = 'Excel5';
    const FORMAT_EXCEL_X = 'Excel2007';

    /**
     * Export form submission targets
     */
    const TARGET_POPUP = '_popup';
    const TARGET_SELF = '_self';
    const TARGET_BLANK = '_blank';

    /**
     * Input parameters from export form
     */
    const PARAM_EXPORT_TYPE = 'export_type';
    const PARAM_EXPORT_COLS = 'export_columns';
    const PARAM_COLSEL_FLAG = 'column_selector_enabled';

    /**
     * @var string the target for submitting the export form, which will trigger
     * the download of the exported file. Must be one of the `TARGET_` constants.
     * Defaults to `ExportMenu::TARGET_POPUP`.
     */
    public $target = self::TARGET_POPUP;

    /**
     * @var bool whether to show a confirmation alert dialog before download. This
     * confirmation dialog will notify user about the type of exported file for download
     * and to disable popup blockers. Defaults to `true`.
     */
    public $showConfirmAlert = true;

    /**
     * @var bool whether to enable the yii gridview formatter component.
     * Defaults to `true`. If set to `false`, this will render content
     * as `raw` format.
     */
    public $enableFormatter = true;

    /**
     * @var bool whether to render the export menu as bootstrap button dropdown
     * widget. Defaults to `true`. If set to `false`, this will generate a simple
     * HTML list of links.
     */
    public $asDropdown = true;
    
    /**
     * @var string the pjax container identifier inside which this menu is being 
     * rendered. If set the jQuery export menu plugin will get auto initialized
     * on pjax request completion.
     */
    public $pjaxContainerId;

    /**
     * @var array the HTML attributes for the export button menu. Applicable only
     * if `asDropdown` is set to `true`. The following special options are
     * available:
     * - label: string, defaults to empty string
     * - icon: string, defaults to `<i class="glyphicon glyphicon-export"></i>`
     * - title: string, defaults to `Export data in selected format`.
     * - menuOptions: array, the HTML attributes for the dropdown menu.
     * - itemsBefore: array, any additional items that will be merged/prepended before with the export dropdown list.
     *     This should be similar to the `items` property as supported by `\yii\bootstrap\ButtonDropdown` widget. Note
     *     the page export items will be automatically generated based on settings in the `exportConfig` property.
     * - itemsAfter: array, any additional items that will be merged/appended after with the export dropdown list. This
     *     should be similar to the `items` property as supported by `\yii\bootstrap\ButtonDropdown` widget. Note the
     *     page export items will be automatically generated based on settings in the `exportConfig` property.
     */
    public $dropdownOptions = ['class' => 'btn btn-default'];

    /**
     * @var bool whether to clear all previous / parent buffers. Defaults to `false`.
     */
    public $clearBuffers = false;

    /**
     * @var bool whether to initialize data provider and clear models before rendering.
     * Defaults to `false`.
     */
    public $initProvider = false;

    /**
     * @var bool whether to show a column selector to select columns for export.
     * Defaults to `true`.
     */
    public $showColumnSelector = true;

    /**
     * @var array the configuration of the column names in the column selector. Note: column names will be
     * auto-generated anyway. Any setting in this property will override the auto-generated column names.
     * This list should be setup as `$key => $value` where:
     * $key: int, is the zero based index of the column as set in `$columns`.
     * $value: string, is the column name/label you wish to display in the column selector.
     */
    public $columnSelector = [];

    /**
     * @var array the HTML attributes for the column selector dropdown button.
     * The following special options are recognized:
     * - label: string, defaults to empty string.
     * - icon: string, defaults to `<i class="glyphicon glyphicon-list"></i>`
     * - title: string, defaults to `Select columns for export`.
     */
    public $columnSelectorOptions = [];

    /**
     * @var array the HTML attributes for the column selector menu list.
     */
    public $columnSelectorMenuOptions = [];

    /**
     * @var array the settings for the toggle all checkbox to check/uncheck
     * the columns as a batch. Should be setup as an associative array which
     * can have the following keys:
     * - `show`: bool, whether the batch toggle checkbox is to be shown. Defaults
     *    to `true`.
     * - `label`: string, the label to be displayed for toggle all. Defaults to
     *    `Toggle All`.
     * - `options`: array, the HTML attributes for the toggle label text. Defaults
     *   to `['class'=>'kv-toggle-all']`
     */
    public $columnBatchToggleSettings = [];

    /**
     * @var array, HTML attributes for the container to wrap the widget.
     * Defaults to ['class'=>'btn-group', 'role'=>'group']
     */
    public $container = ['class' => 'btn-group', 'role' => 'group'];

    /**
     * @var string, the template for rendering the content in the container. This will
     * be parsed only if `asDropdown` is `true`. The following tags will be replaced:
     * - {columns}: will be replaced with the column selector dropdown
     * - {menu}: will be replaced with export menu dropdown
     */
    public $template = "{columns}\n{menu}";

    /**
     * @var array the HTML attributes for the export form.
     */
    public $exportFormOptions = [];

    /**
     * @var array the selected column indexes for export. If not set this will default to
     * all columns.
     */
    public $selectedColumns;

    /**
     * @var array the column indexes for export that will be disabled for selection
     * in the column selector.
     */
    public $disabledColumns = [];

    /**
     * @var array the column indexes for export that will be hidden for selection
     * in the column selector, but will still be displayed in export output.
     */
    public $hiddenColumns = [];

    /**
     * @var array the column indexes for export that will not be exported at all nor
     * will they be shown in the column selector
     */
    public $noExportColumns = [];

    /**
     * @var string the view file for rendering the export form
     */
    public $exportFormView = '_form';

    /**
     * @var string the view file for rendering the columns selection
     */
    public $exportColumnsView = '_columns';

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
     * - filename: the base file name for the generated file. Defaults to 'grid-export'. This will be used to generate
     *     a default file name for downloading.
     * - extension: the extension for the file name
     * - alertMsg: string, the message prompt to show before saving. If this is empty or not set it will not be
     *     displayed.
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
     * - allowPopups: string, the message to be shown to disable browser popups for download. Defaults to `Disable any
     *     popup blockers in your browser to ensure proper download.`.
     * - confirmDownload: string, the message to be shown for confirming to proceed with the download. Defaults to `Ok
     *     to proceed?`.
     * - downloadProgress: string, the message to be shown in a popup dialog when download request is executed.
     *     Defaults to `Generating file. Please wait...`.
     * - downloadComplete: string, the message to be shown in a popup dialog when download request is completed.
     *     Defaults to
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
     * @var array the the internalization configuration for this widget
     */
    public $i18n = [];

    /**
     * @var string translation message file category name for i18n
     */
    protected $_msgCat = 'kvexport';

    /**
     * @var DataProvider the modified data provider for usage with export.
     */
    protected $_provider;

    /**
     * @var string the data output format type. Defaults to
     * `ExportMenu::FORMAT_EXCEL_X`.
     */
    private $_exportType = self::FORMAT_EXCEL_X;

    /**
     * @var array the default export configuration
     */

    private $_defaultExportConfig = [];

    /**
     * @var PHPExcel object instance
     */
    private $_objPHPExcel;

    /**
     * @var PHPExcel_Writer_IWriter object instance
     */
    private $_objPHPExcelWriter;

    /**
     * @var PHPExcel_Worksheet object instance
     */
    private $_objPHPExcelSheet;

    /**
     * @var int the header beginning row
     */
    private $_headerBeginRow = 1;

    /**
     * @var int the table beginning row
     */
    private $_beginRow = 1;

    /**
     * @var int the current table end row
     */
    private $_endRow = 1;

    /**
     * @var int the current table end column
     */
    private $_endCol = 1;

    /**
     * @var bool whether the column selector is enabled
     */
    private $_columnSelectorEnabled = true;

    /**
     * @var array the visble columns for export
     */
    private $_visibleColumns;

    /**
     * @var array the default style configuration
     */
    private $_defaultStyleOptions = [
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

    /**
     * @var bool flag to identify if download is triggered
     */
    private $_triggerDownload = false;

    /**
     * @inheritdoc
     */
    public function init()
    {
        $this->_columnSelectorEnabled = $this->showColumnSelector && $this->asDropdown;
        $this->_triggerDownload = !empty($_POST) &&
            !empty($_POST[$this->exportRequestParam]) &&
            $_POST[$this->exportRequestParam];
        if ($this->_triggerDownload) {
            Yii::$app->controller->layout = false;
            $this->_exportType = $_POST[self::PARAM_EXPORT_TYPE];
            $this->_columnSelectorEnabled = $_POST[self::PARAM_COLSEL_FLAG];
            $this->initSelectedColumns();
        }
        parent::init();
    }

    /**
     * Initialize columns selected for export
     */
    protected function initSelectedColumns()
    {
        if (!$this->_columnSelectorEnabled) {
            return;
        }
        $this->selectedColumns = array_keys($this->columnSelector);
        if (empty($_POST[self::PARAM_EXPORT_COLS])) {
            return;
        }
        $this->selectedColumns = explode(',', $_POST[self::PARAM_EXPORT_COLS]);
    }

    /**
     * @inheritdoc
     */
    public function run()
    {
        $this->initI18N(__DIR__);
        $this->initColumnSelector();
        $this->setVisibleColumns();
        $this->initExport();
        if (!$this->_triggerDownload) {
            $this->registerAssets();
            echo $this->renderExportMenu();
            return;
        }
        ob_end_clean();
        $config = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
        if (empty($config)) {
            throw new InvalidConfigException("The '{$this->pdfLibrary}' was not found or installed at path '{$path}'.");
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
        if ($this->_exportType === self::FORMAT_TEXT) {
            $writer->setDelimiter("\t");
        }
        if ($this->autoWidth) {
            foreach ($this->getVisibleColumns() as $n => $column) {
                $sheet->getColumnDimension(self::columnName($n + 1))->setAutoSize(true);
            }
        }
        $this->raiseEvent('onRenderSheet', [$sheet, $this]);
        if (!$this->stream) {
            $writer->save($this->filename . '.' . $config['extension']);
        } else {
            if ($this->clearBuffers) {
                while (ob_get_level() > 0) {
                    ob_end_clean();
                }
            } else {
                ob_end_clean();
            }
            $this->setHttpHeaders();
            $writer->save('php://output');
            Yii::$app->end();
        }
    }

    /**
     * Initialize column selector list
     */
    protected function initColumnSelector()
    {
        if (!$this->_columnSelectorEnabled) {
            return;
        }
        $selector = [];
        Html::addCssClass($this->columnSelectorOptions, 'btn btn-default dropdown-toggle');
        $header = ArrayHelper::getValue($this->columnSelectorOptions, 'header', Yii::t('kvexport', 'Select Columns'));
        $this->columnSelectorOptions['header'] = (empty($header) || $header === false) ? '' :
            '<li class="dropdown-header">' . $header . '</li><li class="kv-divider"></li>';
        $id = $this->options['id'] . '-cols';
        Html::addCssClass($this->columnSelectorMenuOptions, 'dropdown-menu kv-checkbox-list');
        $this->columnSelectorMenuOptions = array_replace_recursive([
            'id' => $id . '-list',
            'role' => 'menu',
            'aria-labelledby' => $id,
        ], $this->columnSelectorMenuOptions);
        $this->columnSelectorOptions = array_replace_recursive([
            'id' => $id,
            'icon' => '<i class="glyphicon glyphicon-list"></i>',
            'title' => Yii::t('kvexport', 'Select columns to export'),
            'type' => 'button',
            'data-toggle' => 'dropdown',
            'aria-haspopup' => 'true',
            'aria-expanded' => 'false',
        ], $this->columnSelectorOptions);
        foreach ($this->columns as $key => $column) {
            $selector[$key] = $this->getColumnLabel($key, $column);
        }
        $this->columnSelector = array_replace($selector, $this->columnSelector);
        if (!isset($this->selectedColumns)) {
            $keys = array_keys($this->columnSelector);
            $this->selectedColumns = array_combine($keys, $keys);
        }
    }

    /**
     * Fetches the column label
     *
     * @param int              $key
     * @param \yii\grid\Column $column
     *
     * @return string
     */
    protected function getColumnLabel($key, $column)
    {
        $label = Yii::t('kvexport', 'Column') . ' ' . ($key + 1);
        if (!empty($column->label)) {
            $label = $column->label;
        } elseif (!empty($column->header)) {
            $label = $column->header;
        } elseif (!empty($column->attribute)) {
            $label = $this->getAttributeLabel($column->attribute);
        } elseif (!$column instanceof \yii\grid\DataColumn) {
            $class = explode("\\", $column::classname());
            $label = Inflector::camel2words(end($class));
        }
        return trim(strip_tags(str_replace(['<br>', '<br/>'], ' ', $label)));
    }

    /**
     * Generates the attribute label
     *
     * @param string $attribute
     *
     * @return string
     */
    protected function getAttributeLabel($attribute)
    {
        $provider = $this->dataProvider;
        /** @var Model $model */
        if ($provider instanceof yii\data\ActiveDataProvider && $provider->query instanceof yii\db\ActiveQueryInterface) {
            $model = new $provider->query->modelClass;
            return $model->getAttributeLabel($attribute);
        } else {
            $models = $provider->getModels();
            if (($model = reset($models)) instanceof Model) {
                return $model->getAttributeLabel($attribute);
            } else {
                return Inflector::camel2words($attribute);
            }
        }
    }

    /**
     * Initializes export settings
     */
    public function initExport()
    {
        $this->_provider = clone($this->dataProvider);
        $this->_provider->pagination = false;
        if ($this->initProvider) {
            $this->_provider->prepare(true);
        }
       $this->styleOptions = ArrayHelper::merge($this->_defaultStyleOptions, $this->styleOptions);
        $this->filterModel = null;
        $this->setDefaultExportConfig();
        $this->exportConfig = ArrayHelper::merge($this->_defaultExportConfig, $this->exportConfig);
        if (empty($this->filename)) {
            $this->filename = Yii::t('kvexport', 'grid-export');
        }
        $target = $this->target == self::TARGET_POPUP ? 'kvExportFullDialog' : $this->target;
        $id = ArrayHelper::getValue($this->exportFormOptions, 'id', $this->options['id'] . '-form');
        Html::addCssClass($this->exportFormOptions, 'kv-export-full-form');
        $this->exportFormOptions += [
            'id' => $id,
            'target' => $target,
            'data-pjax' => false
        ];
    }

    /**
     * Sets the default export configuration
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
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Tab Delimited Text')],
                'alertMsg' => Yii::t('kvexport', 'The TEXT export file will be generated for download.'),
                'mime' => 'text/plain',
                'extension' => 'txt',
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
                'mime' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
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
            'allowPopups' => Yii::t('kvexport',
                'Disable any popup blockers in your browser to ensure proper download.'),
            'confirmDownload' => Yii::t('kvexport', 'Ok to proceed?'),
            'downloadProgress' => Yii::t('kvexport', 'Generating the export file. Please wait...'),
            'downloadComplete' => Yii::t('kvexport',
                'Request submitted! You may safely close this dialog after saving your downloaded file.'),
        ];
        $formId = $this->exportFormOptions['id'];
        $options = Json::encode([
            'formId' => $formId,
            'messages' => $this->messages
        ]);
        $menu = 'kvexpmenu_' . hash('crc32', $options);
        $view->registerJs("var {$menu} = {$options};\n", View::POS_HEAD);
        $script = "";
        foreach ($this->exportConfig as $format => $setting) {
            if (empty($setting) || $setting === false) {
                continue;
            }
            $id = $this->options['id'] . '-' . strtolower($format);
            $options = [
                'settings' => new JsExpression($menu),
                'alertMsg' => $setting['alertMsg'],
                'target' => $this->target,
                'showConfirmAlert' => $this->showConfirmAlert
            ];
            if ($this->_columnSelectorEnabled) {
                $options['columnSelectorId'] = $this->columnSelectorOptions['id'];
            }
            $options = Json::encode($options);
            $script .= "jQuery('#{$id}').exportdata({$options});\n";
        }
        if ($this->_columnSelectorEnabled) {
            $id = $this->columnSelectorMenuOptions['id'];
            ExportColumnAsset::register($view);
            $script .= "jQuery('#{$id}').exportcolumns({});\n";
        }
        if (!empty($script) && isset($this->pjaxContainerId)) {
            $script .= "jQuery('#{$this->pjaxContainerId}').on('pjax:complete', function() {
                {$script}
            });\n";
        }
        $view->registerJs($script);
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
            if (empty($settings) || $settings === false) {
                continue;
            }
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
        $form = $this->render($this->exportFormView, [
            'options' => $this->exportFormOptions,
            'exportType' => $this->_exportType,
            'columnSelectorEnabled' => $this->_columnSelectorEnabled,
            'exportRequestParam' => $this->exportRequestParam,
            'exportTypeParam' => self::PARAM_EXPORT_TYPE,
            'exportColsParam' => self::PARAM_EXPORT_COLS,
            'colselFlagParam' => self::PARAM_COLSEL_FLAG,
        ]);
        if ($this->asDropdown) {
            $icon = ArrayHelper::remove($this->dropdownOptions, 'icon', '<i class="glyphicon glyphicon-export"></i>');
            $label = ArrayHelper::remove($this->dropdownOptions, 'label', '');
            $label = empty($label) ? $icon : $icon . ' ' . $label;
            if (empty($this->dropdownOptions['title'])) {
                $this->dropdownOptions['title'] = Yii::t('kvexport', 'Export data in selected format');
            }
            $menuOptions = ArrayHelper::remove($this->dropdownOptions, 'menuOptions', []);
            $itemsBefore = ArrayHelper::remove($this->dropdownOptions, 'itemsBefore', []);
            $itemsAfter = ArrayHelper::remove($this->dropdownOptions, 'itemsAfter', []);
            $items = ArrayHelper::merge($itemsBefore, $items, $itemsAfter);
            $content = strtr($this->template, [
                    '{menu}' => ButtonDropdown::widget([
                        'label' => $label,
                        'dropdown' => ['items' => $items, 'encodeLabels' => false, 'options' => $menuOptions],
                        'options' => $this->dropdownOptions,
                        'encodeLabel' => false
                    ]),
                    '{columns}' => $this->renderColumnSelector()
                ]) . "\n" . $form;
            return Html::tag('div', $content, $this->container);
        } else {
            return $items . "\n" . $form;
        }
    }

    /**
     * Renders the columns selector
     */
    public function renderColumnSelector()
    {
        if (!$this->_columnSelectorEnabled) {
            return '';
        }
        return $this->render($this->exportColumnsView, [
            'options' => $this->columnSelectorOptions,
            'menuOptions' => $this->columnSelectorMenuOptions,
            'columnSelector' => $this->columnSelector,
            'batchToggle' => $this->columnBatchToggleSettings,
            'selectedColumns' => $this->selectedColumns,
            'disabledColumns' => $this->disabledColumns,
            'hiddenColumns' => $this->hiddenColumns,
            'noExportColumns' => $this->noExportColumns
        ]);
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
        $lastModifiedBy = 'krajee';
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
     *
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
     * Raises a callable event
     *
     * @param string $event the event name
     * @param array  $params the parameters to the callable function
     */
    protected function raiseEvent($event, $params)
    {
        if (isset($this->$event) && is_callable($this->$event)) {
            call_user_func_array($this->$event, $params);
        }
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
     * Generates the output data header content.
     */
    public function generateHeader()
    {
        $columns = $this->getVisibleColumns();
        if (count($columns) == 0) {
            return;
        }
        $cells = [];
        $sheet = $this->_objPHPExcelSheet;
        $style = ArrayHelper::getValue($this->styleOptions, $this->_exportType, []);
        $colFirst = self::columnName(1);
        if (!empty($this->caption)) {
            $cell = $sheet->setCellValue($colFirst . $this->_beginRow, $this->caption, true);
            $this->_beginRow += 2;
        }
        $this->_endCol = 0;
        foreach ($this->getVisibleColumns() as $column) {
            $this->_endCol++;
            /* @var $column Column */
            $head = ($column instanceof \yii\grid\DataColumn) ? $this->getColumnHeader($column) : $column->header;
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
     * Gets the visible columns for export
     */
    public function getVisibleColumns()
    {
        if (!$this->_columnSelectorEnabled) {
            return $this->columns;
        }
        return $this->_visibleColumns;
    }

    /**
     * Sets visible columns for export
     */
    protected function setVisibleColumns()
    {
        if (!$this->_columnSelectorEnabled) {
            $this->_visibleColumns = $this->columns;
            return;
        }
        $cols = [];
        foreach ($this->columns as $key => $column) {
            if (!in_array($key, $this->noExportColumns) && in_array($key, $this->selectedColumns)) {
                $cols[] = $column;
            }
        }
        $this->_visibleColumns = $cols;
    }

    /**
     * Returns an excel column name.
     *
     * @param int $index the column index number
     *
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
        return 'A';
    }

    /**
     * Gets the column header content
     *
     * @param \yii\grid\DataColumn $col
     *
     * @return string
     */
    public function getColumnHeader($col)
    {
        /* @var $model yii\base\Model */
        if ($col->header !== null || ($col->label === null && $col->attribute === null)) {
            return trim($col->header) !== '' ? $col->header : $col->grid->emptyCell;
        }
        $provider = $this->dataProvider;
        if ($col->label === null) {
            if ($provider instanceof ActiveDataProvider && $provider->query instanceof ActiveQueryInterface) {
                $model = new $provider->query->modelClass;
                $label = $model->getAttributeLabel($col->attribute);
            } else {
                $models = $provider->getModels();
                if (($model = reset($models)) instanceof Model) {
                    $label = $model->getAttributeLabel($col->attribute);
                } else {
                    $label = Inflector::camel2words($col->attribute);
                }
            }
        } else {
            $label = $col->label;
        }
        return $label;
    }

    /**
     * Generates the output data body content.
     *
     * @return int the number of output rows.
     */
    public function generateBody()
    {
        $columns = $this->getVisibleColumns();
        $models = array_values($this->_provider->getModels());
        if (count($columns) == 0) {
            $cell = $this->_objPHPExcelSheet->setCellValue('A1', $this->emptyText, true);
            $model = reset($models);
            $this->raiseEvent('onRenderDataCell', [$cell, $this->emptyText, $model, null, 0, $this]);
            return 0;
        }
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
     *
     * @param mixed   $model the data model to be rendered
     * @param mixed   $key the key associated with the data model
     * @param integer $index the zero-based index of the data model among the model array returned by [[dataProvider]].
     */
    public function generateRow($model, $key, $index)
    {
        $cells = [];
        /* @var $column Column */
        $this->_endCol = 0;
        foreach ($this->getVisibleColumns() as $column) {
            if ($column instanceof \yii\grid\SerialColumn || $column instanceof \kartik\grid\SerialColumn) {
                $value = $column->renderDataCell($model, $key, $index);
            } elseif ($column instanceof \yii\grid\ActionColumn) {
                $value = '';
            } else {
                $format = $this->enableFormatter && isset($column->format) ? $column->format : 'raw';
                $value = ($column->content === null) ? (method_exists($column, 'getDataCellValue') ?
                    $this->formatter->format($column->getDataCellValue($model, $key, $index), $format) :
                    $column->renderDataCell($model, $key, $index)) :
                    call_user_func($column->content, $model, $key, $index, $column);
            }
            if (empty($value) && !empty($column->attribute) && $column->attribute !== null) {
                $value = ArrayHelper::getValue($model, $column->attribute, '');
            }
            $this->_endCol++;
            $cell = $this->_objPHPExcelSheet->setCellValue(self::columnName($this->_endCol) . ($index + $this->_beginRow + 1),
                strip_tags($value), true);
            $this->raiseEvent('onRenderDataCell', [$cell, $value, $model, $key, $index, $this]);
        }
    }

    /**
     * Generates the output footer row after a specific row number
     *
     * @param int $row the row number after which the footer is to be generated
     */
    public function generateFooter($row)
    {
        $columns = $this->getVisibleColumns();
        if (count($columns) == 0) {
            return;
        }
        $this->_endCol = 0;
        foreach ($this->getVisibleColumns() as $n => $column) {
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
     * Gets the PHP Excel object
     *
     * @return PHPExcel the current PHPExcel object instance
     */
    public function getPHPExcel()
    {
        return $this->_objPHPExcel;
    }

    /**
     * Gets the PHP Excel writer object
     *
     * @return PHPExcel_Writer_IWriter the current PHPExcel_Writer_IWriter object instance
     */
    public function getPHPExcelWriter()
    {
        return $this->_objPHPExcelWriter;
    }

    /**
     * Gets the PHP Excel sheet object
     *
     * @return PHPExcel_Worksheet the current PHPExcel_Worksheet object instance
     */
    public function getPHPExcelSheet()
    {
        return $this->_objPHPExcelSheet;
    }

    /**
     * Destroys PHP Excel Object Instance
     */
    public function destroyPHPExcel()
    {
        $this->_objPHPExcel->disconnectWorksheets();
        unset($this->_objPHPExcel);
    }
}