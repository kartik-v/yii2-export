<?php

/**
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2018
 * @version   1.3.2
 */

namespace kartik\export;

use Closure;
use kartik\base\TranslationTrait;
use kartik\dialog\Dialog;
use kartik\dynagrid\Dynagrid;
use kartik\grid\GridView;
use kartik\mpdf\Pdf;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\BaseWriter;
use PhpOffice\PhpSpreadsheet\Writer\Csv as WriterCsv;
use PhpOffice\PhpSpreadsheet\Writer\Html as WriterHtml;
use Yii;
use yii\base\InvalidConfigException;
use yii\base\Model;
use yii\base\Widget;
use yii\data\ActiveDataProvider;
use yii\data\BaseDataProvider;
use yii\db\ActiveQueryInterface;
use yii\grid\ActionColumn;
use yii\grid\Column;
use yii\grid\DataColumn;
use yii\grid\SerialColumn;
use yii\helpers\ArrayHelper;
use yii\helpers\Html;
use yii\helpers\Inflector;
use yii\helpers\Json;
use yii\helpers\Url;
use yii\web\JsExpression;
use yii\web\View;

/**
 * Export menu widget. Export tabular data to various formats using the `\PhpOffice\PhpSpreadsheet\Spreadsheet library
 * by reading data from a dataProvider - with configuration very similar to a GridView.
 *
 * @author Kartik Visweswaran <kartikv2@gmail.com>
 * @since  1.0
 */
class ExportMenu extends GridView
{
    use TranslationTrait;

    /**
     * HTML (Hyper Text Markup Language) export format
     */
    const FORMAT_HTML = 'Html';
    /**
     * CSV (comma separated values) export format
     */
    const FORMAT_CSV = 'Csv';
    /**
     * Text export format
     */
    const FORMAT_TEXT = 'Txt';
    /**
     * PDF (Portable Document Format) export format
     */
    const FORMAT_PDF = 'Pdf';
    /**
     * Microsoft Excel 95+ export format
     */
    const FORMAT_EXCEL = 'Xls';
    /**
     * Microsoft Excel 2007+ export format
     */
    const FORMAT_EXCEL_X = 'Xlsx';
    /**
     * Set download target for grid export to a popup browser window
     */
    const TARGET_POPUP = '_popup';
    /**
     * Set download target for grid export to the same open document on the browser
     */
    const TARGET_SELF = '_self';
    /**
     * Set download target for grid export to a new window that auto closes after download
     */
    const TARGET_BLANK = '_blank';
    /**
     * Export type input parameter for export form
     */
    const PARAM_EXPORT_TYPE = 'export_type';
    /**
     * Export columns input parameter for export form
     */
    const PARAM_EXPORT_COLS = 'export_columns';
    /**
     * Column selector flag parameter for export form
     */
    const PARAM_COLSEL_FLAG = 'column_selector_enabled';

    /**
     * @var string the target for submitting the export form, which will trigger the download of the exported file.
     * Must be one of the `TARGET_` constants. Defaults to [[TARGET_POPUP]]. Note if you set [[stream]] to `false`,
     * then this will be overridden to [[TARGET_SELF]].
     */
    public $target = self::TARGET_POPUP;

    /**
     * @var array configuration settings for the Krajee dialog widget that will be used to render alerts and
     * confirmation dialog prompts
     * @see http://demos.krajee.com/dialog
     */
    public $krajeeDialogSettings = [];

    /**
     * @var boolean whether to show a confirmation alert dialog before download. This confirmation dialog will notify
     * user about the type of exported file for download and to disable popup blockers. Defaults to `true`.
     */
    public $showConfirmAlert = true;

    /**
     * @var boolean whether to enable the yii gridview formatter component. Defaults to `true`. If set to `false`, this
     * will render content as `raw` format.
     */
    public $enableFormatter = true;

    /**
     * @var boolean whether to render the export menu as bootstrap button dropdown widget. Defaults to `true`. If set to
     * `false`, this will generate a simple HTML list of links.
     */
    public $asDropdown = true;

    /**
     * @var string the pjax container identifier inside which this menu is being rendered. If set the jQuery export
     * menu plugin will get auto initialized on pjax request completion.
     */
    public $pjaxContainerId;

    /**
     * @var boolean whether to clear all previous / parent buffers. Defaults to `false`.
     */
    public $clearBuffers = false;

    /**
     * @var array the HTML attributes for the export button menu. Applicable only if [[asDropdown]] is set to `true`.
     * The following special options are available:
     * - `label`: _string_, defaults to empty string
     * - `icon`: _string_, defaults to `<i class="glyphicon glyphicon-export"></i>`
     * - `title`: _string_, defaults to `Export data in selected format`.
     * - `menuOptions`: _array_, the HTML attributes for the dropdown menu.
     * - `itemsBefore`: _array_, any additional items that will be merged/prepended before with the export dropdown list.
     *   This should be similar to the `items` property as supported by [[ButtonDropdown]] widget. Note the page export
     *   items will be automatically generated based on settings in the [[exportConfig]] property.
     * - `itemsAfter`: _array_, any additional items that will be merged/appended after with the export dropdown list. This
     *   should be similar to the `items` property as supported by [[ButtonDropdown]] widget. Note the
     *   page export items will be automatically generated based on settings in the `exportConfig` property.
     */
    public $dropdownOptions = [];

    /**
     * @var boolean whether to initialize data provider and clear models before rendering. Defaults to `false`.
     */
    public $initProvider = false;

    /**
     * @var boolean whether to show a column selector to select columns for export. Defaults to `true`.
     */
    public $showColumnSelector = true;

    /**
     * @var boolean enable or disable cell formatting by auto detecting the grid column alignment and format.
     * If set to `false` the format will not be applied but improve performance configured in [[dynagridOptions]].
     */
    public $enableAutoFormat = true;

    /**
     * @var array the configuration of the column names in the column selector. Note: column names will be generated
     * automatically by default. Any setting in this property will override the auto-generated column names. This
     * list should be setup as `$key => $value` where:
     * - `$key`: _integer_, is the zero based index of the column as set in `$columns`.
     * - `$value`: _string_, is the column name/label you wish to display in the column selector.
     */
    public $columnSelector = [];

    /**
     * @var array the HTML attributes for the column selector dropdown button. The following special options are
     * recognized:
     * - `label`: _string_, defaults to empty string.
     * - `icon`: _string_, defaults to `<i class="glyphicon glyphicon-list"></i>`
     * - `title`: _string_, defaults to `Select columns for export`.
     */
    public $columnSelectorOptions = [];

    /**
     * @var array the HTML attributes for the column selector menu list.
     */
    public $columnSelectorMenuOptions = [];

    /**
     * @var array the settings for the toggle all checkbox to check / uncheck the columns as a batch. Should be setup as
     * an associative array which can have the following keys:
     * - `show`: _boolean_, whether the batch toggle checkbox is to be shown. Defaults to `true`.
     * - `label`: _string_, the label to be displayed for toggle all. Defaults to `Toggle All`.
     * - `options`: _array_, the HTML attributes for the toggle label text. Defaults to `['class'=>'kv-toggle-all']`
     */
    public $columnBatchToggleSettings = [];

    /**
     * @var array, HTML attributes for the container to wrap the widget. Defaults to:
     * `['class'=>'btn-group', 'role'=>'group']`
     */
    public $container = ['class' => 'btn-group', 'role' => 'group'];

    /**
     * @var string, the template for rendering the content in the container. This will be parsed only if `asDropdown`
     * is `true`. The following tokens will be replaced:
     * - `{columns}`: will be replaced with the column selector dropdown
     * - `{menu}`: will be replaced with export menu dropdown
     */
    public $template = "{columns}\n{menu}";

    /**
     * @var integer timeout for the export function (in seconds), if timeout is < 0, the default PHP timeout will be used.
     */
    public $timeout = -1;

    /**
     * @var array the HTML attributes for the export form.
     */
    public $exportFormOptions = [];

    /**
     * @var array the configuration of additional hidden inputs that will be rendered with the export form. This will be
     * set as an array of hidden input `$name => $setting` pairs where:
     * - `$name`: _string_, is the hidden input name attribute.
     * - `$setting`: _array_, containing the following settings:
     *    - `value`: _string_, the value of the hidden input. Defaults to NULL.
     *    - `options`: _array_, the HTML attributes for the hidden input. Defaults to an empty array `[]`.
     *
     * An example of how you can configure this property is shown below:
     *
     * ```
     * [
     *    'inputName1' => ['value' => 'inputValue1'],
     *    'inputName2' => ['value' => 'inputValue2', 'options' => ['id' => 'inputId2']],
     * ]
     * ```
     */
    public $exportFormHiddenInputs = [];

    /**
     * @var array the selected column indexes for export. If not set this will default to all columns.
     */
    public $selectedColumns;

    /**
     * @var array the column indexes for export that will be disabled for selection in the column selector.
     */
    public $disabledColumns = [];

    /**
     * @var array the column indexes for export that will be hidden for selection in the column selector, but will
     * still be displayed in export output.
     */
    public $hiddenColumns = [];

    /**
     * @var array the column indexes for export that will not be exported at all nor will they be shown in the column
     * selector
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
     * @var boolean whether to use font awesome icons for rendering the icons as defined in [[exportConfig]]. If set to
     * `true`, you must load the FontAwesome CSS separately in your application.
     */
    public $fontAwesome = false;

    /**
     * @var boolean whether to strip HTML tags from each of the source column data before rendering the PHP
     * Spreadsheet Cell.
     */
    public $stripHtml = true;

    /**
     * @var array the export configuration. The array keys must be the one of the `format` constants (CSV, HTML, TEXT,
     * EXCEL, PDF) and the array value is a configuration array consisting of these settings:
     * - `label`: _string_, the label for the export format menu item displayed
     * - `icon`: _string_, the glyphicon or font-awesome name suffix to be displayed before the export menu item label. If
     *   set to an empty string_, this will not be displayed.
     * - `iconOptions`: _array_, HTML attributes for export menu icon.
     * - `linkOptions`: _array_, HTML attributes for each export item link.
     * - `filename`: _string_, the base file name for the generated file. Defaults to 'grid-export'. This will be used to generate
     *   a default file name for downloading.
     * - `extension`: _string_, the extension for the file name
     * - `alertMsg`: _string_, the message prompt to show before saving. If this is empty or not set it will not be
     *   displayed.
     * - `mime`: _string_, the mime type (for the file format) to be set before downloading.
     * - `writer`: _string_, the PhpSpreadsheet writer type
     * - `options`: _array_, HTML attributes for the export menu item.
     */
    public $exportConfig = [];

    /**
     * @var string the request parameter ($_GET or $_POST) that will be submitted during export. If not set this will
     *  be auto generated. This should be unique for each export menu widget (for multiple export menu widgets on
     *  same page).
     */
    public $exportRequestParam;

    /**
     * @var array the output style configuration options for each data cell. It must be the style configuration
     * array as required by `\PhpOffice\PhpSpreadsheet\Spreadsheet`.
     */
    public $styleOptions = [];

    /**
     * @var array the output style configuration options for the header row. It must be the style configuration array as
     * required by `\PhpOffice\PhpSpreadsheet\Spreadsheet`.
     */
    public $headerStyleOptions = [];

    /**
     * @var array the output style configuration options for the entire spreadsheet box range. It must be the style
     * configuration array as required by `\PhpOffice\PhpSpreadsheet\Spreadsheet`.
     */
    public $boxStyleOptions = [];

    /**
     * @var array an array of rows to prepend in front of the grid used to create things like a title. Each array
     * should be set with the following settings:
     * - value: string, the value of the merged row
     * - styleOptions: array, array of configuration options to set the style. See $styleOptions on how to configure.
     */
    public $contentBefore = [];

    /**
     * @var array an array of rows to append after the footer row. Each array
     * should be set with the following settings:
     * - value: string, the value of the merged row
     * - styleOptions: array, array of configuration options to set the style. See $styleOptions on how to configure.
     */
    public $contentAfter = [];

    /**
     * @var boolean whether to auto-size the excel output column widths. Defaults to `true`.
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
     * @var string the folder to save the exported file. Defaults to '@app/runtime/export/'. If the specified folder
     * does not exist, the extension will attempt to create it - else an exception will be thrown.
     */
    public $folder = '@app/runtime/export';

    /**
     * @var string the web accessible path for the saved file location. This property will be parsed only if [[stream]]
     * is false. Note the [[afterSaveView]] property that will render the displayed file link.
     */
    public $linkPath = '/runtime/export';

    /**
     * @var string the name of the file to be appended to [[linkPath]] to generate the complete link. If not set, this
     * will default to the [[filename]].
     */
    public $linkFileName;

    /**
     * @var boolean whether to stream output to the browser.
     */
    public $stream = true;

    /**
     * @var boolean whether to delete file after saving file to [[folder]] and when [[stream]] is `false`. This property
     * will be validated only when [[stream]] is `false`.
     */
    public $deleteAfterSave = false;

    /**
     * @var string|bool the view file to show details of exported file link. This property will be validated only when
     * [[stream]] is `false`. You can set this to `false` to not display any file link details for view. This defaults
     * to the `_view` PHP file in the `views` folder of the extension.
     */
    public $afterSaveView = '_view';

    /**
     * @var integer  fetch models from the dataprovider using batches of this size. Set this to `0` (the default) to
     * disable. If `$dataProvider` does not have a pagination object, this parameter is ignored. Setting this
     * property helps reduce memory overflow issues by allowing parsing of models in batches, rather than fetching
     * all models in one go.
     */
    public $batchSize = 0;

    /**
     * @var array, the configuration of various messages that will be displayed at runtime:
     * - allowPopups: string, the message to be shown to disable browser popups for download. Defaults to `Disable any
     *   popup blockers in your browser to ensure proper download.`.
     * - confirmDownload: string, the message to be shown for confirming to proceed with the download. Defaults to `Ok
     *   to proceed?`.
     * - downloadProgress: string, the message to be shown in a popup dialog when download request is executed.
     *   Defaults to `Generating file. Please wait...`.
     * - downloadComplete: string, the message to be shown in a popup dialog when download request is completed.
     *   Defaults to `All done! Click anywhere here to close this window, once you have downloaded the file.`.
     */
    public $messages = [];

    /**
     * @var Closure the callback function on initializing the PhpSpreadsheet library. The anonymous function should have the
     * following signature:
     * ```php
     * function ($spreadsheet, $widget)
     * ```
     * where:
     * - `$spreadsheet`: \PhpOffice\PhpSpreadsheet\Spreadsheet, the Spreadsheet object instance
     * - `$widget`: ExportMenu, the current ExportMenu object instance
     */
    public $onInitExcel = null;

    /**
     * @var Closure the callback function on initializing the writer. The anonymous function should have the following
     * signature:
     * ```php
     * function ($writer, $widget)
     * ```
     * where:
     * - `$writer`: \PhpOffice\PhpSpreadsheet\Writer\BaseWriter, the BaseWriter object instance
     * - `$widget`: ExportMenu, the current ExportMenu object instance
     */
    public $onInitWriter = null;

    /**
     * @var Closure the callback function to be executed on initializing the active sheet. The anonymous function
     * should have the following signature:
     * ```php
     * function ($sheet, $widget)
     * ```
     * where:
     * - `$sheet`: \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet, the Worksheet object instance
     * - `$widget`: ExportMenu, the current ExportMenu object instance
     */
    public $onInitSheet = null;

    /**
     * @var Closure the callback function to be executed on rendering the header cell output content. The anonymous
     * function should have the following signature:
     * ```php
     * function ($cell, $content, $widget)
     * ```
     * where:
     * - `$cell`: \PhpOffice\PhpSpreadsheet\Cell\Cell, is the current Spreadsheet cell being rendered
     * - `$content`: string, is the header cell content being rendered
     * - `$widget`: ExportMenu, the current ExportMenu object instance
     */
    public $onRenderHeaderCell = null;

    /**
     * @var Closure the callback function to be executed on rendering each body data cell content. The anonymous
     * function should have the following signature:
     * ```php
     * function ($cell, $content, $model, $key, $index, $widget)
     * ```
     * where:
     * - `$cell`: \PhpOffice\PhpSpreadsheet\Cell\Cell, the current Spreadsheet cell being rendered
     * - `$content`: string, the data cell content being rendered
     * - `$model`: Model, the data model to be rendered
     * - `$key`: mixed, the key associated with the data model
     * - `$index`: integer, the zero-based index of the data model among the model array returned by [[dataProvider]].
     * - `$widget`: ExportMenu, the current ExportMenu object instance
     */
    public $onRenderDataCell = null;

    /**
     * @var Closure the callback function to be executed on rendering the footer cell output content. The anonymous
     * function should have the following signature:
     * ```php
     * function ($cell, $content, $widget)
     * ```
     * where:
     * - `$sheet`: \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet, the Worksheet object instance
     * - `$content`: string, the footer cell content being rendered
     * - `$widget`: ExportMenu, the current ExportMenu object instance
     */
    public $onRenderFooterCell = null;

    /**
     * @var Closure the callback function to be executed on rendering the sheet. The anonymous function should have the
     * following signature:
     * ```php
     * function ($sheet, $widget)
     * ```
     * where:
     * - `$sheet`: \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet, the Worksheet object instance
     * - `$widget`: ExportMenu, the current ExportMenu object instance
     */
    public $onRenderSheet = null;

    /**
     * @var Closure the callback function to be executed after the output file is generated. This function must return
     * a boolean status of `true` or `false`. A `false` status will abort the post file generation activities. The
     * anonymous function should have the following signature:
     * ```php
     * function ($fileExt, $widget)
     * ```
     * where:
     * - `$fileExt`: _string_, is the generated file extension.
     * - `$widget`: ExportMenu, the current ExportMenu object instance
     */
    public $onGenerateFile = null;

    /**
     * @var array the \PhpOffice\PhpSpreadsheet\Spreadsheet document properties
     */
    public $docProperties = [];

    /**
     * @var boolean enable dynagrid for column selection. If set to `true` the inbuilt export menu column selector
     * functionality will be disabled and not rendered and column settings for dynagrid will be used as per settings
     * configured in [[dynagridOptions]].
     */
    public $dynagrid = false;

    /**
     * @var array dynagrid widget options. Applicable only if [[dynagrid]] is set to `true`.
     */
    public $dynagridOptions = ['options' => ['id' => 'dynagrid-export-menu']];

    /**
     * @var array the PhpSpreadsheet style configuration for a grouped grid row
     */
    public $groupedRowStyle = [
        'font' => [
            'bold' => false,
            'color' => [
                'argb' => Color::COLOR_DARKBLUE,
            ],
        ],
        'fill' => [
            'type' => Fill::FILL_SOLID,
            'color' => [
                'argb' => Color::COLOR_WHITE,
            ],
        ],
    ];

    /**
     * @var BaseDataProvider the modified data provider for usage with export.
     */
    protected $_provider;

    /**
     * @var string the data output format type. Defaults to `ExportMenu::FORMAT_EXCEL_X`.
     */
    protected $_exportType = self::FORMAT_EXCEL_X;

    /**
     * @var array the default export configuration
     */

    protected $_defaultExportConfig = [];

    /**
     * @var Spreadsheet object instance
     */
    protected $_objSpreadsheet;

    /**
     * @var BaseWriter object instance
     */
    protected $_objWriter;

    /**
     * @var Worksheet object instance
     */
    protected $_objWorksheet;

    /**
     * @var integer the header beginning row
     */
    protected $_headerBeginRow = 1;

    /**
     * @var integer  the table beginning row
     */
    protected $_beginRow = 1;

    /**
     * @var integer  the current table end row
     */
    protected $_endRow = 0;

    /**
     * @var integer  the current table end column
     */
    protected $_endCol = 1;

    /**
     * @var boolean whether the column selector is enabled
     */
    protected $_columnSelectorEnabled = true;

    /**
     * @var array the visble columns for export
     */
    protected $_visibleColumns;

    /**
     * @var array columns to be grouped
     */
    protected $_groupedColumn = [];

    /**
     *
     * @var array grouped row values
     */
    protected $_groupedRow = null;

    /**
     * @var boolean flag to identify if download is triggered
     */
    protected $_triggerDownload = false;

    /**
     * Appends slash to path if it does not exist
     *
     * @param string $path
     * @param string $s the path separator
     *
     * @return string
     */
    public static function slash($path, $s = DIRECTORY_SEPARATOR)
    {
        $path = trim($path);
        if (substr($path, -1) !== $s) {
            $path .= $s;
        }
        return $path;
    }

    /**
     * Returns an excel column name.
     *
     * @param integer $index the column index number
     *
     * @return string
     */
    public static function columnName($index)
    {
        $i = intval($index) - 1;
        if ($i >= 0 && $i < 26) {
            return chr(ord('A') + $i);
        }
        if ($i > 25) {
            return (self::columnName($i / 26)) . (self::columnName($i % 26 + 1));
        }
        return 'A';
    }

    /**
     * @inheritdoc
     */
    public function init()
    {
        $this->_msgCat = 'kvexport';
        if (empty($this->options['id'])) {
            $this->options['id'] = $this->getId();
        }
        if (empty($this->exportRequestParam)) {
            $this->exportRequestParam = 'exportFull_' . $this->options['id'];
        }
        $this->_columnSelectorEnabled = $this->showColumnSelector && $this->asDropdown;
        $this->_triggerDownload = Yii::$app->request->post($this->exportRequestParam, false);
        if (!$this->stream) {
            $this->target = self::TARGET_SELF;
        }
        if ($this->_triggerDownload) {
            if ($this->stream) {
                Yii::$app->controller->layout = false;
            }
            $this->_exportType = $_POST[self::PARAM_EXPORT_TYPE];
            $this->_columnSelectorEnabled = $_POST[self::PARAM_COLSEL_FLAG];
            $this->initSelectedColumns();
        }
        if ($this->dynagrid) {
            $this->_columnSelectorEnabled = false;
            $options = $this->dynagridOptions;
            $options['columns'] = $this->columns;
            $options['storage'] = 'db';
            $options['gridOptions']['dataProvider'] = $this->dataProvider;
            $dynagrid = new DynaGrid($options);
            $this->columns = $dynagrid->getColumns();
        }
        parent::init();
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
        if ($this->timeout >= 0) {
            set_time_limit($this->timeout);
        }
        if ($this->stream) {
            $this->clearOutputBuffers();
        }
        $config = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
        if (empty($config['writer'])) {
            throw new InvalidConfigException(
                "The 'writer' setting for '\PhpOffice\PhpSpreadsheet\Spreadsheet' must be setup in 'exportConfig'."
            );
        }
        $this->initPhpSpreadsheet();
        $this->initPhpSpreadsheetWriter($config['writer']);
        $this->initPhpSpreadsheetWorksheet();
        $this->generateBeforeContent();
        $this->generateHeader();
        $this->generateBody();
        $row = $this->generateFooter();
        $this->generateAfterContent($row);
        $writer = $this->_objWriter;
        $sheet = $this->_objWorksheet;
        if ($this->autoWidth) {
            foreach ($this->getVisibleColumns() as $n => $column) {
                $sheet->getColumnDimension(self::columnName($n + 1))->setAutoSize(true);
            }
        }
        $this->raiseEvent('onRenderSheet', [$sheet, $this]);
        $this->folder = trim(Yii::getAlias($this->folder));
        if (!file_exists($this->folder) && !mkdir($this->folder, 0777, true)) {
            throw new InvalidConfigException(
                "Invalid permissions to write to '{$this->folder}' as set in `ExportMenu::folder` property."
            );
        }
        $file = self::slash($this->folder) . $this->filename . '.' . $config['extension'];
        $writer->save($file);
        if ($this->stream) {
            $this->clearOutputBuffers();
            $this->setHttpHeaders();
            if ($this->_exportType === self::FORMAT_PDF) {
                $this->renderPDF($file);
            } else {
                readfile($file);
            }
            $this->cleanup($file, $config);
            exit();
        } else {
            $this->registerAssets();
            echo $this->renderExportMenu();
            if ($this->_triggerDownload && $this->afterSaveView !== false) {
                if ($this->_exportType === self::FORMAT_PDF) {
                    $this->renderPDF($file);
                }
                $config = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
                if (!empty($config)) {
                    $l = $this->linkFileName;
                    $fileName = (!isset($l) || $l === '' ? $this->filename : $l) . '.' . $config['extension'];
                    echo $this->render(
                        $this->afterSaveView,
                        [
                            'isBs4' => $this->isBs4(),
                            'file' => $fileName,
                            'icon' => $config['icon'],
                            'href' => Url::to([self::slash($this->linkPath, '/') . $fileName]),
                        ]
                    );
                }
            }
        }
        $this->cleanup($file, $config);
    }

    /**
     * Initializes export settings
     */
    public function initExport()
    {
        $this->_provider = clone($this->dataProvider);
        if ($this->batchSize && $this->_provider->pagination) {
            /** @noinspection PhpUndefinedFieldInspection */
            $this->_provider->pagination = clone($this->dataProvider->pagination);
            $this->_provider->pagination->pageSize = $this->batchSize;
        } else {
            $this->_provider->pagination = false;
        }
        if ($this->initProvider) {
            $this->_provider->prepare(true);
        }
        $this->setDefaultStyles('header');
        $this->setDefaultStyles('box');
        $this->filterModel = null;
        $this->setDefaultExportConfig();
        $this->exportConfig = ArrayHelper::merge($this->_defaultExportConfig, $this->exportConfig);
        if (!isset($this->filename)) {
            $this->filename = Yii::t('kvexport', 'grid-export');
        }
        $target = $this->target == self::TARGET_POPUP ? 'kvExportFullDialog' : $this->target;
        $id = ArrayHelper::getValue($this->exportFormOptions, 'id', $this->options['id'] . '-form');
        Html::addCssClass($this->exportFormOptions, 'kv-export-full-form');
        $this->exportFormOptions += [
            'id' => $id,
            'target' => $target,
        ];
    }

    /**
     * Renders the export menu
     *
     * @return string the export menu markup
     * @throws InvalidConfigException
     * @throws \Exception
     */
    public function renderExportMenu()
    {
        $items = $this->asDropdown ? [] : '';
        $isBs4 = $this->isBs4();
        Html::addCssClass($this->dropdownOptions, ['btn', $this->getDefaultBtnCss()]);
        foreach ($this->exportConfig as $format => $settings) {
            if (!isset($settings) || $settings === false) {
                continue;
            }
            $label = '';
            if (isset($settings['icon'])) {
                $iconOptions = ArrayHelper::getValue($settings, 'iconOptions', []);
                Html::addCssClass($iconOptions, $settings['icon']);
                $label = Html::tag('i', '', $iconOptions) . ' ';
            }
            if (isset($settings['label'])) {
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
                    'options' => $options,
                ];
            } else {
                $tag = ArrayHelper::remove($options, 'tag', 'li');
                if ($tag !== false) {
                    $items .= Html::tag($tag, Html::a($label, '#', $linkOptions), $options);
                } else {
                    $items .= Html::a($label, '#', $linkOptions);
                }
            }
        }
        $form = $this->render(
            $this->exportFormView,
            [
                'options' => $this->exportFormOptions,
                'exportType' => $this->_exportType,
                'columnSelectorEnabled' => $this->_columnSelectorEnabled,
                'exportRequestParam' => $this->exportRequestParam,
                'exportTypeParam' => self::PARAM_EXPORT_TYPE,
                'exportColsParam' => self::PARAM_EXPORT_COLS,
                'colselFlagParam' => self::PARAM_COLSEL_FLAG,
                'exportFormHiddenInputs' => $this->exportFormHiddenInputs,
            ]
        );
        $iconCss = $isBs4 ? 'fas fa-external-link-alt' : 'glyphicon glyphicon-export';
        if ($this->asDropdown) {
            $icon = ArrayHelper::remove($this->dropdownOptions, 'icon', '<i class="' . $iconCss . '"></i>');
            $label = ArrayHelper::remove($this->dropdownOptions, 'label', null);
            $label = $label === null ? $icon : $icon . ' ' . $label;
            if (!isset($this->dropdownOptions['title'])) {
                $this->dropdownOptions['title'] = Yii::t('kvexport', 'Export data in selected format');
            }
            $menuOptions = ArrayHelper::remove($this->dropdownOptions, 'menuOptions', []);
            $itemsBefore = ArrayHelper::remove($this->dropdownOptions, 'itemsBefore', []);
            $itemsAfter = ArrayHelper::remove($this->dropdownOptions, 'itemsAfter', []);
            $items = ArrayHelper::merge($itemsBefore, $items, $itemsAfter);
            $opts = [
                'label' => $label,
                'dropdown' => ['items' => $items, 'encodeLabels' => false, 'options' => $menuOptions,],
                'encodeLabel' => false,
            ];

            if (!isset($this->exportContainer['class'])) {
                $this->exportContainer['class'] = 'btn-group';
            }
            /**
             * @var Widget $class
             */
            $class = $isBs4 ? 'kartik\bs4dropdown\ButtonDropdown' : 'yii\bootstrap\ButtonDropdown';
            if (!class_exists($class)) {
                throw new InvalidConfigException("The '{$class}' does not exist and must be installed for dropdown rendering when 'ExportMenu::asDropdown' is set to 'true'.");
            }
            if ($isBs4) {
                $opts['buttonOptions'] = $this->dropdownOptions;
                $opts['renderContainer'] = false;
                $out = Html::tag('div', $class::widget($opts), $this->exportContainer);
            } else {
                $opts['options'] = $this->dropdownOptions;
                $opts['containerOptions'] = $this->exportContainer;
                $out = $class::widget($opts);
            }
            $replacePairs = ['{menu}' => $out, '{columns}' => $this->renderColumnSelector()];
            $content = strtr($this->template, $replacePairs) . "\n" . $form;
            return Html::tag('div', $content, $this->container);
        } else {
            return $items . "\n" . $form;
        }
    }

    /**
     * Renders the columns selector
     *
     * @return string the column selector markup
     */
    public function renderColumnSelector()
    {
        if (!$this->_columnSelectorEnabled) {
            return '';
        }
        return $this->render(
            $this->exportColumnsView,
            [
                'options' => $this->columnSelectorOptions,
                'menuOptions' => $this->columnSelectorMenuOptions,
                'columnSelector' => $this->columnSelector,
                'batchToggle' => $this->columnBatchToggleSettings,
                'selectedColumns' => $this->selectedColumns,
                'disabledColumns' => $this->disabledColumns,
                'hiddenColumns' => $this->hiddenColumns,
                'noExportColumns' => $this->noExportColumns,
            ]
        );
    }

    /**
     * Initializes PhpSpreadsheet Object Instance
     */
    public function initPhpSpreadsheet()
    {
        $this->_objSpreadsheet = new Spreadsheet();
        $creator = $title = $subject = $category = $keywords = $manager = '';
        $description = Yii::t('kvexport', 'Grid export generated by Krajee ExportMenu widget (yii2-export)');
        $company = 'Krajee Solutions';
        $created = date('Y-m-d H:i:s');
        $lastModifiedBy = 'krajee';
        extract($this->docProperties);
        $properties = $this->_objSpreadsheet->getProperties();
        /** @noinspection PhpParamsInspection */
        $properties->setCreator($creator)
            ->setTitle($title)
            ->setSubject($subject)
            ->setDescription($description)
            ->setCategory($category)
            ->setKeywords($keywords)
            ->setManager($manager)
            ->setCompany($company)
            ->setCreated($created)
            ->setLastModifiedBy($lastModifiedBy);
        $this->raiseEvent('onInitExcel', [$this->_objSpreadsheet, $this]);
    }

    /**
     * Initializes PhpSpreadsheet Writer Object Instance
     *
     * @param string $type the writer type as set in export config
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function initPhpSpreadsheetWriter($type)
    {
        /**
         * @var WriterCsv $writer
         */
        $writer = $this->_objWriter = IOFactory::createWriter($this->_objSpreadsheet, $type);
        if ($this->_exportType === self::FORMAT_TEXT) {
            $delimiter = $this->getSetting('delimiter', "\t");
            $writer->setDelimiter($delimiter);
        }
        $this->raiseEvent('onInitWriter', [$this->_objWriter, $this]);
    }

    /**
     * Initializes PhpSpreadsheet Worksheet Instance
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function initPhpSpreadsheetWorksheet()
    {
        $this->_objWorksheet = $this->_objSpreadsheet->getActiveSheet();
        $this->raiseEvent('onInitSheet', [$this->_objWorksheet, $this]);
    }

    /**
     * Generates the before content at the top of the exported sheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function generateBeforeContent()
    {
        $colFirst = self::columnName(1);
        $sheet = $this->_objWorksheet;
        foreach ($this->contentBefore as $contentBefore) {
            $this->setOutCellValue($sheet, $colFirst . $this->_beginRow, $contentBefore['value']);
            $opts = $this->getStyleOpts($contentBefore);
            $sheet->getStyle($colFirst . $this->_beginRow)->applyFromArray($opts);
            $this->_beginRow += 1;
        }
    }

    /**
     * Generates the output data header content.
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function generateHeader()
    {
        $columns = $this->getVisibleColumns();
        if (count($columns) == 0) {
            return;
        }
        $sheet = $this->_objWorksheet;
        $styleOpts = ArrayHelper::getValue($this->headerStyleOptions, $this->_exportType, []);
        $colFirst = self::columnName(1);

        $this->_endCol = 0;
        foreach ($this->getVisibleColumns() as $column) {
            if (!empty($column->hiddenFromExport)) {
                continue;
            }
            $opts = $styleOpts;
            $this->_endCol++;
            /**
             * @var \kartik\grid\DataColumn $column
             */
            $head = ($column instanceof DataColumn) ? $this->getColumnHeader($column) : $column->header;
            $id = self::columnName($this->_endCol) . $this->_beginRow;
            $cell = $this->setOutCellValue($sheet, $id, $head);
            if (isset($column->hAlign) && !isset($opts['alignment']['horizontal'])) {
                $opts['alignment']['horizontal'] = $column->hAlign;
            }
            if (isset($column->vAlign) && !isset($opts['alignment']['vertical'])) {
                $opts['alignment']['vertical'] = $column->vAlign;
            }
            // Apply formatting to header cell
            $sheet->getStyle($id)->applyFromArray($opts);
            $this->raiseEvent('onRenderHeaderCell', [$cell, $head, $this]);
        }
        for ($i = $this->_headerBeginRow; $i < ($this->_beginRow); $i++) {
            $sheet->mergeCells($colFirst . $i . ':' . self::columnName($this->_endCol) . $i);
        }

        // Freeze the top row
        $sheet->freezePane($colFirst . ($this->_beginRow + 1));
    }

    /**
     * Gets the visible columns for export
     *
     * @return array the columns configuration
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
    public function setVisibleColumns()
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
     * Gets the column header content
     *
     * @param DataColumn $col
     *
     * @return string
     */
    public function getColumnHeader($col)
    {
        if ($col->header !== null || ($col->label === null && $col->attribute === null)) {
            return trim($col->header) !== '' ? $col->header : $col->grid->emptyCell;
        }
        $provider = $this->dataProvider;
        if ($col->label === null) {
            if ($provider instanceof ActiveDataProvider && $provider->query instanceof ActiveQueryInterface) {
                /**
                 * @var \yii\db\ActiveRecord $model
                 */
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
     * @return integer the number of output rows.
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function generateBody()
    {
        $this->_endRow = 0;
        $columns = $this->getVisibleColumns();
        $models = array_values($this->_provider->getModels());
        if (count($columns) == 0) {
            $cell = $this->setOutCellValue($this->_objWorksheet, 'A1', $this->emptyText);
            $model = reset($models);
            $this->raiseEvent('onRenderDataCell', [$cell, $this->emptyText, $model, null, 0, $this]);
            return 0;
        }
        // do not execute multiple COUNT(*) queries
        $totalCount = $this->_provider->getTotalCount();
        $this->findGroupedColumn();
        while (count($models) > 0) {
            $keys = $this->_provider->getKeys();
            foreach ($models as $index => $model) {
                $key = $keys[$index];
                $this->generateRow($model, $key, $this->_endRow);
                $this->_endRow++;
                if ($index === $totalCount) {
                    //a little hack to generate last grouped footer
                    $this->checkGroupedRow($model, $models[0], $key, $this->_endRow);
                } elseif (isset($models[$index + 1])) {
                    $this->checkGroupedRow($model, $models[$index + 1], $key, $this->_endRow);
                }
                if (!is_null($this->_groupedRow)) {
                    $this->_endRow++;
                    $this->_objWorksheet->fromArray($this->_groupedRow, null, 'A' . ($this->_endRow + 1), true);
                    $cell = 'A' . ($this->_endRow + 1) . ':' . self::columnName(count($columns)) . ($this->_endRow + 1);
                    $this->_objWorksheet->getStyle($cell)->applyFromArray($this->groupedRowStyle);
                    $this->_groupedRow = null;
                }
            }
            if ($this->_provider->pagination) {
                $this->_provider->pagination->page++;
                $this->_provider->refresh();
                $this->_provider->setTotalCount($totalCount);
                $models = $this->_provider->getModels();
            } else {
                $models = [];
            }
        }
        $this->generateBox();
        return $this->_endRow;
    }

    /**
     * Generates an output data row with the given data model and key.
     *
     * @param mixed $model the data model to be rendered
     * @param mixed $key the key associated with the data model
     * @param integer $index the zero-based index of the data model among the model array returned by [[dataProvider]].
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function generateRow($model, $key, $index)
    {
        /**
         * @var Column $column
         */
        $this->_endCol = 0;
        foreach ($this->getVisibleColumns() as $column) {
            if (!empty($column->hiddenFromExport)) {
                continue;
            }
            $format = $this->enableFormatter && isset($column->format) ? $column->format : 'raw';
            $value = null;
            if ($column instanceof ActionColumn) {
                $value = null;
            } elseif ($column instanceof SerialColumn) {
                $value = $index + 1;
                $pagination = $column->grid->dataProvider->getPagination();
                if ($pagination !== false) {
                    $value += $pagination->getOffset();
                }
            } elseif (isset($column->content)) {
                $value = call_user_func($column->content, $model, $key, $index, $column);
            } elseif (method_exists($column, 'getDataCellValue')) {
                $value = $column->getDataCellValue($model, $key, $index);
            } elseif (isset($column->attribute)) {
                $value = ArrayHelper::getValue($model, $column->attribute, '');
            }
            $this->_endCol++;
            if (isset($value) && $value !== '' && isset($format)) {
                $value = $this->formatter->format($value, $format);
            } else {
                $value = '';
            }
            $cell = $this->setOutCellValue(
                $this->_objWorksheet,
                self::columnName($this->_endCol) . ($index + $this->_beginRow + 1),
                $value
            );
            if ($this->enableAutoFormat) {
                $this->autoFormat($model, $key, $index, $column, $cell);
            }
            $this->raiseEvent('onRenderDataCell', [$cell, $value, $model, $key, $index, $this]);
        }
    }

    /**
     * Generates the output footer row after a specific row number
     *
     * @return integer the row number after which the footer is to be generated
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function generateFooter()
    {
        $row = $this->_endRow + $this->_beginRow;
        $footerExists = false;
        $columns = $this->getVisibleColumns();
        if (count($columns) == 0) {
            return 0;
        }
        $this->_endCol = 0;
        foreach ($this->getVisibleColumns() as $n => $column) {
            $this->_endCol = $this->_endCol + 1;
            if ($column->footer) {
                $footerExists = true;
                $footer = trim($column->footer) !== '' ? $column->footer : $column->grid->blankDisplay;
                $cell = $this->setOutCellValue(
                    $this->_objSpreadsheet->getActiveSheet(),
                    self::columnName($this->_endCol) . ($row + 1),
                    $footer
                );
                $this->raiseEvent('onRenderFooterCell', [$cell, $footer, $this]);
            }
        }
        if ($footerExists) {
            $row++;
        }
        return $row;
    }

    /**
     * Generates the after content at the bottom of the exported sheet
     *
     * @param integer $row the row number after which the content is to be generated
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function generateAfterContent($row)
    {
        $colFirst = self::columnName(1);
        $row++;
        $afterContentBeginRow = $row;
        $sheet = $this->_objWorksheet;
        foreach ($this->contentAfter as $contentAfter) {
            $this->setOutCellValue($sheet, $colFirst . $row, $contentAfter['value']);
            $opts = $this->getStyleOpts($contentAfter);
            $sheet->getStyle($colFirst . $row)->applyFromArray($opts);
            $row += 1;
        }
        for ($i = $afterContentBeginRow; $i < $row; $i++) {
            $sheet->mergeCells($colFirst . $i . ':' . self::columnName($this->_endCol) . $i);
        }
    }

    /**
     * Gets the currently selected export type
     *
     * @return string
     */
    public function getExportType()
    {
        return $this->_exportType;
    }

    /**
     * Gets the PhpSpreadsheet object
     *
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet the current \PhpOffice\PhpSpreadsheet\Spreadsheet object instance
     */
    public function getPhpSpreadsheet()
    {
        return $this->_objSpreadsheet;
    }

    /**
     * Gets the PhpSpreadsheet writer object
     *
     * @return \PhpOffice\PhpSpreadsheet\Writer\BaseWriter the current \PhpOffice\PhpSpreadsheet\Writer\BaseWriter object instance
     */
    public function getPhpSpreadsheetWriter()
    {
        return $this->_objWriter;
    }

    /**
     * Gets the PhpSpreadsheet sheet object
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet the current \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet object instance
     */
    public function getPhpSpreadsheetWorksheet()
    {
        return $this->_objWorksheet;
    }

    /**
     * Sets the PhpSpreadsheet object
     *
     * @param $obj \PhpOffice\PhpSpreadsheet\Spreadsheet the \PhpOffice\PhpSpreadsheet\Spreadsheet object instance
     */
    public function setPhpSpreadsheet($obj)
    {
        $this->_objSpreadsheet = $obj;
    }

    /**
     * Sets the PhpSpreadsheet writer object
     *
     * @param $obj \PhpOffice\PhpSpreadsheet\Writer\BaseWriter the \PhpOffice\PhpSpreadsheet\Writer\BaseWriter object instance
     */
    public function setPhpSpreadsheetWriter($obj)
    {
        $this->_objWriter = $obj;
    }

    /**
     * Sets the PhpSpreadsheet sheet object
     *
     * @param $obj \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet the \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet object instance
     */
    public function setPhpSpreadsheetWorksheet($obj)
    {
        $this->_objWorksheet = $obj;
    }

    /**
     * Destroys PhpSpreadsheet Object Instance
     */
    public function destroyPhpSpreadsheet()
    {
        if (isset($this->_objSpreadsheet)) {
            $this->_objSpreadsheet->disconnectWorksheets();
        }
        unset($this->_provider, $this->_objWriter, $this->_objWorksheet, $this->_objSpreadsheet);
    }

    /**
     * Sets default styles
     *
     * @param string $section the php spreadsheet section
     */
    protected function setDefaultStyles($section)
    {
        $defaultStyle = [];
        $opts = '';
        if ($section === 'header') {
            $opts = 'headerStyleOptions';
            $defaultStyle = [
                'font' => ['bold' => true],
                'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'color' => [
                        'argb' => 'FFE5E5E5',
                    ],
                ],
                'borders' => [
                    'outline' => [
                        'borderStyle' => Border::BORDER_MEDIUM,
                        'color' => ['argb' => Color::COLOR_BLACK],
                    ],
                    'inside' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => ['argb' => Color::COLOR_BLACK],
                    ],
                ],
            ];
        } elseif ($section === 'box') {
            $opts = 'boxStyleOptions';
            $defaultStyle = [
                'borders' => [
                    'outline' => [
                        'borderStyle' => Border::BORDER_MEDIUM,
                        'color' => ['argb' => Color::COLOR_BLACK],
                    ],
                    'inside' => [
                        'borderStyle' => Border::BORDER_DOTTED,
                        'color' => ['argb' => Color::COLOR_BLACK],
                    ],
                ],
            ];
        }
        if (empty($opts)) {
            return;
        }
        $defaultStyleOptions = [
            self::FORMAT_HTML => $defaultStyle,
            self::FORMAT_PDF => $defaultStyle,
            self::FORMAT_EXCEL => $defaultStyle,
            self::FORMAT_EXCEL_X => $defaultStyle,
        ];
        $this->$opts = array_replace_recursive($defaultStyleOptions, $this->$opts);
    }

    /**
     * Generates the box
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function generateBox()
    {
        // Set autofilter on
        $from = self::columnName(1) . $this->_beginRow;
        $to = self::columnName($this->_endCol) . ($this->_endRow + $this->_beginRow);
        $box = "{$from}:{$to}";
        $this->_objWorksheet->setAutoFilter($box);
        if (isset($this->boxStyleOptions[$this->_exportType])) {
            $this->_objWorksheet->getStyle($box)->applyFromArray($this->boxStyleOptions[$this->_exportType]);
        }

        if (isset($this->headerStyleOptions[$this->_exportType])) {
            $to = self::columnName($this->_endCol) . $this->_beginRow;
            $box = "{$from}:{$to}";
            $this->_objWorksheet->getStyle($box)->applyFromArray($this->headerStyleOptions[$this->_exportType]);
        }
    }

    /**
     * Autoformats a cell by auto detecting the grid column alignment and format
     *
     * @param mixed $model the data model to be rendered
     * @param mixed $key the key associated with the data model
     * @param integer $index the zero-based index of the data model among the model array returned by [[dataProvider]].
     * @param Column $column
     * @param Cell $cell
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function autoFormat($model, $key, $index, $column, $cell)
    {
        $ord = $cell->getCoordinate();
        $style = $this->_objWorksheet->getStyle($ord);
        $opts = ArrayHelper::getValue($this->styleOptions, $this->_exportType, []);
        if (isset($column->exportMenuStyle)) {
            $opts = $column->exportMenuStyle;
            if ($opts instanceof Closure) {
                $opts = call_user_func($opts, $model, $key, $index, $column);
            }
        }
        if (isset($column->hAlign) && !isset($opts['alignment']['horizontal'])) {
            $opts['alignment']['horizontal'] = $column->hAlign;
        }
        if (isset($column->vAlign) && !isset($opts['alignment']['vertical'])) {
            $opts['alignment']['vertical'] = $column->vAlign;
        }
        if (isset($column->format) && !isset($opts['numberFormat']) && !($column->format instanceof Closure)) {
            $fmt = (array)$column->format;
            $f = $fmt[0];
            $code = null;
            if ($f === 'integer') {
                $code = '0';
            } elseif ($f === 'percent' || $f === 'decimal' || $f === 'currency') {
                $code = '';
                if ($f === 'currency') {
                    $code = ArrayHelper::getValue($fmt, 1, $this->formatter->currencyCode) . ' ';
                }
                $decimals = ArrayHelper::getValue($fmt, 1, ($f === 'percent' ? 0 : 2));
                $d = intval($decimals);
                $code .= '#' . $this->formatter->thousandSeparator . '##0';
                if ($d > 0) {
                    $code .= $this->formatter->decimalSeparator . str_repeat('0', $d);
                }
                if ($f === 'percent') {
                    $code .= '%';
                }
            }
            if ($code !== null) {
                $opts['numberFormat'] = ['formatCode' => $code];
            }
        }
        $style->applyFromArray($opts);
    }

    /**
     * Gets the setting property value for the current export format
     *
     * @param string $key the setting property key for the current export format
     * @param string $default the default value for the property
     *
     * @return mixed
     */
    protected function getSetting($key, $default = null)
    {
        $settings = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
        return ArrayHelper::getValue($settings, $key, $default);
    }

    /**
     * Parse PDF
     *
     * @param string $file the output filename on server with path
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function renderPDF($file)
    {
        //  Default PDF paper size
        $spreadsheet = $this->_objSpreadsheet;
        $sheet = $this->_objWorksheet;
        /**
         * @var WriterHtml $w
         */
        $w = $this->_objWriter;
        $page = $sheet->getPageSetup();
        $orientation = $page->getOrientation() == PageSetup::ORIENTATION_LANDSCAPE ? 'L' : 'P';
        $properties = $spreadsheet->getProperties();
        $settings = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
        $useInlineCss = ArrayHelper::getValue($settings, 'useInlineCss', false);
        $config = ArrayHelper::getValue($settings, 'pdfConfig', []);
        $w->setUseInlineCss($useInlineCss);
        $config = array_replace_recursive(
            [
                'orientation' => strtoupper($orientation),
                'methods' => [
                    'SetTitle' => $properties->getTitle(),
                    'SetAuthor' => $properties->getCreator(),
                    'SetCreator' => $properties->getCreator(),
                    'SetSubject' => $properties->getSubject(),
                    'SetKeywords' => $properties->getKeywords(),
                ],
                'cssFile' => '',
                'content' => $w->generateHTMLHeader(false) . $w->generateSheetData() . $w->generateHTMLFooter(),
            ],
            $config
        );
        if (!$this->stream) {
            $config['destination'] = Pdf::DEST_FILE;
            $config['filename'] = $file;
        } else {
            $config['destination'] = Pdf::DEST_DOWNLOAD;
            $extension = ArrayHelper::getValue($settings, 'extension', 'pdf');
            $config['filename'] = $this->filename . '.' . $extension;
        }
        $pdf = new Pdf($config);
        echo $pdf->render();
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
        if (!isset($_POST[self::PARAM_EXPORT_COLS]) or $_POST[self::PARAM_EXPORT_COLS] === '') {
            return;
        }
        $this->selectedColumns = Json::decode($_POST[self::PARAM_EXPORT_COLS]);
    }

    /**
     * Clear output buffers
     */
    protected function clearOutputBuffers()
    {
        if ($this->clearBuffers) {
            while (ob_get_level() > 0) {
                ob_end_clean();
            }
        } else {
            ob_end_clean();
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
        Html::addCssClass($this->columnSelectorOptions, ['btn', $this->getDefaultBtnCss(), 'dropdown-toggle']);
        $header = ArrayHelper::getValue($this->columnSelectorOptions, 'header', Yii::t('kvexport', 'Select Columns'));
        $this->columnSelectorOptions['header'] = (!isset($header) || $header === false) ? '' :
            '<li class="dropdown-header">' . $header . '</li><li class="kv-divider"></li>';
        $id = $this->options['id'] . '-cols';
        Html::addCssClass($this->columnSelectorMenuOptions, 'dropdown-menu kv-checkbox-list');
        $this->columnSelectorMenuOptions = array_replace_recursive(
            [
                'id' => $id . '-list',
                'role' => 'menu',
                'aria-labelledby' => $id,
            ],
            $this->columnSelectorMenuOptions
        );
        $this->columnSelectorOptions = array_replace_recursive(
            [
                'id' => $id,
                'icon' => '<i class="glyphicon glyphicon-list"></i>',
                'title' => Yii::t('kvexport', 'Select columns to export'),
                'type' => 'button',
                'data-toggle' => 'dropdown',
                'aria-haspopup' => 'true',
                'aria-expanded' => 'false',
            ],
            $this->columnSelectorOptions
        );
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
     * @param integer $key
     * @param Column $column
     *
     * @return string
     */
    protected function getColumnLabel($key, $column)
    {
        $key++;
        $label = Yii::t('kvexport', 'Column') . ' ' . $key;
        if (isset($column->label)) {
            $label = $column->label;
        } elseif (isset($column->header)) {
            $label = $column->header;
        } elseif (isset($column->attribute)) {
            $label = $this->getAttributeLabel($column->attribute);
        } elseif (!$column instanceof DataColumn) {
            $class = explode('\\', get_class($column));
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
        /**
         * @var Model $model
         */
        $provider = $this->dataProvider;
        if ($provider instanceof ActiveDataProvider && $provider->query instanceof ActiveQueryInterface) {
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
     * Sets the default export configuration
     * @throws InvalidConfigException
     */
    protected function setDefaultExportConfig()
    {
        $isFa = $this->fontAwesome;
        $isBs4 = $this->isBs4();
        $this->_defaultExportConfig = [
            self::FORMAT_HTML => [
                'label' => Yii::t('kvexport', 'HTML'),
                'icon' => $isBs4 ? 'fas fa-file-alt' : ($isFa ? 'fa fa-file-text' : 'glyphicon glyphicon-save'),
                'iconOptions' => ['class' => 'text-info'],
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Hyper Text Markup Language')],
                'alertMsg' => Yii::t('kvexport', 'The HTML export file will be generated for download.'),
                'mime' => 'text/html',
                'extension' => 'html',
                'writer' => self::FORMAT_HTML,
            ],
            self::FORMAT_CSV => [
                'label' => Yii::t('kvexport', 'CSV'),
                'icon' => $isBs4 ? 'fas fa-file-code' : ($isFa ? 'fa fa-file-code-o' : 'glyphicon glyphicon-floppy-open'),
                'iconOptions' => ['class' => 'text-primary'],
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Comma Separated Values')],
                'alertMsg' => Yii::t('kvexport', 'The CSV export file will be generated for download.'),
                'mime' => 'application/csv',
                'extension' => 'csv',
                'writer' => self::FORMAT_CSV,
            ],
            self::FORMAT_TEXT => [
                'label' => Yii::t('kvexport', 'Text'),
                'icon' => $isBs4 ? 'far fa-file-alt' : ($isFa ? 'fa fa-file-text-o' : 'glyphicon glyphicon-floppy-save'),
                'iconOptions' => ['class' => 'text-muted'],
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Tab Delimited Text')],
                'alertMsg' => Yii::t('kvexport', 'The TEXT export file will be generated for download.'),
                'mime' => 'text/plain',
                'extension' => 'txt',
                'writer' => self::FORMAT_CSV,
                'delimiter' => "\t",
            ],
            self::FORMAT_PDF => [
                'label' => Yii::t('kvexport', 'PDF'),
                'icon' => $isBs4 ? 'far fa-file-pdf' : ($isFa ? 'fa fa-file-pdf-o' : 'glyphicon glyphicon-floppy-disk'),
                'iconOptions' => ['class' => 'text-danger'],
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Portable Document Format')],
                'alertMsg' => Yii::t('kvexport', 'The PDF export file will be generated for download.'),
                'mime' => 'application/pdf',
                'extension' => 'pdf',
                'writer' => self::FORMAT_HTML,
                'useInlineCss' => true,
                'pdfConfig' => [],
            ],
            self::FORMAT_EXCEL => [
                'label' => Yii::t('kvexport', 'Excel 95 +'),
                'icon' => $isBs4 ? 'far fa-file-excel' : ($isFa ? 'fa fa-file-excel-o' : 'glyphicon glyphicon-floppy-remove'),
                'iconOptions' => ['class' => 'text-success'],
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Microsoft Excel 95+ (xls)')],
                'alertMsg' => Yii::t('kvexport', 'The EXCEL 95+ (xls) export file will be generated for download.'),
                'mime' => 'application/vnd.ms-excel',
                'extension' => 'xls',
                'writer' => self::FORMAT_EXCEL,
            ],
            self::FORMAT_EXCEL_X => [
                'label' => Yii::t('kvexport', 'Excel 2007+'),
                'icon' => $isBs4 ? 'fas fa-file-excel' : ($isFa ? 'fa fa-file-excel-o' : 'glyphicon glyphicon-floppy-remove'),
                'iconOptions' => ['class' => 'text-success'],
                'linkOptions' => [],
                'options' => ['title' => Yii::t('kvexport', 'Microsoft Excel 2007+ (xlsx)')],
                'alertMsg' => Yii::t('kvexport', 'The EXCEL 2007+ (xlsx) export file will be generated for download.'),
                'mime' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'extension' => 'xlsx',
                'writer' => self::FORMAT_EXCEL_X,
            ],
        ];
    }

    /**
     * Registers client assets needed for Export Menu widget
     * @throws \Exception
     */
    protected function registerAssets()
    {
        $view = $this->getView();
        Dialog::widget($this->krajeeDialogSettings);
        ExportMenuAsset::register($view);
        $this->messages += [
            'allowPopups' => Yii::t(
                'kvexport',
                'Disable any popup blockers in your browser to ensure proper download.'
            ),
            'confirmDownload' => Yii::t('kvexport', 'Ok to proceed?'),
            'downloadProgress' => Yii::t('kvexport', 'Generating the export file. Please wait...'),
            'downloadComplete' => Yii::t(
                'kvexport',
                'Request submitted! You may safely close this dialog after saving your downloaded file.'
            ),
        ];
        $formId = $this->exportFormOptions['id'];
        $options = Json::encode(
            [
                'formId' => $formId,
                'messages' => $this->messages,
                'dialogLib' => new JsExpression(
                    ArrayHelper::getValue($this->krajeeDialogSettings, 'libName', 'krajeeDialog')
                ),
            ]
        );
        $menu = 'kvexpmenu_' . hash('crc32', $options);
        $view->registerJs("var {$menu} = {$options};\n", View::POS_HEAD);
        $script = '';
        foreach ($this->exportConfig as $format => $setting) {
            if (!isset($setting) || $setting === false) {
                continue;
            }
            $id = $this->options['id'] . '-' . strtolower($format);
            $options = [
                'settings' => new JsExpression($menu),
                'alertMsg' => $setting['alertMsg'],
                'target' => $this->target,
                'showConfirmAlert' => $this->showConfirmAlert,
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
     * Raises a callable event
     *
     * @param string $event the event name
     * @param array $params the parameters to the callable function
     *
     * @return mixed
     */
    protected function raiseEvent($event, $params)
    {
        if (isset($this->$event) && is_callable($this->$event)) {
            return call_user_func_array($this->$event, $params);
        }
        return true;
    }

    /**
     * Parses and returns the style options for `contentBefore` or `contentAfter`
     *
     * @param array $settings the settings to parse (for `contentBefore` or `contentAfter`)
     *
     * @return array
     */
    protected function getStyleOpts($settings = [])
    {
        $styleOpts = ArrayHelper::getValue($settings, 'styleOptions', []);
        return ArrayHelper::getValue($styleOpts, $this->_exportType, []);
    }

    /**
     * Search all groupable columns
     */
    protected function findGroupedColumn()
    {
        foreach ($this->getVisibleColumns() as $key => $column) {
            if (isset($column->group) && $column->group == true) {
                $this->_groupedColumn[$key] = ['firstLine' => -1, 'value' => null];
            } else {
                $this->_groupedColumn[$key] = null;
            }
        }
        $this->_groupedColumn[] = null; //prevent the overflow
        $this->_groupedColumn[] = null; //prevent the overflow
    }

    /**
     * Validates a grouped row
     *
     * @param Model|array $model the data model
     * @param Model|array $nextModel the next data model
     * @param integer $key the key associated with the data model
     * @param integer $index the zero-based index of the data model among the model array returned by
     * [[dataProvider]].
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function checkGroupedRow($model, $nextModel, $key, $index)
    {
        $endCol = 0;
        /**
         * @var Column $column
         */
        foreach ($this->getVisibleColumns() as $column) {
            if ((isset($this->_groupedColumn[$endCol])) && (!is_null($this->_groupedColumn[$endCol]))) {
                $value = ($column->content === null) ? (method_exists($column, 'getDataCellValue') ?
                    $this->formatter->format($column->getDataCellValue($model, $key, $index), 'raw') :
                    $column->renderDataCell($model, $key, $index)) :
                    call_user_func($column->content, $model, $key, $index, $column);
                $nextValue = ($column->content === null) ? (method_exists($column, 'getDataCellValue') ?
                    $this->formatter->format($column->getDataCellValue($nextModel, $key, $index), 'raw') :
                    $column->renderDataCell($nextModel, $key, $index)) :
                    call_user_func($column->content, $nextModel, $key, $index, $column);
                if (is_null($this->_groupedColumn[$endCol]['value'])) {
                    $this->_groupedColumn[$endCol]['value'] = $value;
                    $this->_groupedColumn[$endCol]['firstLine'] = $index;
                }
                if ($this->_groupedColumn[$endCol]['value'] != $nextValue) {
                    $groupFooter = isset($column->groupFooter) ? $column->groupFooter : null;
                    if ($groupFooter instanceof Closure) {
                        $groupFooter = call_user_func($groupFooter, $model, $key, $index, $this);
                    }
                    if (isset($groupFooter['content'])) {
                        $this->generateGroupedRow($groupFooter['content'], $endCol);
                    }
                    $this->_groupedColumn[$endCol]['firstLine'] = $index;
                }
                $this->_groupedColumn[$endCol]['value'] = $nextValue;
            }
            $endCol++;
        }
    }

    /**
     * Generate a grouped row
     *
     * @param array $groupFooter footer row
     * @param integer $groupedCol the zero-based index of grouped column
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function generateGroupedRow($groupFooter, $groupedCol)
    {
        $endGroupedCol = 0;
        $this->_groupedRow = [];
        $fLine = ArrayHelper::getValue($this->_groupedColumn[$groupedCol], 'firstLine', -1);
        $fLine = ($fLine == $this->_beginRow) ? $this->_beginRow + 1 : ($fLine + 3);
        $firstLine = ($this->_endRow == ($this->_beginRow + 3) && $fLine == 2) ? $this->_beginRow + 3 : $fLine;
        $endLine = $this->_endRow + 1;
        list($endLine, $firstLine) = ($endLine > $firstLine) ? [$endLine, $firstLine] : [$firstLine, $endLine];
        foreach ($this->getVisibleColumns() as $key => $column) {
            $value = isset($groupFooter[$key]) ? $groupFooter[$key] : '';
            $endGroupedCol++;
            $groupedRange = self::columnName($key + 1) . $firstLine . ':' . self::columnName($key + 1) . $endLine;
            //$lastCell = self::columnName($key + 1) . $endLine - 1;
            if (isset($column->group) && $column->group) {
                $this->_objWorksheet->mergeCells($groupedRange);
            }
            switch ($value) {
                case self::F_SUM:
                    $value = "=sum($groupedRange)";
                    break;
                case self::F_COUNT:
                    $value = '=countif(' . $groupedRange . ',"*")';
                    break;
                case self::F_AVG:
                    $value = "=AVERAGE($groupedRange)";
                    break;
                case self::F_MAX:
                    $value = "=max($groupedRange)";
                    break;
                case self::F_MIN:
                    $value = "=min($groupedRange)";
                    break;
            }
            if ($value instanceof \Closure) {
                $value = call_user_func($value, $groupedRange, $this);
            }
            $this->_groupedRow[] = !isset($value) || $value === '' ? '' : strip_tags($value);
        }
    }

    /**
     * Set HTTP headers for download
     */
    protected function setHttpHeaders()
    {
        $config = ArrayHelper::getValue($this->exportConfig, $this->_exportType, []);
        $extension = ArrayHelper::getValue($config, 'extension', 'xlsx');
        $mime = ArrayHelper::getValue($config, 'mime', null);
        if (strstr($_SERVER['HTTP_USER_AGENT'], 'MSIE') == false) {
            header('Cache-Control: no-cache');
            header('Pragma: no-cache');
        } else {
            header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
            header('Pragma: public');
        }
        header('Expires: Sat, 26 Jul 1979 05:00:00 GMT');
        header("Content-Encoding: {$this->encoding}");
        if (!empty($mime)) {
            header("Content-Type: {$mime}; charset={$this->encoding}");
        }
        header("Content-Disposition: attachment; filename=\"{$this->filename}.{$extension}\"");
        header('Cache-Control: max-age=0');
    }

    /**
     * Parses format and sets the value of a PHP Spreadsheet Cell
     *
     * @param Worksheet $sheet
     * @param string $index coordinate of the cell, eg: 'A1'
     * @param mixed $value value of the cell
     *
     * @return Cell
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function setOutCellValue($sheet, $index, $value)
    {
        if ($this->stripHtml) {
            $value = strip_tags($value);
        }
        $value = html_entity_decode($value, ENT_QUOTES, 'UTF-8');
        $cell = $sheet->getCell($index);
        $cell->setValue($value);
        return $cell;
    }

    /**
     * Cleans up the export file and current object instance
     *
     * @param string $file the file exported
     * @param array $config the export configuration
     */
    protected function cleanup($file, $config)
    {
        if ($this->raiseEvent('onGenerateFile', [$config['extension'], $this]) === false) {
            return;
        }
        if ($this->stream || $this->deleteAfterSave) {
            @unlink($file);
        }
        $this->destroyPhpSpreadsheet();
    }
}
