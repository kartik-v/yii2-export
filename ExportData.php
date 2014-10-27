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
use yii\base\InvalidConfigException;
use yii\base\InvalidValueException;
use yii\helpers\Html;
use yii\helpers\Json;
use yii\helpers\Inflector;
use yii\helpers\ArrayHelper;
use yii\data\DataProvider;
use yii\grid\GridView;
use yii\bootstrap\ButtonDropdown;

/**
 * Export data library. 
 *
 * @author Kartik Visweswaran <kartikv2@gmail.com>
 * @since 1.0
 */
class ExportData extends GridView
{
    const FORMAT_HTML = 'HTML';
    const FORMAT_CSV = 'CSV';
    const FORMAT_TXT = 'TXT';
    const FORMAT_PDF = 'PDF';
    const FORMAT_EXCEL = 'Excel5';
    const FORMAT_EXCEL_X = 'Excel2007';
    
    /**
     * @var string the data output format type. Defaults to `ExportData::FORMAT_EXCEL_X`.
     */
    public $exportType = self::FORMAT_EXCEL_X;
     
    /**
     * @var bool whether to render the export menu as bootstrap button dropdown widget. Defaults to `true`.
     * If set to `false`, this will generate a simple HTML list of links.
     */
    public $asButtonDropdown = true;
    
    /**
     * @var array the HTML attributes for the export button menu. Applicable only if `isButtonDropdown`
     * is set to `true`. The following special options are available:
     * - label: string, defaults to `<i class="glyphicon glyphicon-export"></i>'
     * - menuOptions: array, the HTML attributes for the dropdown menu.
     */
    public $buttonDropdownOptions = ['class'=>'btn btn-default'];
    
    /**
     * @var array the output style configuration options. It must be the style configuration 
     * array as required by PHPExcel.
     */
    public $styleOptions = [];
    
   /**
     * @var bool whether to autosize the excel output column widths. Defaults to `true`.
     */
    public $autoWidth = true;

    /**
     * @var string the exported output file name. Defaults to 'grid-export';
     */
    public $filename;
    
    /**
     * @var bool whether to stream output to the browser
     */
    public $stream = true;   
    
    /**
     * @var Closure the callback function on rendering the header cell output content. The anonymous function 
     * should have the following signature:
     *
     * ```php
     * function ($cell, $content, $grid)
     * ```
     *
     * - `$cell`: the current PHPExcel cell being rendered
     * - `$content`: the header cell content being rendered
     * - `$grid`: the GridView object
     */
    public $onRenderHeaderCell = null;
   
    /**
     * @var Closure the callback function on rendering the data cell output content. The anonymous function 
     * should have the following signature:
     *
     * ```php
     * function ($cell, $content, $grid)
     * ```
     *
     * - `$cell`: the current PHPExcel cell being rendered
     * - `$content`: the data cell content being rendered
     * - `$grid`: the GridView object
     */    
    public $onRenderDataCell = null;
   
    /**
     * @var Closure the callback function on rendering the footer cell output content. The anonymous function 
     * should have the following signature:
     *
     * ```php
     * function ($cell, $content, $grid)
     * ```
     *
     * - `$cell`: the current PHPExcel cell being rendered
     * - `$content`: the footer cell content being rendered
     * - `$grid`: the GridView object
     */    
    public $onRenderFooterCell = null;

    /**
     * @var array the PHPExcel document properties
     */
    public $docProperties = [
        'creator' => '',
        'title' => '',
        'subject' => '',
        'description' => '',
        'category' => '',
        'keywords' => '',
        'lastModifiedBy' => ''
    ];
    
    /**
     * @var array the export configuration
     */
    public $exportConfig = [];
    
    /**
     * @var string the request parameter ($_GET or $_POST) that
     * will be submitted during export
     */
    public $exportRequestParam = 'exportFull';
    
    /**
     * @var array the the internalization configuration for this module
     */
    public $i18n = [];

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
     * @var PHPExcel object instance
     */
    protected $_objWriter;
    
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
    
    /**
     * @inherit doc
     */
    public function run() 
    {
        $triggerDownload = !empty($_POST) && !empty($_POST[$this->exportRequestParam]) && $_POST[$this->exportRequestParam];
        $this->initI18N();
        $this->initExport();
        $this->registerAssets();
        echo $this->renderExportMenu();
        if (!$triggerDownload) {
            return;
        }
        $this->initPHPExcel();
        $this->generateHeader();
        $row = $this->generateBody();
        $this->generateFooter($row);
        if ($this->autoWidth) {
            foreach ($this->columns as $n => $column) {
                $this->_objPHPExcel->getActiveSheet()->getColumnDimension(self::columnName($n + 1))->setAutoSize(true);
            }
        }
        if (empty($this->exportConfig[$this->exportType]['writer'])) { 
            return;
        }
        $objWriter = PHPExcel_IOFactory::createWriter($this->_objPHPExcel, $this->exportConfig[$this->exportType]['writer']);
        if (!$this->stream) {
            $objWriter->save($this->filename);
        }
        else {
            ob_end_clean();
            $this->setHttpHeaders();
            $objWriter->save('php://output');
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
     * @return string the export menu markup
     */
    public function renderExportMenu()
    {
        $items = $this->asButtonDropdown ? [] : '';
        foreach ($this->exportConfig as $format => $settings) {
            $icon = '';
            $label = '';
            if (!empty($settings['icon'])) {
                $label = "<i class='glyphicon glyphicon-{$icon}'></i> ";
            }
            if (!empty($settings['label'])) {
                $label .= $settings['label'];
            }
            $fmt = strtolower($format);
            $linkOptions = ArrayHelper::getValue($settings, 'linkOptions', []);
            $linkOptions['id'] = $this->options['id'] . '-' . $fmt;
            $options = ArrayHelper::getValue($settings, 'options', []);
            Html::addCssClass($linkOptions, "export-full-{$fmt}");
            if ($this->asButtonDropdown) {
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
            'target' => '_blank',
            'data-pjax' => false,
            'id'=>$this->options['id'] . '-form'
        ]) . Html::hiddenInput($this->exportRequestParam, 1) . '</form>';
        
        if ($this->asButtonDropdown) {
            $title = ArrayHelper::remove($this->buttonDropdownOptions, 'label', '<i class="glyphicon glyphicon-export"></i>');
            $menuOptions = ArrayHelper::remove($this->buttonDropdownOptions, 'menuOptions', []);
            return ButtonDropdown::widget([
                'label' => $title,
                'dropdown' => ['items' => $items, 'encodeLabels' => false, 'options'=>$menuOptions],
                'options' => $this->buttonDropdownOptions,
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
     * @return PHPExcel the object instance
     */
    public function getPHPExcel()
    {
        return $this->_objPHPExcel();
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
        $description = '';
        $category = '';
        $keywords = '';
        $lastModifiedBy = '';
        extract($this->docProperties);
        $this->_objPHPExcel->getProperties()
            ->setCreator($creator)
            ->setTitle($title)
            ->setSubject($subject)
            ->setDescription($description)
            ->setCategory($category)
            ->setKeywords($keywords)
            ->setLastModifiedBy($lastModifiedBy);
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
    public function generateHeader() {
        $cells = [];
        foreach ($this->columns as $column) {
            /* @var $column Column */
            $cells[] = $column->renderHeaderCell();
        }
        $content = Html::tag('tr', implode('', $cells), $this->headerRowOptions);
        $sheet = $this->_objPHPExcel->getActiveSheet();
        $style = $this->styleOptions[$this->exportType];
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
            $this->raiseEvent('onRenderHeaderCell', $cell, $head);
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
    public function generateBody() {
        $models = array_values($this->_provider->getModels());
        $keys = $this->_provider->getKeys();
        $this->_endRow = 0;
        foreach ($models as $index => $model) {
            $key = $keys[$index];
            $this->generateRow($model, $key, $index);
            $this->_endRow++;
        }
        // Set autofilter on
        $this->_objPHPExcel->getActiveSheet()->setAutoFilter(
            self::columnName(1).
            $this->_beginRow.
            ":".
            self::columnName($this->_endCol).
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
    public function generateRow($model, $key, $index) {
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
            $cell = $this->_objPHPExcel->getActiveSheet()->setCellValue(self::columnName($this->_endCol) . ($index + $this->_beginRow + 1), strip_tags($value), true);
            $this->raiseEvent('onRenderDataCell', $cell, $value);
        }
    }


    /**
     * Generates the output footer row after a specific row number
     * @param int $row the row number after which the footer is to be generated
     */
    public function generateFooter($row) {
        $this->_endCol = 0;
        foreach ($this->columns as $n => $column) {
            $this->_endCol = $this->_endCol + 1;
            if ($column->footer) {
                $footer = trim($column->footer) !== '' ? $column->footer : $column->grid->blankDisplay;
                $cell = $this->_objPHPExcel->getActiveSheet()->setCellValue(self::columnName($this->_endCol) . ($row + 2), $footer, true);
                $this->raiseEvent('onRenderFooterCell', $cell, $footer);
            }
        }
    }
    
    /**
     * Returns an excel column name.
     * 
     * @param int $index the column index number
     * @return string
     */
    public static function columnName($index) {
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
     * Sets the HTTP headers needed by file download action.
     */
    protected function setHttpHeaders()
    {
        if (empty($this->exportConfig[$this->exportType])) {
            return;
        }
        extract($this->exportConfig[$this->exportType]);
        Yii::$app->getResponse()->getHeaders()
            ->set('Pragma', 'public')
            ->set('Expires', '0')
            ->set('Cache-Control', 'must-revalidate, post-check=0, pre-check=0')
            ->set('Content-Disposition', 'attachment; filename="' . $this->filename . '.' . $extension . '"')
            ->set('Content-type', $mime . '; charset=utf-8');
    }
    
    /**
     * Raises a callable event
     * @param string $event the event name
     * @param PHPExcelCell $cell the output cell 
     * @param string $output the output content for the cell
     */
    protected function raiseEvent($event, $cell, $output) {
        if (is_callable($this->$event)) {
            call_user_func($this->$event, $cell, $output, $this);
        }    
    }
    
    /**
     * Gets the default export configuration
     */
    protected function setDefaultExportConfig()
    {
        $this->_defaultExportConfig = [
            self::FORMAT_HTML => [
                'label' => Yii::t('kvexport', 'HTML'),
                'icon' => 'floppy-saved',
                'linkOptions'=>[],
                'options' => ['title' => Yii::t('kvexport', 'Save as HTML')],
                'confirmMsg' => Yii::t('kvexport', 'Export data as a HTML file. Ok to proceed?'),
                'mime' => 'text/html',
                'extension' => 'html',
                'writer' => 'HTML',
                'header' => true,
                'cssFile' => '',
            ],
            self::FORMAT_CSV => [
                'label' => Yii::t('kvexport', 'CSV'),
                'icon' => 'floppy-open',
                'colDelimiter' => ",",
                'rowDelimiter' => "\r\n",
                'linkOptions'=>[],
                'options' => ['title' => Yii::t('kvexport', 'Save as CSV')],
                'confirmMsg' => Yii::t('kvexport', 'Export data as Comma Separated Values (CSV). Ok to proceed?'),
                'mime' => 'application/csv',
                'extension' => 'csv',
                'writer' => 'CSV',
                'header' => false,
            ],
            self::FORMAT_PDF => [
                'label' => Yii::t('kvexport', 'PDF'),
                'icon' => 'floppy-save',
                'linkOptions'=>[],
                'options' => ['title' => Yii::t('kvexport', 'Save as PDF')],
                'confirmMsg' => Yii::t('kvexport', 'Export data as Portable Document Format (PDF). Ok to proceed?'),
                'mime' => 'application/pdf',
                'extension' => 'pdf',
                'writer' => 'PDF',
                'header' => true,
            ],
            self::FORMAT_EXCEL => [
                'label' => Yii::t('kvexport', 'Excel 95 +'),
                'icon' => 'floppy-remove',
                'linkOptions'=>[],
                'worksheet' => Yii::t('kvexport', 'ExportWorksheet'),
                'options' => ['title' => Yii::t('kvexport', 'Save as Excel (xls)')],
                'confirmMsg' => Yii::t('kvexport', 'Export data as Excel 95+ (xls) format. Ok to proceed?'),
                'mime' => 'application/vnd.ms-excel',
                'extension' => 'xls',
                'writer' => 'Excel5',
                'header' => true,
            ],
            self::FORMAT_EXCEL_X => [
                'label' => Yii::t('kvexport', 'Excel 2007+'),
                'icon' => 'floppy-remove',
                'linkOptions'=>[],
                'worksheet' => Yii::t('kvexport', 'ExportWorksheet'),
                'options' => ['title' => Yii::t('kvexport', 'Save as Excel (xlsx)')],
                'confirmMsg' => Yii::t('kvexport', 'Export data as Excel 2007+ (xlsx) format. Ok to proceed?'),
                'mime' => 'application/vnd.ms-excel',
                'extension' => 'xlsx',
                'writer' => 'Excel2007',
                'header' => true,
            ],
        ];
    }
 
    /**
     * Registers client assets needed for Export Data widget
     */
    protected function registerAssets()
    {
        $view = $this->getView();
        ExportDataAsset::register($view);
        
        foreach ($this->exportConfig as $format => $setting) {
            $id = $this->options['id'] . '-' . strtolower($format);
            $options = Json::encode(['formId'=>$this->options['id'] . '-form', 'confirmMsg' => $setting['confirmMsg']]);
            $view->registerJs("jQuery('#{$id}').exportdata({$options});");
        }
    }
}