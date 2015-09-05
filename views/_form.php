<?php
/**
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015
 * @version   1.2.4
 *
 * Export Submission Form
 *
 * Author: Kartik Visweswaran
 * Copyright: 2015, Kartik Visweswaran, Krajee.com
 * For more JQuery plugins visit http://plugins.krajee.com
 * For more Yii related demos visit http://demos.krajee.com
 */
use yii\helpers\Html;

/**
 * @var array  $options
 * @var string $exportTypeParam
 * @var string $exportType
 * @var string $exportRequestParam
 * @var string $exportColsParam
 * @var string $colselFlagParam
 * @var string $columnSelectorEnabled
 */
echo Html::beginForm('', 'post', $options);
echo Html::hiddenInput($exportTypeParam, $exportType);
echo Html::hiddenInput($exportRequestParam, 1);
echo Html::hiddenInput($exportColsParam, '');
echo Html::hiddenInput($colselFlagParam, $columnSelectorEnabled);
echo Html::endForm();
?>