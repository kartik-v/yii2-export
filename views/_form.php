<?php
/**
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2017
 * @version   1.2.8
 *
 * Export Submission Form
 *
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
