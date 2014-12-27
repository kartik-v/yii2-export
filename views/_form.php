<?php
/*!
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2014
 * @version 1.2.0
 *
 * Export Submission Form
 *
 * Author: Kartik Visweswaran
 * Copyright: 2014, Kartik Visweswaran, Krajee.com
 * For more JQuery plugins visit http://plugins.krajee.com
 * For more Yii related demos visit http://demos.krajee.com
 */
use yii\helpers\Html;  
echo Html::beginForm('', 'post', $options);
echo Html::hiddenInput('export_type', $exportType);
echo Html::hiddenInput($exportRequestParam, 1); 
echo Html::hiddenInput('export_columns', '');
echo Html::endForm();
?>