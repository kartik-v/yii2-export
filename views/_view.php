<?php
/**
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2017
 * @version   1.2.7
 *
 * Export Submission View
 *
 */

use yii\helpers\Html;

/**
 * @var string $icon
 * @var string $file
 * @var string $href
 */
?>
<div class="alert alert-success alert-dismissible" role="alert" style="margin:10px 0">
    <button type="button" class="close" data-dismiss="alert" aria-label="Close">
        <span aria-hidden="true">&times;</span>
    </button>
    <strong><?= Yii::t('kvexport', 'Exported File') ?>: </strong>
    <span class="h4" data-toggle="tooltip" title="<?= Yii::t('kvexport', 'Download exported file') ?>">
        <?= Html::a("<i class='{$icon}'></i> {$file}", $href, ['class' => 'label label-success']) ?>
    </span>
</div>