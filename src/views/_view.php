<?php
/**
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2018
 * @version   1.3.7
 *
 * Export Submission View
 *
 */

use yii\helpers\Html;

/**
 * @var bool $isBs4
 * @var string $icon
 * @var string $file
 * @var string $href
 */
$badgePrefix = $isBs4 ? 'badge badge-' : 'label label-';
?>
<div class="alert alert-success alert-dismissible" role="alert" style="margin:10px 0">
    <button type="button" class="close" data-dismiss="alert" aria-label="Close">
        <span aria-hidden="true">&times;</span>
    </button>
    <strong><?= Yii::t('kvexport', 'Exported File') ?>: </strong>
    <span class="h5" data-toggle="tooltip" title="<?= Yii::t('kvexport', 'Download exported file') ?>">
        <?= Html::a("<i class='{$icon}'></i> {$file}", $href, ['class' => $badgePrefix . 'success']) ?>
    </span>
</div>