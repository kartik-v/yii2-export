<?php
/**
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2023
 * @version   1.4.3
 *
 * Export Submission View
 *
 */

use yii\helpers\Html;

/**
 * @var bool $notBs3
 * @var string $icon
 * @var string $file
 * @var string $href
 */
$badgePrefix = $notBs3 ? 'badge bg-' : 'label label-';
?>
<div class="alert alert-success alert-dismissible" role="alert" style="margin:10px 0">
    <button type="button" class="close" data-dismiss="alert" aria-label="Close">
        <span aria-hidden="true">&times;</span>
    </button>
    <strong><?= Yii::t('kvexport', 'Exported File') ?>: </strong>
    <span class="h5" data-toggle="tooltip" title="<?= Yii::t('kvexport', 'Download exported file') ?>">
        <?= Html::a("<i class='{$icon}'></i> {$file}", $href, ['class' => $badgePrefix.'success']) ?>
    </span>
</div>