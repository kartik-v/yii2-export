<div class="alert alert-success alert-dismissible" role="alert" style="margin:10px 0">
    <button type="button" class="close" data-dismiss="alert" aria-label="Close">
        <span aria-hidden="true">&times;</span>
    </button>
    <strong><?= Yii::t('kvexport', 'Exported File')?>: </strong>
    <span class="h4" data-toggle="tooltip" title="<?= Yii::t('kvexport', 'Download exported file')?>">
        <?= \yii\helpers\Html::a("<i class='{$icon}'></i> {$file}", $href, ['class'=>'label label-success']) ?>
    </span>
</div>