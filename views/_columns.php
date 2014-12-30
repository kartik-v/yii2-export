<?php
use yii\helpers\Html;
use yii\helpers\ArrayHelper;

$label = ArrayHelper::remove($options, 'label');
$icon = ArrayHelper::remove($options, 'icon');
if (!empty($icon)) {
    $label = $icon . ' ' . $label;
}
echo Html::beginTag('div', ['class'=>'btn-group', 'role'=>'group']);
echo Html::button($label . ' <span class="caret"></span>', $options);
foreach($columnSelector as $value => $label) {
    if (in_array($value, $hiddenColumns)) {
        $checked = in_array($value, $selectedColumns);
        echo Html::checkbox('export_columns_selector[]', $checked, ['data-key'=>$value, 'style'=>'display:none']) . "\n";
        unset($columnSelector[$value]);
    }
    if (in_array($value, $noExportColumns)) {
        unset($columnSelector[$value]);
    }
}
echo Html::beginTag('ul', $menuOptions);
?>
<li>
    <div class="checkbox"> 
        <label>
        <?= Html::checkbox('export_columns_toggle', true) ?>
        <span class="kv-toggle-all"><?= Yii::t('kvexport', 'Toggle All') ?></span>
        </label>
    </div>
</li>
<li class="divider"></li>
<?php
foreach($columnSelector as $value => $label) {
    $checked = in_array($value, $selectedColumns);
    $disabled = in_array($value, $disabledColumns);
    $labelTag = $disabled ? '<label class="disabled">' : '<label>';
    echo '<li><div class="checkbox">' . $labelTag .
        Html::checkbox('export_columns_selector[]', $checked, ['data-key'=>$value, 'disabled'=>$disabled]) . 
        "\n" . $label . '</label></div></li>';
}
echo Html::endTag('ul');
echo Html::endTag('div');
?>
