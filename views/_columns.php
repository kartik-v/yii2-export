<?php
use yii\helpers\Html;
use yii\helpers\ArrayHelper;
$menuOptions = [
    'role' => 'menu',
    'class' => 'dropdown-menu kv-checkbox-list',
    'aria-labelledby' => $options['id']
];
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
