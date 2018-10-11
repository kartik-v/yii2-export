<?php
/**
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2018
 * @version   1.3.4
 * 
 * Column Selector View
 *
 */

use yii\helpers\Html;
use yii\helpers\ArrayHelper;

/**
 * @var bool $isBs4
 * @var array $options
 * @var array $batchToggle
 * @var array $columnSelector
 * @var array $hiddenColumns
 * @var array $selectedColumns
 * @var array $disabledColumns
 * @var array $noExportColumns
 * @var array $menuOptions
 */

$label = ArrayHelper::remove($options, 'label');
$icon = ArrayHelper::remove($options, 'icon');
$showToggle = ArrayHelper::remove($batchToggle, 'show', true);
if (!empty($icon)) {
    $label = $icon . ' ' . $label;
}
echo Html::beginTag('div', ['class' => 'btn-group', 'role' => 'group']);
echo Html::button($label . ' <span class="caret"></span>', $options);
foreach ($columnSelector as $value => $label) {
    if (in_array($value, $hiddenColumns)) {
        $checked = in_array($value, $selectedColumns);
        echo Html::checkbox('export_columns_selector[]', $checked, ['data-key' => $value, 'style' => 'display:none']);
        unset($columnSelector[$value]);
    }
    if (in_array($value, $noExportColumns)) {
        unset($columnSelector[$value]);
    }
}
echo Html::beginTag('ul', $menuOptions);
?>

<?php if ($showToggle): ?>
    <?php
    $toggleOptions = ArrayHelper::remove($batchToggle, 'options', []);
    $toggleLabel = ArrayHelper::remove($batchToggle, 'label', Yii::t('kvexport', 'Toggle All'));
    Html::addCssClass($toggleOptions, 'kv-toggle-all');
    ?>
    <li>
        <div class="checkbox">
            <label>
                <?= Html::checkbox('export_columns_toggle', true) ?>
                <?= Html::tag('span', $toggleLabel, $toggleOptions) ?>
            </label>
        </div>
    </li>
    <li class="<?= $isBs4 ? 'dropdown-' : '' ?>divider"></li>
<?php endif; ?>

<?php
foreach ($columnSelector as $value => $label) {
    $checked = in_array($value, $selectedColumns);
    $disabled = in_array($value, $disabledColumns);
    $labelTag = $disabled ? '<label class="disabled">' : '<label>';
    echo '<li><div class="checkbox">' . $labelTag .
        Html::checkbox('export_columns_selector[]', $checked, ['data-key' => $value, 'disabled' => $disabled]) .
        "\n" . $label . '</label></div></li>';
}
echo Html::endTag('ul');
echo Html::endTag('div');
?>
