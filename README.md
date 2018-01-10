yii2-export
===========

[![Stable Version](https://poser.pugx.org/kartik-v/yii2-export/v/stable)](https://packagist.org/packages/kartik-v/yii2-export)
[![Unstable Version](https://poser.pugx.org/kartik-v/yii2-export/v/unstable)](https://packagist.org/packages/kartik-v/yii2-export)
[![License](https://poser.pugx.org/kartik-v/yii2-export/license)](https://packagist.org/packages/kartik-v/yii2-export)
[![Total Downloads](https://poser.pugx.org/kartik-v/yii2-export/downloads)](https://packagist.org/packages/kartik-v/yii2-export)
[![Monthly Downloads](https://poser.pugx.org/kartik-v/yii2-export/d/monthly)](https://packagist.org/packages/kartik-v/yii2-export)
[![Daily Downloads](https://poser.pugx.org/kartik-v/yii2-export/d/daily)](https://packagist.org/packages/kartik-v/yii2-export)

A library to export server/db data in various formats (e.g. excel, html, pdf, csv etc.) using the [\PhpOffice\PhpSpreadsheet\Spreadsheet library](https://phpexcel.codeplex.com/). The widget allows you to configure the dataProvider, columns just like a yii\grid\GridView. However, it just displays the export actions in form of a ButtonDropdown menu, for embedding into any of your GridView or other components.

In addition, with release v1.2.0, the extension also displays a handy grid columns selector for controlling the columns for export. The features available with the column selector are:

- shows a column picker dropdown list to allow selection of columns for export.
- new `container` property allows you to group the export menu and column selector dropdowns.
- new `template` property for manipulating the display of menu, column selector or additional buttons in button group.
- allows configuration of column picker dropdown button through `columnSelectorOptions`
- auto-generates column labels in the column selector. But you can override displayed column labels for each column key through `columnSelector` property settings.
- allows preselected columns through `selectedColumns` (you must set the selected column keys)
- allows columns to be disabled in column selector through `disabledColumns` (you must set the disabled column keys)
- allows columns to be hidden in column selector through `hiddenColumns` (you must set the hidden column keys)
- allows columns to be hidden from both export and column selector through `noExportColumns` (you must set the no export column keys)
- toggle display of the column selector through `showColumnSelector` property
- column selector is displayed only if `asDropdown` is set to `true`.

The extension offers configurable user interfaces for advanced cases using view templates.

- `exportFormView` allows you to setup your own custom view file for rendering the export form.
- `exportColumnsView` allows you to setup your own custom view file for rendering the column selector dropdown.

## Demo
You can see detailed [documentation](http://demos.krajee.com/export) and [demonstration](http://demos.krajee.com/export-demo) on usage of the extension.

## Latest Release
>NOTE: The latest version of the extension is v1.2.8. Refer the [CHANGE LOG](https://github.com/kartik-v/yii2-export/blob/master/CHANGE.md) for details.

## Installation

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

> Note: Read this [web tip /wiki](http://webtips.krajee.com/setting-composer-minimum-stability-application/) on setting the `minimum-stability` settings for your application's composer.json.

Either run

```
$ php composer.phar require kartik-v/yii2-export "@dev"
```

or add

```
"kartik-v/yii2-export": "@dev"
```

to the `require` section of your `composer.json` file.

## Usage

### ExportMenu

```php
use kartik\export\ExportMenu;
$gridColumns = [
    ['class' => 'yii\grid\SerialColumn'],
    'id',
    'name',
    'color',
    'publish_date',
    'status',
    ['class' => 'yii\grid\ActionColumn'],
];

// Renders a export dropdown menu
echo ExportMenu::widget([
    'dataProvider' => $dataProvider,
    'columns' => $gridColumns
]);

// You can choose to render your own GridView separately
echo \kartik\grid\GridView::widget([
    'dataProvider' => $dataProvider,
    'filterModel' => $searchModel,
    'columns' => $gridColumns
]);
```

## License

**yii2-export** is released under the BSD 3-Clause License. See the bundled `LICENSE.md` for details.