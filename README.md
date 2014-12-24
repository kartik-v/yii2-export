yii2-export
===========

A library to export server/db data in various formats (e.g. excel, html, pdf, csv etc.) using the [PHPExcel library](https://phpexcel.codeplex.com/). The widget allows you to configure the dataProvider, columns just like a yii\grid\GridView. However, it just displays the export actions in form of a ButtonDropdown menu, for embedding into any of your GridView or other components.

## Demo
You can see detailed [documentation](http://demos.krajee.com/export) and [demonstration](http://demos.krajee.com/export-demo) on usage of the extension.

## Latest Release
>NOTE: The latest version of the extension is v1.1.0 released on 25-Dec-2014. Refer the [CHANGE LOG](https://github.com/kartik-v/yii2-export/blob/master/CHANGE.md) for details.

## Installation

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

> Note: Read this [web tip /wiki](http://webtips.krajee.com/setting-composer-minimum-stability-application/) on setting the `minimum-stability` settings for your application's composer.json.

Either run

```
$ php composer.phar require kartik-v/yii2-export "dev-master"
```

or add

```
"kartik-v/yii2-export": "dev-master"
```

to the ```require``` section of your `composer.json` file.

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