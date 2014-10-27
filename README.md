yii2-export
===========

A library to export server/db data in various formats (e.g. excel, html, pdf, csv etc.). The widget allows you to configure the `dataProvider`, `columns` just like one would do within a `yii\grid\GridView`. However, the widget is made lean to not render or process default grid table markup. Instead it will display a user friendly export actions ButtonDropdown menu or list. This can be embedded into any of your views that may or may not contain a GridView. It uses PHPExcel library for converting data to the right format for export.

### Demo
You can see detailed [documentation](http://demos.krajee.com/export) on usage of the extension.

## Latest Release
>NOTE: The latest version of the module is v1.2.0 released on 25-Oct-2014. Refer the [CHANGE LOG](https://github.com/kartik-v/yii2-export/blob/master/CHANGE.md) for details.

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

### Slider

```php
use kartik\export\ExportGrid;
```

## License

**yii2-export** is released under the BSD 3-Clause License. See the bundled `LICENSE.md` for details.