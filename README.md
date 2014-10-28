yii2-export
===========

A library to export server/db data in various formats (e.g. excel, html, pdf, csv etc.). The widget allows you to configure the dataProvider, columns just like a yii\grid\GridView. However, it just displays the export actions in form of a ButtonDropdown menu, for embedding into any of your GridView or other components.

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
use kartik\export\ExportData;
```

## License

**yii2-export** is released under the BSD 3-Clause License. See the bundled `LICENSE.md` for details.