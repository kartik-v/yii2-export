<h1 align="center">
    <a href="http://demos.krajee.com" title="Krajee Demos" target="_blank">
        <img src="http://kartik-v.github.io/bootstrap-fileinput-samples/samples/krajee-logo-b.png" alt="Krajee Logo"/>
    </a>
    <br>
    yii2-export
    <hr>
    <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=DTP3NZQ6G2AYU"
       title="Donate via Paypal" target="_blank"><img height="60" src="https://kartik-v.github.io/bootstrap-fileinput-samples/samples/donate.png" alt="Donate"/></a>
    &nbsp; &nbsp; &nbsp;
    <a href="https://www.buymeacoffee.com/kartikv" title="Buy me a coffee" ><img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" height="60" alt="kartikv" /></a>
</h1>

<div align="center">

[![Stable Version](https://poser.pugx.org/kartik-v/yii2-export/v/stable)](https://packagist.org/packages/kartik-v/yii2-export)
[![Unstable Version](https://poser.pugx.org/kartik-v/yii2-export/v/unstable)](https://packagist.org/packages/kartik-v/yii2-export)
[![License](https://poser.pugx.org/kartik-v/yii2-export/license)](https://packagist.org/packages/kartik-v/yii2-export)
[![Total Downloads](https://poser.pugx.org/kartik-v/yii2-export/downloads)](https://packagist.org/packages/kartik-v/yii2-export)
[![Monthly Downloads](https://poser.pugx.org/kartik-v/yii2-export/d/monthly)](https://packagist.org/packages/kartik-v/yii2-export)
[![Daily Downloads](https://poser.pugx.org/kartik-v/yii2-export/d/daily)](https://packagist.org/packages/kartik-v/yii2-export)

</div>

A library to export server/db data in various formats (e.g. excel, html, pdf, csv etc.) using the [PhpSpreadsheet](https://github.com/PHPOffice/phpspreadsheet) library. The widget allows you to configure the dataProvider, columns just like a yii\grid\GridView. However, it just displays the export actions in form of a ButtonDropdown menu, for embedding into any of your GridView or other components.

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

- `exportColumnsView` allows you to setup your own custom view file for rendering the column selector dropdown.
- `afterSaveView` allows you to setup your own after save view file if you are configuring to save exported file on server.

## Demo
You can see detailed [documentation](http://demos.krajee.com/export) and [demonstration](http://demos.krajee.com/export-demo) on usage of the extension.

## Release Changes
> NOTE: Refer the [CHANGE LOG](https://github.com/kartik-v/yii2-export/blob/master/CHANGE.md) for details on changes to various releases.

## Installation

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

> Note: Read this [web tip /wiki](http://webtips.krajee.com/setting-composer-minimum-stability-application/) on setting the `minimum-stability` settings for your application's composer.json.

### Pre-requisites

Install the necessary pre-requisite (Krajee Dropdown Extension) based on your bootstrap version:

- For Bootstrap v5.x install the extension `kartik-v/yii2-bootstrap5-dropdown`
- For Bootstrap v4.x install the extension `kartik-v/yii2-bootstrap4-dropdown`
- For Bootstrap v3.x install the extension `kartik-v/yii2-dropdown-x`

For example if you are using the Bootstrap v5.x add the following to the `require` section of your `composer.json` file:

```
"kartik-v/yii2-bootstrap5-dropdown": "@dev"
```

### Install

Either run:

```
$ php composer.phar require kartik-v/yii2-export "@dev"
```

or add

```
"kartik-v/yii2-export": "@dev"
```

to the `require` section of your `composer.json` file.

> Note: you must run `composer update` to have the latest stable dependencies like `kartik-v/yii2-krajee-base`

## Pre-requisites

The `yii2-export` extension is dependent on `yii2-grid` extension module. In order to start using `yii2-export`, you need to ensure setup of the `gridview` module in your application modules configuration file. For example:

```php
'modules' => [
    'gridview' => [
        'class' => 'kartik\grid\Module',
        // other module settings
    ]
]
```

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
    'columns' => $gridColumns,
    'clearBuffers' => true, //optional
]);

// You can choose to render your own GridView separately
echo \kartik\grid\GridView::widget([
    'dataProvider' => $dataProvider,
    'filterModel' => $searchModel,
    'columns' => $gridColumns
]);
```

## Contributors

### Code Contributors

This project exists thanks to all the people who contribute. [[Contribute](CONTRIBUTING.md)].
<a href="https://github.com/kartik-v/yii2-export/graphs/contributors"><img src="https://opencollective.com/yii2-export/contributors.svg?width=890&button=false" /></a>

### Financial Contributors

Become a financial contributor and help us sustain our community. [[Contribute](https://opencollective.com/yii2-export/contribute)]

#### Individuals

<a href="https://opencollective.com/yii2-export"><img src="https://opencollective.com/yii2-export/individuals.svg?width=890"></a>

#### Organizations

Support this project with your organization. Your logo will show up here with a link to your website. [[Contribute](https://opencollective.com/yii2-export/contribute)]

<a href="https://opencollective.com/yii2-export/organization/0/website"><img src="https://opencollective.com/yii2-export/organization/0/avatar.svg"></a>
<a href="https://opencollective.com/yii2-export/organization/1/website"><img src="https://opencollective.com/yii2-export/organization/1/avatar.svg"></a>
<a href="https://opencollective.com/yii2-export/organization/2/website"><img src="https://opencollective.com/yii2-export/organization/2/avatar.svg"></a>
<a href="https://opencollective.com/yii2-export/organization/3/website"><img src="https://opencollective.com/yii2-export/organization/3/avatar.svg"></a>
<a href="https://opencollective.com/yii2-export/organization/4/website"><img src="https://opencollective.com/yii2-export/organization/4/avatar.svg"></a>
<a href="https://opencollective.com/yii2-export/organization/5/website"><img src="https://opencollective.com/yii2-export/organization/5/avatar.svg"></a>
<a href="https://opencollective.com/yii2-export/organization/6/website"><img src="https://opencollective.com/yii2-export/organization/6/avatar.svg"></a>
<a href="https://opencollective.com/yii2-export/organization/7/website"><img src="https://opencollective.com/yii2-export/organization/7/avatar.svg"></a>
<a href="https://opencollective.com/yii2-export/organization/8/website"><img src="https://opencollective.com/yii2-export/organization/8/avatar.svg"></a>
<a href="https://opencollective.com/yii2-export/organization/9/website"><img src="https://opencollective.com/yii2-export/organization/9/avatar.svg"></a>

## License

**yii2-export** is released under the BSD-3-Clause License. See the bundled `LICENSE.md` for details.
