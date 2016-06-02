Change Log: `yii2-export`
=========================

## version 1.2.6

**Date:** 02-Jun-2016

- (enh #135): Add Indonesian translations.
- (enh #137): Add Estonian translations.
- (enh #138): Add dynagrid selection support.
- (enh #141): Add grouping option in export column.

## version 1.2.5

**Date:** 18-Apr-2016

- (enh #115): Add Polish translations.
- (enh #118): Validation for empty value.
- (enh #117): Add German translations.
- (enh #119): Add Hungarian translations.
- (enh #121): Add Dutch translations.
- (enh #123): Allow the exported filename to have spaces.
- (enh #124): Add Italian translations.
- (enh #133): Modify default `pdfLibrary` setting for mPDF.

## version 1.2.4

**Date:** 04-Feb-2016

- (enh #78): Add Portugese Brazilian translations.
- (enh #87): Cache dataProvider total count (for performance).
- (enh #89): New property `onInitExcel` as an event for `initPHPExcel` method.
- (enh #98): More correct models count for generateBody.
- (enh #99): New setter methods `setPHPExcel`, `setPHPExcelWriter`, `setPHPExcelSheet`
- PHP comment formatting and PHPDoc updates.
- (enh #100): Add Czech language translations.
- (enh #100): Add Czech language translations.
- (enh #112): Added option to configure timeout.
- (enh #114): Add composer branch alias to allow getting latest `dev-master` updates.

## version 1.2.3

**Date:** 19-Jul-2015

- (enh #43): Added new `clearBuffers` property for better fix of #40.
- (enh #44): Improve validation to retrieve the right translation messages folder.
- (enh #45): Fix buffer clearing.
- (enh #46): New `pjaxContainerId` property added to widget to enable refreshing via pjax.
- (enh #47): Set asset bundle dependencies with yii2-krajee-base.
- (enh #49): Set a better PHP Excel version dependency.
- (enh #50): Better exit and resetting of memory after output generation.
- (enh #51): Fix to correct right filtering of exported data via pjax.
- (enh #52): Bind export elements better on jQuery events.
- (bug #62): Correct export request param for allowing multiple export menus on the same page.
- (bug #64): Alternative buffer emulation by setting `stream` to `false` and `streamAfterSave` to `true`.
- (enh #69): Various enhancements to export functionality.
- (enh #72): Configurable menu container tag when `asDropdown` is `false`. Fixes #73.
- (enh #75): Add Spanish translations.
- (enh #76): Allow fetching models in batches. Fixes #70.

## version 1.2.2

**Date:** 14-Feb-2015

- (enh #37): Added zh-CN translations
- (enh #39): Set AssetBundle dependency to kartik\base\AssetBundle.
- (enh #40): Fix buffer clearing.
- (enh #41): New bool property `initProvider` to clear previously fetched models before render.
- Set copyright year to current.

## version 1.2.1

**Date:** 20-Jan-2015

- (bug #34): Set lastModifiedBy to default to username instead of datetime 

## version 1.2.0

**Date:** 12-Jan-2015

- (enh #17): New column selector feature to allow selection of columns
    - shows a column picker dropdown list to allow selection of columns for export.
    - new `container` property allows you to group the export menu and column selector dropdowns.
    - allows configuration of column picker dropdown button through `columnSelectorOptions`
    - auto-generates column labels in the column selector. But you can override displayed column labels for each column key through `columnSelector` property settings.
    - allows preselected columns through `selectedColumns` (you must set the selected column keys)
    - allows columns to be disabled in column selector through `disabledColumns` (you must set the disabled column keys)
    - allows columns to be hidden in column selector through `hiddenColumns` (you must set the hidden column keys)
    - allows columns to be hidden from both export and column selector through `noExportColumns` (you must set the no export column keys)
    - toggle display of the column selector through `showColumnSelector` property
    - column selector is displayed only if `asDropdown` is set to `true`.
- (enh #18): Configurable user interfaces for advanced cases using view templates
    - `exportFormView` allows you to setup your own custom view file for rendering the export form.
    - `exportColumnsView` allows you to setup your own custom view file for rendering the column selector dropdown.
- added ability to configure export form HTML options.
- (enh #25): Template to configure the export menu and column selector button groups.
- (enh #26): Russian translation added.
- (enh #27): Add ability to toggle (check/uncheck) all columns in column selector.
- (bug #28): Validation of column name correctly when 0 or 1 column selected.
- (enh #29): Display emptyText when no columns selected.
- (enh #30): Create new jquery plugin for export columns selector.
- (enh #31): Configure separate `AssetBundle` for export columns selector.
- (enh #32): Add`columnBatchToggleSettings` to configure column toggle all checkbox.
- (enh #33): Updated Russian translations.
- Revamp to use new Krajee base Module and TranslationTrait.
- Code formatting updates as per Yii2 coding style.

## version 1.1.0

**Date:** 26-Dec-2014

- (enh #10): Set composer json dependency for yii2-grid.
- (enh #11): Option to set target for export form submission.
- (enh #12): Translations for Portugese (pt-PT).
- (enh #13): Enhance download popup dialog to properly reset.
- (enh #14): New property `showConfirmAlert` that controls display of the javascript confirmation dialog before download.
- (bug #16): Prevent default on menu item click when showConfirmAlert is set to false.
- (bug #19): Correct rendering of menu when format config is disabled.
- (bug #21): Set correct reference to `ActiveDataProvider` and `ActiveQueryInterface`.
- (enh #22): Add property `enableFormatter` to enable/disable yii grid formatter.

## version 1.0.0

**Date:** 17-Dec-2014

- Initial release