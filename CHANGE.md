version 1.2.0
=============
**Date:** 12-Jan-2014

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
- Revamp to use new Krajee base Module and TranslationTrait.
- Code formatting updates as per Yii2 coding style.


version 1.1.0
=============
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

version 1.0.0
=============
**Date:** 17-Dec-2014

- Initial release