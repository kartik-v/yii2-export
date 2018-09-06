/*!
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2018
 * @version   1.3.0
 *
 * Export Columns Selector Validation Module.
 *
 * Author: Kartik Visweswaran
 * Copyright: 2015, Kartik Visweswaran, Krajee.com
 * For more JQuery plugins visit http://plugins.krajee.com
 * For more Yii related demos visit http://demos.krajee.com
 */
(function ($) {
    "use strict";

    var ExportColumns = function (element, options) {
        var self = this;
        self.$element = $(element);
        self.options = options;
        self.listen();
    };

    ExportColumns.prototype = {
        constructor: ExportColumns,
        listen: function () {
            var self = this, $el = self.$element, $tog = $el.find('input[name="export_columns_toggle"]');
            $el.off('click').on('click', function (e) {
                e.stopPropagation();
            });
            $tog.off('change').on('change', function () {
                var checked = $tog.is(':checked');
                $el.find('input[name="export_columns_selector[]"]:not([disabled])').prop('checked', checked);
            });
        }
    };

    //ExportColumns plugin definition
    $.fn.exportcolumns = function (option) {
        var args = Array.apply(null, arguments);
        args.shift();
        return this.each(function () {
            var $this = $(this),
                data = $this.data('exportcolumns'),
                options = typeof option === 'object' && option;

            if (!data) {
                $this.data('exportcolumns', (data = new ExportColumns(this,
                    $.extend({}, $.fn.exportcolumns.defaults, options, $(this).data()))));
            }

            if (typeof option === 'string') {
                data[option].apply(data, args);
            }
        });
    };

    $.fn.exportcolumns.defaults = {};

})(window.jQuery);