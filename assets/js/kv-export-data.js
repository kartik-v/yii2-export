/*!
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2014
 * @version 1.0.0
 *
 * Export Data Validation Module.
 *
 * Author: Kartik Visweswaran
 * Copyright: 2014, Kartik Visweswaran, Krajee.com
 * For more JQuery plugins visit http://plugins.krajee.com
 * For more Yii related demos visit http://demos.krajee.com
 */
(function ($) {

    var ExportData = function (element, options) {
        this.$element = $(element);
        this.confirmMsg = options.confirmMsg;
        this.$form = $("#" + options.formId);
        this.listen();
    };
    
    ExportData.prototype = {
        constructor: ExportData,
        listen: function () {
            var self = this;
            self.$form.appendTo('body');
            self.$element.on('click', function(e) {
                e.preventDefault();
                if (confirm(self.confirmMsg)) {
                    self.$form.trigger('submit');
                }
            });
        }
    };

    //ExportData plugin definition
    $.fn.exportdata = function (option) {
        var args = Array.apply(null, arguments);
        args.shift();
        return this.each(function () {
            var $this = $(this),
                data = $this.data('exportdata'),
                options = typeof option === 'object' && option;

            if (!data) {
                $this.data('exportdata', (data = new ExportData(this, $.extend({}, $.fn.exportdata.defaults, options, $(this).data()))));
            }

            if (typeof option === 'string') {
                data[option].apply(data, args);
            }
        });
    };

    $.fn.exportdata.defaults = {
        filename: 'export',
        alertMsg: '',
        confirmMsg: '',
        formId: ''
    };

})(window.jQuery);
