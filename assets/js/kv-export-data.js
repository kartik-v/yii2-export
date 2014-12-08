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
    var isEmpty = function (value, trim) {
        return value === null || value === undefined || value == []
        || value === '' || trim && $.trim(value) === '';
    }, 
    popupDialog = function (url, name, w, h) {
        var left = (screen.width / 2) - (w / 2);
        var top = 60; //(screen.height / 2) - (h / 2);
        return window.open(url, name, 'toolbar=no, location=no, directories=no, status=yes, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=' 
            + w +', height=' + h + ', top=' + top + ', left=' + left);
    },
    popupTemplate = '<html style="display:table;width:100%;height:100%;">' +
        '<title>Export Data - &copy; Krajee</title>' +
        '<body style="display:table-cell;font-family:Helvetica,Arial,sans-serif;color:#888;font-weight:bold;line-height:1.4em;text-align:center;vertical-align:middle;width:100%;height:100%;padding:0 10px;">' +
        '{msg}' +
        '</body>' +
        '</html>';
    
    var ExportData = function (element, options) {
        this.$element = $(element);
        var settings = options.settings;
        this.alertMsg = options.alertMsg;
        this.$form = $("#" + settings.formId);
        this.messages = settings.messages;
        this.popup = '';
        this.listen();
    };
    
    ExportData.prototype = {
        constructor: ExportData,
        popupDialog: function (url, name, w, h) {
            var left = (screen.width / 2) - (w / 2);
            var top = 60;
            return window.open(url, name, 'toolbar=no, location=no, directories=no, status=yes, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=' 
                + w +', height=' + h + ', top=' + top + ', left=' + left);
        },
        kvConfirm: function(msg) {
            if (isEmpty(msg)) {
                return true;
            }
            try {
               return confirm(msg);
            } 
            catch (err) {
                return true;
            }
        },
        notify: function (e) {
            var self = this, msgs = self.messages;
            var msg1 = isEmpty(self.alertMsg) ? '' : self.alertMsg,
                msg2 = isEmpty(msgs.allowPopups) ? '' : msgs.allowPopups,
                msg3 = isEmpty(msgs.confirmDownload) ? '' : msgs.confirmDownload,
                msg = '', out;
            if (msg1.length && msg2.length) {
                msg = msg1 + '\n\n' + msg2;
            } else {
                if (!msg1.length && msg2.length) {
                    msg = msg2;
                } else {
                    msg = (msg1.length && !msg2.length) ? msg1 : ''
                }
            }
            if (msg3.length) {
                msg = msg + '\n\n' + msg3;
            }
            e.preventDefault();
            if (isEmpty(msg)) {
                return true;
            }
            return self.kvConfirm(msg);
        },
        setPopupAlert: function (msg) {
            var self = this;
            if (self.popup.document == undefined) {
                return;
            }
            if (arguments.length && arguments[1]) {
                var el =  self.popup.document.getElementsByTagName('body');
                setTimeout(function () {
                    el[0].innerHTML = msg;
                }, 4000);
            } else {
                var newmsg = popupTemplate.replace('{msg}', msg);
                self.popup.document.write(newmsg);
            }
        },
        listen: function () {
            var self = this;
            self.$form.appendTo('body');
            self.$form.on('submit', function() {
                setTimeout(function () {
                    self.setPopupAlert(self.messages.downloadComplete, true);
                }, 1000);
            });
            self.$element.on('click', function(e) {
                if (self.notify(e)) {
                    var fmt = $(this).data('format');
                    self.$form.find('[name="export_type"]').val(fmt);
                    self.popup = popupDialog('', 'kvDownloadDialog', 350, 120);
                    self.popup.focus();
                    self.setPopupAlert(self.messages.downloadProgress);
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
        settings: {
            formId: '',
            messages: {
                allowPopups: '',
                confirmDownload: '',
                downloadProgress: '',
                downloadComplete: ''
            }
        }
    };

})(window.jQuery);
