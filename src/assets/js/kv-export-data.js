/*!
 * @package   yii2-export
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2018
 * @version   1.3.3
 *
 * Export Data Validation Module.
 *
 * Author: Kartik Visweswaran
 * Copyright: 2015, Kartik Visweswaran, Krajee.com
 * For more JQuery plugins visit http://plugins.krajee.com
 * For more Yii related demos visit http://demos.krajee.com
 */
(function ($) {
    "use strict";

    var $h = {
        isEmpty: function (value, trim) {
            return value === null || value === undefined || value === [] || value === '' || trim && $.trim(value) === '';
        },
        popupDialog: function (url, name, w, h) {
            var left = (screen.width / 2) - (w / 2), top = 60,
                existWin = window.open('', name, '', true);
            existWin.close();
            return window.open(url, name,
                'toolbar=no, location=no, directories=no, status=yes, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=' + w + ', height=' + h + ', top=' + top + ', left=' + left);
        },
        popupTemplate: '<html style="display:table;width:100%;height:100%;">' +
        '<title>Export Data - &copy; Krajee</title>' +
        '<body style="display:table-cell;font-family:Helvetica,Arial,sans-serif;color:#888;font-weight:bold;line-height:1.4em;text-align:center;vertical-align:middle;width:100%;height:100%;padding:0 10px;">' +
        '{msg}' +
        '</body>' +
        '</html>'
    };
    var ExportData = function (element, options) {
        var self = this;
        self.$element = $(element);
        $.each(options, function (key, val) {
            self[key] = val;
        });
        self.popup = '';
        self.listen();
    };

    ExportData.prototype = {
        constructor: ExportData,
        setPopupAlert: function (msg) {
            var self = this;
            if (!self.popup.document) {
                return;
            }
            if (arguments.length && arguments[1]) {
                var el = self.popup.document.getElementsByTagName('body');
                setTimeout(function () {
                    el[0].innerHTML = msg;
                }, 1800);
            } else {
                var newmsg = $h.popupTemplate.replace('{msg}', msg);
                self.popup.document.write(newmsg);
            }
        },
        listen: function () {
            var self = this;
            self.listenClick();
        },
        processExport: function (fmt) {
            var self = this, $selected, cols = [], $csrf, yiiLib = window.yii, isPopup, cfg = self.settings,
                frmConfig, expCols, $form, getInput = function (name, value) {
                    return $('<input/>', {'type': 'hidden', 'name': name, 'value': value});
                };
            isPopup = cfg.target === '_popup';
            if (isPopup) {
                self.popup = $h.popupDialog('', 'kvExportFullDialog', 350, 120);
                self.popup.focus();
                self.setPopupAlert(self.settings.messages.downloadProgress);
            }
            $csrf = yiiLib ? getInput(yiiLib.getCsrfParam() || '_csrf', yiiLib.getCsrfToken() || null) : null;
            expCols = '';
            if (!$h.isEmpty(cfg.colSelId)) {
                $selected = $('#' + cfg.colSelId).parent().find('input[name="export_columns_selector[]"]');
                $selected.each(function () {
                    var $el = $(this);
                    if ($el.is(':checked')) {
                        cols.push($el.attr('data-key'));
                    }
                });
                expCols = JSON.stringify(cols);
            }
            frmConfig = $.extend(true, {}, cfg.formOptions, {
                action: window.location.href,
                method: 'post',
                css: {display: 'none'}
            });
            $form = $('<form/>', frmConfig).append($csrf)
                .append(getInput(cfg.exportTypeParam, fmt), getInput(cfg.exportRequestParam, 1))
                .append(getInput(cfg.exportColsParam, expCols), getInput(cfg.colSelFlagParam, cfg.colSelEnabled))
                .appendTo('body');
            if (!$h.isEmpty(cfg.exportFormHiddenInputs)) {
                $.each(cfg.exportFormHiddenInputs, function (key, setting) {
                    var opts = {'name': key, 'value': setting.value || null, 'type': 'hidden'};
                    opts = $.extend({}, opts, setting.options || {});
                    $form.append($('<input/>', opts));
                });
            }
            $form.submit().remove();
            if (isPopup) {
                self.setPopupAlert(self.settings.messages.downloadComplete, true);
            }
        },
        listenClick: function () {
            var self = this;
            self.$element.off('click.exportmenu').on('click.exportmenu', function (e) {
                var fmt, msgs, msg = '', msg1, msg2, msg3, lib = window[self.settings.dialogLib];
                fmt = $(this).data('format');
                e.preventDefault();
                e.stopPropagation();
                if (!self.settings.showConfirmAlert) {
                    self.processExport(fmt);
                    return;
                }
                msgs = self.settings.messages;
                msg1 = $h.isEmpty(self.alertMsg) ? '' : self.alertMsg;
                msg2 = $h.isEmpty(msgs.allowPopups) ? '' : msgs.allowPopups;
                msg3 = $h.isEmpty(msgs.confirmDownload) ? '' : msgs.confirmDownload;
                if (msg1.length && msg2.length) {
                    msg = msg1 + '\n\n' + msg2;
                } else {
                    if (!msg1.length && msg2.length) {
                        msg = msg2;
                    } else {
                        msg = (msg1.length && !msg2.length) ? msg1 : '';
                    }
                }
                if (msg3.length) {
                    msg = msg + '\n\n' + msg3;
                }
                if ($h.isEmpty(msg)) {
                    self.processExport(fmt);
                    return;
                }
                lib.confirm(msg, function (result) {
                    if (result) {
                        self.processExport(fmt);
                    }
                });
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
                $this.data('exportdata',
                    (data = new ExportData(this, $.extend({}, $.fn.exportdata.defaults, options, $(this).data()))));
            }

            if (typeof option === 'string') {
                data[option].apply(data, args);
            }
        });
    };

    $.fn.exportdata.defaults = {
        alertMsg: '',
        settings: {
            formOptions: {},
            exportType: '',
            colSelEnabled: '',
            exportRequestParam: '',
            exportTypeParam: '',
            exportColsParam: '',
            exportFormHiddenInputs: {},
            colSelId: null,
            target: '_popup',
            showConfirmAlert: true,
            dialogLib: 'krajeeDialog',
            messages: {
                allowPopups: '',
                confirmDownload: '',
                downloadProgress: '',
                downloadComplete: ''
            }
        }
    };
})(window.jQuery);