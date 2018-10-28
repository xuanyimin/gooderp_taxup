odoo.define('cn_invoice_verification.cn_account_invoice', function (require) {
"use strict";
var ListView = require('web.ListView');
var FormView = require('web.FormView');
var form_relational = require('web.form_relational');
var data = require('web.data');
var Model = require('web.Model');
var menu = require('web.Menu');
var Qweb = core.qweb;

ListView.include({
    render_buttons: function($node) {
                var self = this;
                this._super($node);
                    this.$buttons.find('.o_list_tender_button_create').click(this.proxy('tree_view_action'));
        },

        tree_view_action: function () {

        new Model("cn.account.invoice").call("_get_name",[self.model, $this.val(), self.datarecord.id]).then(
                            function() {
                                self.reload();
                                self.$el.find('input').val('');
                            }
                        );
        }

});
})