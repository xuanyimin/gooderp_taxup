# -*- coding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2016  德清武康开源软件
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it ied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#will be useful,
#    but WITHOUT ANY WARRANTY; without even the impl
##############################################################################

from odoo import api, fields, models, tools, _
from odoo.exceptions import UserError

#贸易方式
MY_TYPE = [
    ('0110', u'一般贸易')]
#成交方式
CJ_TYPE = [
    ('FOB', u'FOB'),
    ('CNF', u'C&F'),
    ('CIF', u'CIF')]
#币制
UNIT_SELECT= [
    ('fd_unit', u'法定单位'),
    ('no2_fd_unit', u'第二单位'),
    ('cj_unit', u'报关单位')
]
YB_BZ =[
    ('USD', u'美元'),
    ('GBP', u'英镑'),
    ('EUR', u'欧元'),
    ('CNY', u'人民币')]

class ExportDeclaration(models.Model):
    '''海关报关单信息'''
    _name = 'export.declaration'
    _order = "name"

    name = fields.Char(u'海关编号', copy=False, required=True,)
    lj_date = fields.Date(u'出口日期', copy=False, required=True,)
    cj_type = fields.Char(u'成交方式',copy=False,)
    ht_no = fields.Char(u'合同协议号', copy=False, )
    bf = fields.Float(u"运费",copy=False)
    zf = fields.Float(u"保费",copy=False)
    export_amount = fields.Float(u"合计金额",copy=False)
    ht_no = fields.Char(u'合同协议号',copy=False,)
    currency_id = fields.Many2one('res.currency', u'币别',)
    my_type = fields.Selection(MY_TYPE, u'贸易方式',copy=False,)
    line_ids = fields.One2many('export.declaration.line', 'order_id', u'报关单明细行',
                               copy=False)
    is_declare = fields.Boolean(u'已申报', default=False)

class ExportDeclarationLine(models.Model):
    '''海关报关单信息'''
    _name = 'export.declaration.line'

    order_id = fields.Many2one('export.declaration', u'订单编号', index=True,
                               required=True, ondelete='cascade',
                               help=u'关联订单的编号')
    spxh = fields.Char(u'商品序号',)
    cmcode = fields.Char(u'商品编码',)
    cm_name = fields.Char(u'商品名称',)
    yb_bz = fields.Many2one('res.currency', u'币别',)
    yb_amt = fields.Float(u'总价',)
    unit_ids = fields.One2many('export.declaration.line.unit', 'line_id', u'报关单明细行单位',
                               copy=False)

    @api.multi
    @api.depends('spxh', 'cmcode', 'cm_name')
    def name_get(self):
        """
        在其他model中用到account时在页面显示 code name balance如：2202 应付账款 当前余额（更有利于会计记账）
        :return:
        """
        result = []
        for line in self:
            account_name = line.spxh + ' ' +line.cmcode + ' ' + line.cm_name
            result.append((line.id, account_name))
        return result

class ExportDeclarationLineUnit(models.Model):
    '''海关报关单信息单位'''
    _name = 'export.declaration.line.unit'

    line_id = fields.Many2one('export.declaration.line', u'报关单明细行', index=True,
                               required=True, ondelete='cascade',
                               help=u'关联报关单')
    note = fields.Selection(UNIT_SELECT, u'类型',)
    qnt = fields.Float(u'数量', )
    unit = fields.Many2one('export.unit', u'单位',)
