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

class ExportUnit(models.Model):
    '''海关报关单商品单位信息'''
    _name = 'export.unit'
    _order = "name"

    code = fields.Char(u'出口商品单位编号', required=True,)
    name = fields.Char(u'出口商品单位名称', required=True,)

class ExportProduct(models.Model):
    '''海关报关单商品海关信息'''
    _name = 'export.product'
    _order = "name"

    code = fields.Char(u'编码', )
    name = fields.Char(u'名称', )
    begin_date = fields.Date(u'起始日期',)
    end_date = fields.Date(u'截止日期',)
    unit = fields.Many2one('export.unit', string=u'单位')
    is_real = fields.Boolean(u'基本商品标志', default=False)
    tax_category = fields.Char(u'税种', )
    taxation_rate = fields.Float(u'征税税率', )
    drawback_rate = fields.Float(u'退税税率', )
    count_rate = fields.Float(u'从量定额征税率', )
    valuation_rate = fields.Float(u'从价定额征税率', )

class ExportPort(models.Model):
    '''海关报关单海关信息'''
    _name = 'export.port'
    _order = "name"

    code = fields.Char(u'编码', )
    name = fields.Char(u'名称', )
