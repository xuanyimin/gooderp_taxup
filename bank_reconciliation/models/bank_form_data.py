# -*- coding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2016  开阖软件(<http://www.osbzr.com>).
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################

from odoo import api, fields, models, tools, _
from odoo.exceptions import UserError
from odoo.tools.config import config
import pymssql
import xlwt
import xlrd
import base64
import datetime
from decimal import *

class create_bank_form_wizard(models.TransientModel):
    _name = 'create.bank.form.wizard'
    _description = 'Bank Form Import'

    excel = fields.Binary(u'导入系统导出的excel文件',)
    type = fields.Selection([('hzbank', u'湖州银行'),
                             ('acbcbank', u'工商银行'),
                             ('nsbank', u'农商银行'),
                             ('xybank', u'兴业银行'),
                             ('bankcomm', u'交通银行'),], u'回单银行', default='hzbank')

    @api.multi
    def create_bank_form(self):
        """
        通过Excel文件导入信息到hospital.invoice
        """
        order_id = self.env['bank.form'].browse(self.env.context.get('active_id'))
        print order_id
        if not order_id:
            return {}
        xls_data = xlrd.open_workbook(
                file_contents=base64.decodestring(self.excel))
        table = xls_data.sheets()[0]
        #取得行数
        ncows = table.nrows
        #取得第1行数据
        if self.type in ['hzbank','acbcbank','nsbank','bankcomm']:
            colnames =  table.row_values(1)
            list =[]
            newcows = 0
            for rownum in range(2,ncows):
                row = table.row_values(rownum)
                if row:
                    app = {}
                    for i in range(len(colnames)):
                       app[colnames[i]] = row[i]
                    #过滤掉不需要的行，详见销货清单的会在清单中再次导入
                    if app.get(u'交易日期') or app.get(u'交易时间'):
                        list.append(app)
                        newcows += 1
        if self.type in ['xybank']:
            colnames = table.row_values(0)
            list = []
            newcows = 0
            for rownum in range(1, ncows):
                row = table.row_values(rownum)
                if row:
                    app = {}
                    for i in range(len(colnames)):
                        app[colnames[i]] = row[i]
                    # 过滤掉不需要的行，详见销货清单的会在清单中再次导入
                    if app.get(u'交易日期') or app.get(u'交易时间'):
                        list.append(app)
                        newcows += 1
        #数据读入。
        for data in range(0,newcows):
            in_xls_data = list[data]
            if self.type == 'hzbank':
                amount_out = in_xls_data.get(u'支出') and in_xls_data.get(u'支出').replace(',', '').strip() or 0.00
                amount_in = in_xls_data.get(u'收入') and in_xls_data.get(u'收入').replace(',', '').strip() or 0.00
                self.env['bank.form.line'].create({
                    'name': in_xls_data.get(u'对方户名') or '',
                    'num': in_xls_data.get(u'对方账/卡号') or '',
                    'amount_in': float(amount_in),
                    'amount_out': float(amount_out),
                    'date': in_xls_data.get(u'交易日期'),
                    'note':in_xls_data.get(u'摘要') or '',
                    'order_id':order_id.id,})
            if self.type == 'xybank':
                amount_out = in_xls_data.get(u'贷方金额') and in_xls_data.get(u'贷方金额').replace(',', '').strip() or 0.00
                amount_in = in_xls_data.get(u'借方金额') and in_xls_data.get(u'借方金额').replace(',', '').strip() or 0.00
                self.env['bank.form.line'].create({
                    'name': in_xls_data.get(u'对方户名') or '',
                    'num': in_xls_data.get(u'对方账号') or '',
                    'amount_in': float(amount_in),
                    'amount_out': float(amount_out),
                    'date': in_xls_data.get(u'交易日期'),
                    'note': in_xls_data.get(u'摘要') or '',
                    'purpose': in_xls_data.get(u'用途') or '',
                    'order_id': order_id.id, })
            if self.type == 'acbcbank':
                amount_out = in_xls_data.get(u'借方发生额') and in_xls_data.get(u'借方发生额').replace(',', '').strip() or 0.00
                amount_in = in_xls_data.get(u'贷方发生额') and in_xls_data.get(u'贷方发生额').replace(',', '').strip() or 0.00
                self.env['bank.form.line'].create({
                    'name': in_xls_data.get(u'对方单位名称') or '',
                    'num': in_xls_data.get(u'对方账号') or '',
                    'amount_in': float(amount_in),
                    'amount_out': float(amount_out),
                    'date': in_xls_data.get(u'交易时间'),
                    'note': in_xls_data.get(u'摘要') or '',
                    'purpose': in_xls_data.get(u'用途') or '',
                    'order_id': order_id.id, })
            if self.type == 'nsbank':
                amount_out = in_xls_data.get(u'汇出金额') and in_xls_data.get(u'汇出金额').replace(',', '').strip() or 0.00
                amount_in = in_xls_data.get(u'汇入金额') and in_xls_data.get(u'汇入金额').replace(',', '').strip() or 0.00
                self.env['bank.form.line'].create({
                    'name': in_xls_data.get(u'对方户名') or '',
                    'num': in_xls_data.get(u'对方账号') or '',
                    'amount_in': float(amount_in),
                    'amount_out': float(amount_out),
                    'date': in_xls_data.get(u'交易时间'),
                    'note': in_xls_data.get(u'摘要') or '',
                    'purpose': in_xls_data.get(u'备注') or '',
                    'order_id': order_id.id, })
            if self.type == 'bankcomm':
                amount_out = in_xls_data.get(u'借贷标志') and in_xls_data.get(u'借贷标志')==u'借' and in_xls_data.get(u'发生额').replace(',', '').strip() or 0.00
                amount_in = in_xls_data.get(u'借贷标志') and in_xls_data.get(u'借贷标志')==u'贷' and in_xls_data.get(u'发生额').replace(',', '').strip() or 0.00
                self.env['bank.form.line'].create({
                    'name': in_xls_data.get(u'对方户名') or '',
                    'num': in_xls_data.get(u'对方账号') or '',
                    'amount_in': float(amount_in),
                    'amount_out': float(amount_out),
                    'date': in_xls_data.get(u'交易时间'),
                    'note': in_xls_data.get(u'摘要') or '',
                    'purpose': in_xls_data.get(u'备注') or '',
                    'order_id': order_id.id, })


    def excel_date(self,data):
        #将excel日期改为正常日期
        if type(data) in (int,float):
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(data,0)
            py_date = datetime.datetime(year, month, day, hour, minute, second)
        else:
            py_date = data
        return py_date


