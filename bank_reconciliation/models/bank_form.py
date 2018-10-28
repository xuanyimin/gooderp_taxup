# -*- coding: utf-8 -*-
##############################################################################
#
#    Copyright (C) 2016  德清武康开源软件().
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundaption, either version 3 of the
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
from odoo.tools.config import config
import pymssql
import xlrd
import xlwt
import re
import base64
from odoo.exceptions import UserError

# 字段只读状态
READONLY_STATES = {
        'done': [('readonly', True)],
    }

#增加引出K3销售相关内容
class BankForm(models.Model):
    _name = 'bank.form'
    _order = "name"

    name = fields.Many2one('bank.account',u'银行科目',required=True)
    period_id = fields.Many2one(
        'finance.period',
        u'会计期间',
        ondelete='restrict',
        required=True,
        states=READONLY_STATES)
    begin_id =  fields.Integer(u'起始凭证号', required=True, copy=False,)
    line_ids = fields.One2many('bank.form.line', 'order_id', u'对帐单明细',
                               states=READONLY_STATES, copy=False)
    state = fields.Selection([('draft', u'草稿'),
                              ('done', u'已结束')], u'状态', default='draft')
    attachment_number = fields.Integer(compute='_compute_attachment_number', string=u'附件号')

    @api.multi
    def action_get_attachment_view(self):
        res = self.env['ir.actions.act_window'].for_xml_id('base', 'action_attachment')
        res['domain'] = [('res_model', '=', 'bank.form'), ('res_id', 'in', self.ids)]
        res['context'] = {'default_res_model': 'bank.form', 'default_res_id': self.id}
        return res

    @api.multi
    def _compute_attachment_number(self):
        attachment_data = self.env['ir.attachment'].read_group(
            [('res_model', '=', 'bank.form'), ('res_id', 'in', self.ids)], ['res_id'], ['res_id'])
        attachment = dict((data['res_id'], data['res_id_count']) for data in attachment_data)
        for expense in self:
            expense.attachment_number = attachment.get(expense.id, 0)

    @api.multi
    def button_excel(self):
        return {
            'name': u'引入excel',
            'view_mode': 'form',
            'view_type': 'form',
            'res_model': 'create.bank.form.wizard',
            'type': 'ir.actions.act_window',
            'target': 'new',
        }

    @api.multi
    def exp_k3_bank_voucher(self):
        xls_data = xlrd.open_workbook('./excel/moban.xls')
        Page1 = xls_data.sheet_by_name('Page1')
        Page2 = xls_data.sheet_by_name('t_Schema')
        # 连接数据库
        conn = self.createConnection()
        excel, colnames = self.env['tax.invoice.out'].readexcel(Page1)  # 读模版，返回字典及表头数组
        workbook = xlwt.Workbook(encoding='utf-8')  # 生成文件
        worksheet = workbook.add_sheet(u'Page1')  # 在文件中创建一个名为Page1的sheet
        worksheet2 = workbook.add_sheet(u't_Schema')
        self.env['tax.invoice.out'].worksheetcopy(Page2, worksheet2)

        i = j = 0
        number = self.begin_id
        for key in colnames:
            worksheet.write(0, j, key)
            j += 1
        for bank_line in self.line_ids:
            i = self.createvoucher(conn, excel[0], worksheet, i, number, colnames, bank_line)
            number += 1

        workbook.save(u'voucher.xls')
        self.closeConnection(conn)
        # 生成附件
        f = open('voucher.xls', 'rb')
        self.env['ir.attachment'].create({
            'datas': base64.b64encode(f.read()),
            'name': u'K3导入收付款凭证',
            'datas_fname': u'%s收入凭证.xls' % (self.name.name),
            'res_model': 'bank.form',
            'res_id': self.id, })

    @api.multi
    def createvoucher(self, conn, excel, worksheet, d, number, colnames, bank_line):
        bank_name = self.name.k3_account_name
        bank_code = self.name.k3_account_code
        kehu_name = bank_line.note
        if bank_line.amount_in:
            partner_in_code = self.search_organization(conn, bank_line.name)
        # 入库且有客户名称能找到：一般单证
        if bank_line.amount_in and partner_in_code:
            account_name = bank_name
            account_code = bank_code
            amount = bank_line.amount_in
            note = u"%s收款" % (bank_line.date)
            xiangmu = ''
            d += 1
            self.createvoucherline(amount, excel, number, account_name, account_code, colnames, worksheet, d, note, xiangmu)
            k3_account_id = self.env['bank.form.config'].search([('type','=','in'),('is_normal', '=', True ),('company_id', '=', self.name.company_id.id)], limit = 1)
            account_name2 = k3_account_id.k3_account_name
            account_code2 = k3_account_id.k3_account_code
            ku_name = bank_line.name
            ke_code = partner_in_code[0]
            amount = bank_line.amount_in
            note = u"%s收款" % (bank_line.date)
            xiangmu2 = u'客户---%s---%s' % (ke_code, ku_name)
            d += 1
            self.createvoucherline2(amount, excel, number, account_name2, account_code2, colnames, worksheet, d, note, xiangmu2)
            bank_line.write({'is_voucher': True})
        elif bank_line.amount_in and not bank_line.name and bank_line.note:
            pass
        # 入库且有客户名称找不到：利息收入
        elif bank_line.amount_in and not partner_in_code and bank_line.note:
            k3_account_id = self.env['bank.form.config'].search(
                [('type', '=', 'in'), ('is_normal', '=', False), ('company_id', '=', self.name.company_id.id),
                 ('name', '=', bank_line.note)], limit=1)
            if not k3_account_id:
                return d
            account_name = bank_name
            account_code = bank_code
            amount = bank_line.amount_in
            note = u"%s收款" % (bank_line.date)
            xiangmu = ''
            d += 1
            self.createvoucherline(amount, excel, number, account_name, account_code, colnames, worksheet, d, note,
                                   xiangmu)
            account_name2 = k3_account_id.k3_account_name
            account_code2 = k3_account_id.k3_account_code
            amount = bank_line.amount_in
            note = u"%s收款" % (bank_line.date)
            xiangmu2 = ''
            d += 1
            self.createvoucherline2(amount, excel, number, account_name2, account_code2, colnames, worksheet, d, note,
                                    xiangmu2)
            bank_line.write({'is_voucher': True})

        if bank_line.amount_out:
            partner_out_code = self.search_supplier(conn, bank_line.name)
            # 入库且有客户名称能找到：一般单证
        if bank_line.amount_out and partner_out_code:
            k3_account_id = self.env['bank.form.config'].search(
                [('type', '=', 'out'), ('is_normal', '=', True), ('company_id', '=', self.name.company_id.id)],
                limit=1)
            account_name = k3_account_id.k3_account_name
            account_code = k3_account_id.k3_account_code
            ku_name = bank_line.name
            ke_code = partner_out_code[0]
            amount = bank_line.amount_out
            note = u"%s付款" % (bank_line.date)
            xiangmu = u'供应商---%s---%s' % (ke_code, ku_name)
            d += 1
            self.createvoucherline(amount, excel, number, account_name, account_code, colnames, worksheet, d, note,
                                   xiangmu)

            account_name2 = bank_name
            account_code2 = bank_code
            amount = bank_line.amount_out
            note = u"%s付款" % (bank_line.date)
            xiangmu2 = u''
            d += 1
            self.createvoucherline2(amount, excel, number, account_name2, account_code2, colnames, worksheet, d,
                                    note, xiangmu2)
            bank_line.write({'is_voucher': True})

        elif bank_line.amount_out and bank_line.name and bank_line.note:
            pass
        # 入库且有客户名称找不到：利息收入
        elif bank_line.amount_out and not partner_out_code and bank_line.note:
            k3_account_id = self.env['bank.form.config'].search(
                [('type', '=', 'out'), ('is_normal', '=', False), ('company_id', '=', self.name.company_id.id),
                 ('name', '=', bank_line.note)], limit=1)
            print k3_account_id,bank_line.note
            if not k3_account_id:
                return d
            account_name = k3_account_id.k3_account_name
            account_code = k3_account_id.k3_account_code
            amount = bank_line.amount_out
            note = u"%s付款" % (bank_line.date)
            xiangmu = ''
            d += 1
            self.createvoucherline(amount, excel, number, account_name, account_code, colnames, worksheet, d, note,
                                   xiangmu)
            account_name2 = bank_name
            account_code2 = bank_code
            amount = bank_line.amount_out
            note = u"%s付款" % (bank_line.date)
            xiangmu2 = ''
            d += 1
            self.createvoucherline2(amount, excel, number, account_name2, account_code2, colnames, worksheet, d,
                                    note,
                                    xiangmu2)
            bank_line.write({'is_voucher': True})

        return d

    @api.multi
    def createvoucherline(self, amount, excel, number, account_name, account_code, colnames, worksheet, d, note, xiangmu):
        # 修改内容。
        excel[u'凭证日期'] = excel[u'业务日期']= self.env['finance.period'].get_period_month_date_range(self.period_id)[1]  # 会计期间的最后一天
        excel[u'会计年度'] = self.period_id.year
        excel[u'会计期间'] = self.period_id.month
        excel[u'凭证号'] = excel[u'序号'] = number
        excel[u'科目代码'] = account_code
        excel[u'科目名称'] = account_name
        excel[u'原币金额'] = amount
        excel[u'借方'] = amount
        excel[u'贷方'] = 0
        excel[u'制单'] = u'宣一敏'
        excel[u'凭证摘要'] = note
        excel[u'附件数'] = '1'
        excel[u'分录序号'] = 0
        excel[u'核算项目'] = xiangmu
        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(d, j, excel[key])
            j += 1

    @api.multi
    def createvoucherline2(self, amount,  excel, number, account_name, account_code, colnames, worksheet, d,  note, xiangmu):
        # 修改内容。
        excel[u'凭证日期'] = excel[u'业务日期']= self.env['finance.period'].get_period_month_date_range(self.period_id)[1]  # 会计期间的最后一天
        excel[u'会计年度'] = self.period_id.year
        excel[u'会计期间'] = self.period_id.month
        excel[u'凭证号'] = excel[u'序号'] = number
        excel[u'科目代码'] = account_code
        excel[u'科目名称'] = account_name
        excel[u'原币金额'] = amount
        excel[u'借方'] = 0
        excel[u'贷方'] = amount
        excel[u'制单'] = u'宣一敏'
        excel[u'凭证摘要'] = note
        excel[u'附件数'] = '1'
        excel[u'分录序号'] = 1
        excel[u'核算项目'] = xiangmu
        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(d, j, excel[key])
            j += 1

    # 查询客户代码数据
    @api.multi
    def search_organization(self, conn, name):
        cursor = conn.cursor()
        sql = "select fnumber from t_organization WHERE fname='%s';"
        cursor.execute(sql % name)
        name_code = cursor.fetchone()
        return name_code

    # 查询客户代码数据
    @api.multi
    def search_supplier(self, conn, name):
        cursor = conn.cursor()
        sql = "select fnumber from t_Supplier WHERE fname='%s';"
        cursor.execute(sql % name)
        name_code = cursor.fetchone()
        return name_code

    # 创建数据库连接
    @api.multi
    def createConnection(self):
        if config['k3_server'] and config['k3_server'] != 'None':
            k3_server = config['k3_server']
        else:
            raise Exception('k3 服务没有找到.')
        if config['k3_user'] and config['k3_user'] != 'None':
            k3_user = config['k3_user']
        else:
            raise Exception('k3 用户没有找到.')
        if config['k3_password'] and config['k3_password'] != 'None':
            k3_password = config['k3_password']
        else:
            raise Exception('k3 用户密码没有找到.')
        conn = pymssql.connect(server=k3_server, user=k3_user, password=k3_password, database=self.name.company_id.code, charset='utf8')
        return conn

    # 关闭数据库连接。
    @api.multi
    def closeConnection(self, conn):
        conn.close()

class BankFormLine(models.Model):
    _name = 'bank.form.line'
    _order = "date"

    amount_in = fields.Float(u'收入金额')
    amount_out = fields.Float(u'支出金额')
    name = fields.Char(u'对方户名')
    num = fields.Char(u'对方账/卡号')
    date = fields.Datetime(u'交易时间')
    note = fields.Char(u'摘要')
    purpose  = fields.Text(u'用途')
    is_voucher = fields.Boolean(string=u'已生成过凭证', default=False)
    order_id = fields.Many2one('bank.form', u'对应对帐单', index=True, copy=False, readonly=True)

class BankFormConfig(models.Model):
    _name = 'bank.form.config'
    _order = "name"

    name = fields.Char(u'摘要内容')
    note = fields.Char(u'对方户名')
    k3_account_code = fields.Char(u'k3科目代码', required=True, )
    k3_account_name = fields.Char(u'k3科目名称', required=True, )
    type = fields.Selection([('in', u'收入'),
                              ('out', u'支出')], u'收支状态', default='in')
    is_normal = fields.Boolean(string=u'正常收支', default=True)
    company_id = fields.Many2one('k3.category',u'对应公司',required=True)

class BankAccount(models.Model):
    _inherit = 'bank.account'

    k3_account_code = fields.Char(u'k3科目代码', required=True, )
    k3_account_name = fields.Char(u'k3科目名称', required=True, )
    company_id = fields.Many2one('k3.category',u'对应公司',required=True)