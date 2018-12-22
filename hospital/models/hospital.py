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

# 字段只读状态
READONLY_STATES = {
        'done': [('readonly', True)],
    }

class HospitalMonth(models.Model):
    '''医疗月度统计'''
    _name = 'hospital.month'
    _order = "name"

    name = fields.Many2one(
        'finance.period',
        u'会计期间',
        ondelete='restrict',
        required=True,
        states=READONLY_STATES)
    begin_id =  fields.Integer(u'起始凭证号', required=True, copy=False,)
    invoice_ids = fields.One2many('hospital.invoice', 'month_id', u'医院发票明细',
                               states=READONLY_STATES, copy=False)
    state = fields.Selection([('draft', u'草稿'),
                              ('done', u'已结束')], u'状态', default='draft')
    attachment_number = fields.Integer(compute='_compute_attachment_number', string=u'附件号')

    @api.multi
    def action_get_attachment_view(self):
        res = self.env['ir.actions.act_window'].for_xml_id('base', 'action_attachment')
        res['domain'] = [('res_model', '=', 'hospital.month'), ('res_id', 'in', self.ids)]
        res['context'] = {'default_res_model': 'hospital.month', 'default_res_id': self.id}
        return res

    @api.multi
    def _compute_attachment_number(self):
        attachment_data = self.env['ir.attachment'].read_group(
            [('res_model', '=', 'hospital.month'), ('res_id', 'in', self.ids)], ['res_id'], ['res_id'])
        attachment = dict((data['res_id'], data['res_id_count']) for data in attachment_data)
        for expense in self:
            expense.attachment_number = attachment.get(expense.id, 0)

    # COPY excel
    @api.multi
    def worksheetcopy(self, worksheet1, worksheet2):
        ncows = worksheet1.nrows
        ncols = worksheet1.ncols
        for i in range(0, ncows):
            row = worksheet1.row_values(i)
            for j in range(0, ncols):
                worksheet2.write(i, j, row[j])

    # 读取excel
    @api.multi
    def readexcel(self, table):
        ncows = table.nrows
        ncols = 0
        colnames = table.row_values(0)
        list = []
        for rownum in range(1, ncows):
            row = table.row_values(rownum)
            if row:
                app = {}
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                list.append(app)
                ncols += 1
        return list, colnames

    @api.multi
    def createvoucher(self, conn, excel, worksheet, d, number, colnames, invoice):
        x = 0
        app = {u'社保记账':0, u'现金支付':0}
        kehu_name = invoice.name
        kehu_code = self.search_organization(conn, kehu_name)

        for line in invoice.pay_ids:
            if line.name == u'社保记账':
                app[u'社保记账'] += line.amount
            else:
                app[u'现金支付'] += line.amount

        if app[u'社保记账']:
            account_name2 = u'社保记账'
            amount2 = app[u'社保记账']
            d += 1
            self.createvoucherline(account_name2, amount2, x, excel, number, invoice, colnames, worksheet, d, False)
            x += 1
        if app[u'现金支付']:
            account_name = u'现金支付'
            amount = app[u'现金支付']
            d += 1
            self.createvoucherline(account_name, amount, x, excel, number, invoice, colnames, worksheet, d, kehu_code)
            x += 1

        for line in invoice.line_ids:
            account_name = line.name
            amount = line.amount
            d += 1
            self.createvoucherline2(account_name, amount, x, excel, number, invoice, colnames, worksheet,d)
            x += 1

        return d

    @api.multi
    def createvoucherline(self, account_name, amount, x, excel, number, invoice, colnames, worksheet, d, kehu_code):
        # 修改内容。
        excel[u'凭证日期'] = excel[u'业务日期']= self.env['finance.period'].get_period_month_date_range(self.name)[1]  # 会计期间的最后一天
        excel[u'会计年度'] = self.name.year
        excel[u'会计期间'] = self.name.month
        excel[u'凭证号'] = excel[u'序号'] = number
        excel[u'科目代码'] = self.env['hospital.config'].search([('name', '=', account_name),('type', '=', invoice.type)], limit=1).k3_account_code
        excel[u'科目名称'] = self.env['hospital.config'].search([('name', '=', account_name),('type', '=', invoice.type)], limit=1).k3_account_name
        excel[u'原币金额'] = amount
        excel[u'借方'] = amount
        excel[u'贷方'] = 0
        excel[u'制单'] = u'宣一敏'
        excel[u'凭证摘要'] = u'%s发票%s'%(self.env['finance.period'].get_period_month_date_range(self.name)[1], invoice.invoice)
        excel[u'附件数'] = '1'
        excel[u'分录序号'] = x

        if kehu_code:
            excel[u'核算项目'] = u'客户---%s---%s' % (kehu_code[0], invoice.name)
        else:
            excel[u'核算项目'] = u''
        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(d, j, excel[key])
            j += 1

    @api.multi
    def createvoucherline2(self, account_name, amount, x, excel, number, invoice, colnames, worksheet,d):
        if not self.env['hospital.config'].search([('name', '=', account_name),('type', '=', invoice.type)]):
            raise UserError(('请到系统增加发票设置%s。'% (account_name)))
        # 修改内容。
        excel[u'凭证日期'] = excel[u'业务日期'] = self.env['finance.period'].get_period_month_date_range(self.name)[
            1]  # 会计期间的最后一天
        excel[u'会计年度'] = self.name.year
        excel[u'会计期间'] = self.name.month
        excel[u'凭证号'] = excel[u'序号'] = number
        excel[u'科目代码'] = self.env['hospital.config'].search([('name', '=', account_name),('type', '=', invoice.type)]).k3_account_code
        excel[u'科目名称'] = self.env['hospital.config'].search([('name', '=', account_name),('type', '=', invoice.type)]).k3_account_name
        excel[u'原币金额'] = amount
        excel[u'借方'] = 0
        excel[u'贷方'] = amount
        excel[u'制单'] = u'宣一敏'
        excel[u'凭证摘要'] = u'%s发票%s' % (self.env['finance.period'].get_period_month_date_range(self.name)[1], account_name)
        excel[u'附件数'] = '1'
        excel[u'分录序号'] = x
        excel[u'核算项目'] = ''
        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(d, j, excel[key])
            j += 1

    #合并正负发票
    @api.multi
    def merge_positive_negative(self):
        for invoice in self.invoice_ids:
            if not invoice.invoice :
                merge_id =  self.env['hospital.invoice'].search([('name', '=', invoice.name),('amount', '=', -invoice.amount)],limit=1)
                if not merge_id:
                    raise UserError(u'请确认此单据有对应正数发票病人：%s, 金额：%s' %(invoice.name,invoice.amount))
                invoice.write({'is_red': True,
                                'note': merge_id.invoice})
                merge_id.write({'is_red': True})


    # 导出K3收入凭证
    @api.multi
    def exp_k3_voucher(self, order=False):
        self.merge_positive_negative()
        xls_data = xlrd.open_workbook('./excel/moban.xls')
        Page1 = xls_data.sheet_by_name('Page1')
        Page2 = xls_data.sheet_by_name('t_Schema')
        # 连接数据库
        conn = self.createConnection()
        excel, colnames = self.readexcel(Page1)  # 读模版，返回字典及表头数组
        workbook = xlwt.Workbook(encoding='utf-8')  # 生成文件
        worksheet = workbook.add_sheet(u'Page1')  # 在文件中创建一个名为Page1的sheet
        worksheet2 = workbook.add_sheet(u't_Schema')
        self.worksheetcopy(Page2, worksheet2)

        i = j = 0
        number = self.begin_id
        for key in colnames:
            worksheet.write(0, j, key)
            j += 1
        for invoice in self.invoice_ids:
            if not invoice.is_red:
                i = self.createvoucher(conn, excel[0], worksheet, i, number, colnames, invoice)
                number += 1

        workbook.save(u'voucher.xls')
        self.closeConnection(conn)
        # 生成附件
        f = open('voucher.xls', 'rb')
        self.env['ir.attachment'].create({
            'datas': base64.b64encode(f.read()),
            'name': u'K3导入收入凭证',
            'datas_fname': u'%s收入凭证.xls' % (self.name.name),
            'res_model': 'hospital.month',
            'res_id': self.id, })

    #导出K3客户
    @api.multi
    def createsalepartner(self, conn, excel, worksheet, code, x, colnames, invoice):

        for i in excel:
            # 修改内容。
            i[u'名称'] = invoice.name  # 名称
            i[u'代码'] = code  # 名称

        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(x, j, i[key])
            j += 1

    #导出K3客户
    @api.multi
    def exp_k3_sale_partner(self, order=False):
        xls_data = xlrd.open_workbook('./excel/sale_partner.xls')
        Page1 = xls_data.sheet_by_name('Page1')
        Page2 = xls_data.sheet_by_name('Page2')
        Page4 = xls_data.sheet_by_name('t_Schema')
        # 连接数据库
        conn = self.createConnection()
        excel, colnames = self.readexcel(Page1)  # 读模版，返回字典及表头数组
        workbook = xlwt.Workbook(encoding='utf-8')  # 生成文件
        worksheet = workbook.add_sheet(u'Page1')  # 在文件中创建一个名为Page1的sheet
        worksheet2 = workbook.add_sheet(u'Page2')
        self.worksheetcopy(Page2, worksheet2)
        worksheet4 = workbook.add_sheet(u't_Schema')
        self.worksheetcopy(Page4, worksheet4)

        i = j = 0
        number = self.search_max_salefnumber(conn)[0]
        c = number.split('.')
        if len(c) == 2:
            a, b = c
        elif len(c) == 3:
            o, a, b = c
        changdu = len(b)
        for key in colnames:
            worksheet.write(0, j, key)
            j += 1
        partner = []
        for invoice in self.invoice_ids:
            if invoice.name not in partner:
                name_code = self.search_organization(conn, invoice.name)
                partner.append(invoice.name)
                if not name_code:
                    i += 1
                    x = int(b) + i
                    changdu2 = len(str(x))
                    j = b[0:(changdu - changdu2)] + str(x)
                    code = '%s.%s' % (a, j)
                    self.createsalepartner(conn, excel, worksheet, code, i, colnames, invoice)

        workbook.save(u'sale_partner.xls')
        self.closeConnection(conn)
        # 生成附件
        f = open('sale_partner.xls', 'rb')
        self.env['ir.attachment'].create({
            'datas': base64.b64encode(f.read()),
            'name': u'K3新增客户',
            'datas_fname': u'%s客户.xls' % (self.name.name),
            'res_model': 'hospital.month',
            'res_id': self.id, })

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
        conn = pymssql.connect(server=k3_server, user=k3_user, password=k3_password, database='AIS20180713180714',
                               charset='utf8')
        return conn

    # 关闭数据库连接。
    @api.multi
    def closeConnection(self, conn):
        conn.close()

    # 查询客户代码数据
    @api.multi
    def search_organization(self, conn, name):
        cursor = conn.cursor()
        sql = "select fnumber from t_organization WHERE fname='%s';"
        cursor.execute(sql % name)
        name_code = cursor.fetchone()
        return name_code

    # 查询客户代码最大编号
    @api.multi
    def search_max_salefnumber(self, conn):
        cursor = conn.cursor()
        cursor.execute("select max(fnumber) from t_organization ;")
        fnumber = cursor.fetchone()
        return fnumber

    @api.multi
    def button_excel(self):
        return {
            'name': u'引入excel',
            'view_mode': 'form',
            'view_type': 'form',
            'res_model': 'create.hospital.invoice.wizard',
            'type': 'ir.actions.act_window',
            'target': 'new',
        }

class HospitalInvoice(models.Model):
    '''医疗发票'''
    _name = 'hospital.invoice'
    _order = "type,invoice"

    month_id = fields.Many2one('hospital.month', u'对应入帐月份', index=True, copy=False, readonly=True)

    internal_id = fields.Char(u'结帐ID', )
    name = fields.Char(u'病人姓名', required=True,)
    invoice = fields.Char(u'票据号', )
    name_id = fields.Char(u'身份证号',)
    amount = fields.Float(u'结帐金额',)
    pay_type = fields.Char(u'病人类别')
    difference_amount = fields.Float(u'尾数处理',)
    type = fields.Selection([('hospitalization', u'住院'),
                              ('outpatient', u'门诊')], u'单据类型', default='hospitalization')
    line_ids = fields.One2many('hospital.invoice.line', 'invoice_id', u'医院发票收入明细',
                                copy=False)
    pay_ids = fields.One2many('hospital.pay.line', 'invoice_id', u'医院发票收款明细',
                                copy=False)
    cost_ids = fields.One2many('hospital.invoice.cost', 'invoice_id', u'医院发票费用明细',
                              copy=False)
    is_red = fields.Boolean(u'不生成单据',
                            help=u'此单据不生成对应单据')
    note = fields.Text(u'备注')

class HospitalInvoiceLine(models.Model):
    '''医疗发票明细'''
    _name = 'hospital.invoice.line'
    _order = "name"

    invoice_id = fields.Many2one('hospital.invoice', u'医院发票', index=True, copy=False, readonly=True)
    name = fields.Char(u'收入内容', required = True)
    amount = fields.Float(u'金额')
    type = fields.Selection([('hospitalization', u'住院'),
                             ('outpatient', u'门诊')], u'单据类型', )

class HospitalPayLine(models.Model):
    '''医疗发票付款明细'''
    _name = 'hospital.pay.line'

    invoice_id = fields.Many2one('hospital.invoice', u'医院发票', index=True, copy=False, readonly=True)
    name = fields.Char(u'收款方式', required=True)
    amount = fields.Float(u'金额')
    type = fields.Selection([('hospitalization', u'住院'),
                             ('outpatient', u'门诊')], u'单据类型', )

class HospitalInvoiceCost(models.Model):
    '''医疗发票明细'''
    _name = 'hospital.invoice.cost'
    _order = "cost_type,cost_time"

    invoice_id = fields.Many2one('hospital.invoice', u'医院发票', index=True, copy=False, readonly=True)
    cost_type = fields.Char(u'发票类别')
    name = fields.Char(u'药品名称')
    name2 = fields.Char(u'规格')
    number = fields.Float(u'数量')
    unit= fields.Char(u'单位')
    amount = fields.Float(u'金额')
    cost_time = fields.Datetime(u'业务时间')
    type = fields.Selection([('hospitalization', u'住院'),
                             ('outpatient', u'门诊')], u'单据类型', )

class HospitalConfig(models.Model):
    '''医疗发票'''
    _name = 'hospital.config'
    _order = "name"

    name = fields.Char(u'收入名称', required=True, )
    k3_account_code = fields.Char(u'k3科目代码', required=True, )
    k3_account_name = fields.Char(u'k3科目名称', required=True, )
    type = fields.Selection([('hospitalization', u'住院'),
                             ('outpatient', u'门诊')], u'单据类型', required=True, default='hospitalization')

class HospitalCash(models.Model):
    '''医疗预付款收据'''
    _name = 'hospital.cash'
    _order = "name"

    number = fields.Char(u'票据号',)
    name = fields.Char(u'姓名')
    amount = fields.Float(u'金额')
    type = fields.Char(u'支付方式')
    date = fields.Datetime(u'操作时间')
    is_red = fields.Boolean(u'不生成单据',
                            help=u'此单据不生成对应单据')
    note = fields.Text(u'备注')
    month_id = fields.Many2one('hospital.cash.month', u'对应入帐月份', index=True, copy=False, readonly=True)

class HospitalCashMonth(models.Model):
    '''医疗月度统计'''
    _name = 'hospital.cash.month'
    _order = "name"

    name = fields.Many2one(
        'finance.period',
        u'会计期间',
        ondelete='restrict',
        required=True,
        states=READONLY_STATES)
    begin_id =  fields.Integer(u'起始凭证号', required=True, copy=False,)
    cash_ids = fields.One2many('hospital.cash', 'month_id', u'医院预收款明细',
                               states=READONLY_STATES, copy=False)
    state = fields.Selection([('draft', u'草稿'),
                              ('done', u'已结束')], u'状态', default='draft')
    attachment_number = fields.Integer(compute='_compute_attachment_number', string=u'附件号')

    @api.multi
    def action_get_attachment_view(self):
        res = self.env['ir.actions.act_window'].for_xml_id('base', 'action_attachment')
        res['domain'] = [('res_model', '=', 'hospital.cash.month'), ('res_id', 'in', self.ids)]
        res['context'] = {'default_res_model': 'hospital.cash.month', 'default_res_id': self.id}
        return res

    @api.multi
    def _compute_attachment_number(self):
        attachment_data = self.env['ir.attachment'].read_group(
            [('res_model', '=', 'hospital.cash.month'), ('res_id', 'in', self.ids)], ['res_id'], ['res_id'])
        attachment = dict((data['res_id'], data['res_id_count']) for data in attachment_data)
        for expense in self:
            expense.attachment_number = attachment.get(expense.id, 0)

    @api.multi
    def search_coustorm(self, conn, name_id):
        cursor = conn.cursor()
        sql = "select VAA05,VAA15 from VAA1 WHERE VAA01='%s';"
        cursor.execute(sql % name_id)
        name_code = cursor.fetchone()
        return name_code

    @api.multi
    def synchro_hospital_cash(self):
        conn = self.hospitalcreateConnection()
        cursor = conn.cursor()

        star_date, end_date = self.env['finance.period'].get_period_month_date_range(self.name)
        star_date = '%s 00:00:00' % (star_date)
        end_date = '%s 23:59:59' % (end_date)
        sql = "select VAA01,VBL03,VBL13,VBL14,VBL18 from V_VBL_FULL WHERE VBL04='4' and VBL27='1' and VBL18>='%s' and VBL18<'%s';"
        cursor.execute(sql % (star_date, end_date))
        cash_ids = cursor.fetchall()
        for line in cash_ids:
            user_id, cash_number, amount, cash_type, date = line
            name, name_id = self.search_coustorm(conn, user_id)
            self.env['hospital.cash'].create({
                'number': cash_number,
                'name': name.encode('latin-1').decode('gbk'),
                'amount': amount,
                'type': cash_type.encode('latin-1').decode('gbk'),
                'date': date,
                'month_id': self.id })
        self.closeConnection(conn)

    # 创建数据库连接
    @api.multi
    def hospitalcreateConnection(self):
        if config['amj_server'] and config['amj_server'] != 'None':
            amj_server = config['amj_server']
        else:
            raise Exception('医院服务器没有找到.')
        if config['amj_user'] and config['amj_user'] != 'None':
            amj_user = config['amj_user']
        else:
            raise Exception('医院用户没有找到.')
        if config['amj_password'] and config['amj_password'] != 'None':
            amj_password = config['amj_password']
        else:
            raise Exception('医院 用户密码没有找到.')
        if config['amj_database'] and config['amj_database'] != 'None':
            amj_database = config['amj_database']
        else:
            raise Exception('医院数据库没有找到.')
        conn = pymssql.connect(server=amj_server, user=amj_user, password=amj_password, database=amj_database,
                               charset='utf8')
        return conn

    # 关闭数据库连接。
    @api.multi
    def closeConnection(self, conn):
        conn.close()

    # 导出K3客户
    @api.multi
    def exp_k3_sale_partner(self, order=False):
        xls_data = xlrd.open_workbook('./excel/sale_partner.xls')
        Page1 = xls_data.sheet_by_name('Page1')
        Page2 = xls_data.sheet_by_name('Page2')
        Page4 = xls_data.sheet_by_name('t_Schema')
        # 连接数据库
        conn = self.env['hospital.month'].createConnection()
        excel, colnames = self.env['hospital.month'].readexcel(Page1)  # 读模版，返回字典及表头数组
        workbook = xlwt.Workbook(encoding='utf-8')  # 生成文件
        worksheet = workbook.add_sheet(u'Page1')  # 在文件中创建一个名为Page1的sheet
        worksheet2 = workbook.add_sheet(u'Page2')
        self.env['hospital.month'].worksheetcopy(Page2, worksheet2)
        worksheet4 = workbook.add_sheet(u't_Schema')
        self.env['hospital.month'].worksheetcopy(Page4, worksheet4)
        i = j = 0
        number = self.env['hospital.month'].search_max_salefnumber(conn)[0]
        c = number.split('.')
        if len(c) == 2:
            a, b = c
        elif len(c) == 3:
            o, a, b = c
        changdu = len(b)
        for key in colnames:
            worksheet.write(0, j, key)
            j += 1
        partner = []
        for cash in self.cash_ids:
            if cash.name not in partner:
                name_code = self.env['hospital.month'].search_organization(conn, cash.name)
                partner.append(cash.name)
                if not name_code:
                    i += 1
                    x = int(b) + i
                    changdu2 = len(str(x))
                    j = b[0:(changdu - changdu2)] + str(x)
                    code = '%s.%s' % (a, j)
                    self.env['hospital.month'].createsalepartner(conn, excel, worksheet, code, i, colnames, cash)

        workbook.save(u'sale_partner.xls')
        self.env['hospital.cash.month'].closeConnection(conn)
        # 生成附件
        f = open('sale_partner.xls', 'rb')
        self.env['ir.attachment'].create({
            'datas': base64.b64encode(f.read()),
            'name': u'K3新增客户',
            'datas_fname': u'%s客户.xls' % (self.name.name),
            'res_model': 'hospital.cash.month',
            'res_id': self.id, })

    # 导出K3收入凭证
    @api.multi
    def exp_k3_cash_voucher(self, order=False):
        self.merge_positive_negative()
        xls_data = xlrd.open_workbook('./excel/moban.xls')
        Page1 = xls_data.sheet_by_name('Page1')
        Page2 = xls_data.sheet_by_name('t_Schema')
        # 连接数据库
        conn = self.env['hospital.month'].createConnection()
        excel, colnames = self.env['hospital.month'].readexcel(Page1)  # 读模版，返回字典及表头数组
        workbook = xlwt.Workbook(encoding='utf-8')  # 生成文件
        worksheet = workbook.add_sheet(u'Page1')  # 在文件中创建一个名为Page1的sheet
        worksheet2 = workbook.add_sheet(u't_Schema')
        self.env['hospital.month'].worksheetcopy(Page2, worksheet2)

        i = j = 0
        number = self.begin_id
        for key in colnames:
            worksheet.write(0, j, key)
            j += 1
        for line in self.cash_ids:
            if not line.is_red:
                i = self.createvoucher(conn, excel[0], worksheet, i, number, colnames, line)
                number += 1

        workbook.save(u'voucher.xls')
        self.closeConnection(conn)
        # 生成附件
        f = open('voucher.xls', 'rb')
        self.env['ir.attachment'].create({
            'datas': base64.b64encode(f.read()),
            'name': u'K3导入收款凭证',
            'datas_fname': u'%s收款凭证.xls' % (self.name.name),
            'res_model': 'hospital.cash.month',
            'res_id': self.id, })

    # 合并正负发票
    @api.multi
    def merge_positive_negative(self):
        for line in self.cash_ids:
            if line.amount <= 0:
                merge_id = self.env['hospital.cash'].search(
                    [('name', '=', line.name), ('amount', '=', -line.amount)], limit=1)
                if not merge_id:
                    raise UserError(u'请确认此单据有对应正数预收款病人：%s, 金额：%s' % (line.name, line.amount))
                line.write({'is_red': True,
                               'note': merge_id.number})
                merge_id.write({'is_red': True})

    @api.multi
    def createvoucher(self, conn, excel, worksheet, d, number, colnames, cash):
        x = j = 0
        kehu_name = cash.name
        kehu_code = self.env['hospital.month'].search_organization(conn, kehu_name)
        if not self.env['hospital.cash.config'].search([('name', '=', cash.type)]):
            raise UserError(('请到系统增加发票设置%s。'% (cash.type)))
        # 修改内容。
        excel[u'凭证日期'] = excel[u'业务日期'] = self.env['finance.period'].get_period_month_date_range(self.name)[
            1]  # 会计期间的最后一天
        excel[u'会计年度'] = self.name.year
        excel[u'会计期间'] = self.name.month
        excel[u'凭证号'] = excel[u'序号'] = number
        excel[u'科目代码'] = self.env['hospital.cash.config'].search([('name', '=', cash.type)]).k3_account_code
        excel[u'科目名称'] = self.env['hospital.cash.config'].search([('name', '=', cash.type)]).k3_account_name
        excel[u'原币金额'] = cash.amount
        excel[u'借方'] = cash.amount
        excel[u'贷方'] = 0
        excel[u'制单'] = u'宣一敏'
        excel[u'凭证摘要'] = u'%s预收款%s' % (cash.date,cash.number)
        excel[u'附件数'] = '1'
        excel[u'分录序号'] = 0
        excel[u'核算项目'] = ''
        d += 1
        for key in colnames:
            # 写入excel
            worksheet.write(d, j, excel[key])
            x += 1
            j += 1
        excel[u'凭证日期'] = excel[u'业务日期'] = self.env['finance.period'].get_period_month_date_range(self.name)[
            1]  # 会计期间的最后一天
        excel[u'会计年度'] = self.name.year
        excel[u'会计期间'] = self.name.month
        excel[u'凭证号'] = excel[u'序号'] = number
        excel[u'科目代码'] = 2203.02
        excel[u'科目名称'] = u'预交款'
        excel[u'原币金额'] = cash.amount
        excel[u'借方'] = 0
        excel[u'贷方'] = cash.amount
        excel[u'制单'] = u'宣一敏'
        excel[u'凭证摘要'] = u'%s预收款%s' % (cash.date, cash.number)
        excel[u'附件数'] = '1'
        excel[u'分录序号'] = 1
        excel[u'核算项目'] = u'客户---%s---%s' % (kehu_code[0], kehu_name)
        d += 1
        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(d, j, excel[key])
            x += 1
            j += 1
        return d

class HospitalCashConfig(models.Model):
    '''医疗发票'''
    _name = 'hospital.cash.config'
    _order = "name"

    name = fields.Char(u'收入名称', required=True, )
    k3_account_code = fields.Char(u'k3科目代码', required=True, )
    k3_account_name = fields.Char(u'k3科目名称', required=True, )
