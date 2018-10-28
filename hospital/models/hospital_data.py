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

class create_hospital_invoice_wizard(models.TransientModel):
    _name = 'create.hospital.invoice.wizard'
    _description = 'Hospital Invoice Import'

    excel = fields.Binary(u'导入系统导出的excel文件',)
    excel2 = fields.Binary(u'导入系统导出的excel文件', )
    excel3 = fields.Binary(u'导入系统导出的excel文件', )
    type = fields.Selection([('hospitalization', u'住院'),
                              ('outpatient', u'门诊')], u'单据类型', default='hospitalization')

    @api.multi
    def create_hospital_invoice(self):
        """
        通过Excel文件导入信息到hospital.invoice
        """
        month = self.env['hospital.month'].browse(self.env.context.get('active_id'))
        if not month:
            return {}
        xls_data = xlrd.open_workbook(
                file_contents=base64.decodestring(self.excel))
        table = xls_data.sheets()[0]
        #取得行数
        ncows = table.nrows
        #取得第1行数据
        colnames =  table.row_values(0)
        list =[]
        newcows = 0
        for rownum in range(1,ncows):
            row = table.row_values(rownum)
            if row:
                app = {}
                for i in range(len(colnames)):
                   app[colnames[i]] = row[i]
                #过滤掉不需要的行，详见销货清单的会在清单中再次导入
                if app.get(u'病人姓名') or app.get(u'姓名'):
                    list.append(app)
                    newcows += 1
        #数据读入。
        for data in range(0,newcows):
            in_xls_data = list[data]
            invoice_ids = self.env['hospital.invoice'].create({
                    'name': in_xls_data.get(u'病人姓名') or in_xls_data.get(u'姓名'),
                    'invoice': in_xls_data.get(u'票据号'),
                    'name_id': in_xls_data.get(u'身份证号'),
                    'pay_type': in_xls_data.get(u'病人类别'),
                    'amount': float(in_xls_data.get(u'结账金额') or 0.00),
                    'difference_amount': float(in_xls_data.get(u'尾数处理') or 0.00),
                    'type':self.type,
                    'month_id':month.id or '',})
            if invoice_ids.type == 'outpatient' and in_xls_data.get(u'结账金额') != in_xls_data.get(u'应收金额'):
                invoice_ids.write({'difference_amount': float(in_xls_data.get(u'应收金额') - in_xls_data.get(u'结账金额'))})
            if in_xls_data.get(u'冲预交'):
                self.env['hospital.pay.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name':u'冲预交',
                    'amount': float(in_xls_data.get(u'冲预交') or 0.00),
                    'type':self.type,})
            if in_xls_data.get(u'个人自付'):
                self.env['hospital.pay.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'个人自付',
                    'amount': float(in_xls_data.get(u'个人自付') or 0.00),
                    'type': self.type,})

            if in_xls_data.get(u'应收金额'):
                self.env['hospital.pay.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'个人自付',
                    'amount': float(in_xls_data.get(u'应收金额') or 0.00),
                    'type': self.type,})

            if in_xls_data.get(u'材料费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'材料费',
                    'amount': float(in_xls_data.get(u'材料费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'床位费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'床位费',
                    'amount': float(in_xls_data.get(u'床位费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'护理费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'护理费',
                    'amount': float(in_xls_data.get(u'护理费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'检查费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'检查费',
                    'amount': float(in_xls_data.get(u'检查费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'检验费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'检验费',
                    'amount': float(in_xls_data.get(u'检验费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'输氧费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'输氧费',
                    'amount': float(in_xls_data.get(u'输氧费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'西药费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'西药费',
                    'amount': float(in_xls_data.get(u'西药费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'诊查费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'诊查费',
                    'amount': float(in_xls_data.get(u'诊查费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'治疗费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'治疗费',
                    'amount': float(in_xls_data.get(u'治疗费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'中成药费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'中成药费',
                    'amount': float(in_xls_data.get(u'中成药费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'中草药费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'中草药费',
                    'amount': float(in_xls_data.get(u'中草药费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'其他费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'其他费',
                    'amount': float(in_xls_data.get(u'其他费') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'手术费'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'手术费',
                    'amount': float(in_xls_data.get(u'手术费') or 0.00),
                    'type': self.type,})

            if in_xls_data.get(u'其他'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'其他',
                    'amount': float(in_xls_data.get(u'其他') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'检验'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'检验',
                    'amount': float(in_xls_data.get(u'检验') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'中成药'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'中成药',
                    'amount': float(in_xls_data.get(u'中成药') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'西药'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'西药',
                    'amount': float(in_xls_data.get(u'西药') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'检查'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'检查',
                    'amount': float(in_xls_data.get(u'检查') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'卫生材料'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'卫生材料',
                    'amount': float(in_xls_data.get(u'卫生材料') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'治疗'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'治疗',
                    'amount': float(in_xls_data.get(u'治疗') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'中草药'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'中草药',
                    'amount': float(in_xls_data.get(u'中草药') or 0.00),
                    'type': self.type,})
            if in_xls_data.get(u'护理'):
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_ids.id,
                    'name': u'护理',
                    'amount': float(in_xls_data.get(u'护理') or 0.00),
                    'type': self.type,})

    def excel_date(self,data):
        #将excel日期改为正常日期
        if type(data) in (int,float):
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(data,0)
            py_date = datetime.datetime(year, month, day, hour, minute, second)
        else:
            py_date = data
        return py_date

    @api.multi
    def synchro_hospital_invoice(self):
        month = self.env['hospital.month'].browse(self.env.context.get('active_id'))
        if not month:
            return {}
        conn = self.createConnection()
        self.create_hospital_invoice2(conn, self.type, month)
        self.closeConnection(conn)

    @api.multi
    def createConnection(self):
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

    # 创建发票数据
    @api.multi
    def create_hospital_invoice2(self, conn, type, period):
        cursor = conn.cursor()
        if type == 'outpatient':
            sql = "select vak01,vaa01,fab03,vak08,vak07,vaa07 from VAK1 WHERE vak06='2' and vak13>='%s' and vak13<'%s' ;"
        else:
            sql = "select vak01,vaa01,fab03,vak08,vak07,vaa07 from VAK1 WHERE vak06>'2' and vak13>='%s' and vak13<'%s' ;"
        star_date,end_date = self.env['finance.period'].get_period_month_date_range(period.name)
        star_date = '%s 00:00:00'%(star_date)
        end_date =  '%s 23:59:59'%(end_date)
        cursor.execute(sql % (star_date,end_date))
        invoice_ids = cursor.fetchall()
        for invoice in invoice_ids:
            internal_id,coustorm_id,invoice_name,amount,difference_amount,caseid = invoice
            name,name_id = self.search_coustorm(conn, coustorm_id)
            if self.search_pay_type(conn, caseid):
                pay_type = self.search_pay_type(conn, caseid)[0]
            else:
                pay_type = ''
            invoice_id = self.env['hospital.invoice'].create({
                'internal_id':internal_id,
                'name': name.encode('latin-1').decode('gbk'),
                'invoice': invoice_name,
                'name_id': name_id and name_id.encode('latin-1').decode('gbk') or '',
                'pay_type': pay_type.encode('latin-1').decode('gbk'),
                'amount': amount,
                'difference_amount': difference_amount,
                'type': self.type,
                'month_id': period.id or '', })
            if type == 'outpatient':
                self.crate_hospital_invoice_cost2(conn, invoice_id, internal_id)
            else:
                self.crate_hospital_invoice_cost(conn, invoice_id, internal_id)

            self.crate_hospital_invoice_pay(conn, invoice_id, internal_id)
            self.crate_hospital_invoice_line(conn, invoice_id)

    @api.multi
    def crate_hospital_invoice_line(self, conn, invoice_id):
        # 合并正负发票
        invoice_line = []
        for line in invoice_id.cost_ids:
            if line.cost_type in invoice_line:
                old_amount = self.env['hospital.invoice.line'].search(
                    [('invoice_id', '=', invoice_id.id), ('name', '=', line.cost_type)], limit=1)
                amount = old_amount.amount + line.amount
                old_amount.write({'amount': amount})
            else:
                self.env['hospital.invoice.line'].create({
                    'invoice_id': invoice_id.id,
                    'name': line.cost_type,
                    'amount': line.amount,
                    'type': invoice_id.type,
                })
                invoice_line.append(line.cost_type)

    @api.multi
    def crate_hospital_invoice_cost(self, conn, invoice_id, internal_id):
        cursor = conn.cursor()
        sql = "select BBY01,VAJ36,VAJ35,VAJ25,VAJ46 from VAJ2 WHERE ACF01 = '2' AND vak01='%s';"
        cursor.execute(sql % internal_id)
        cost_ids = cursor.fetchall()
        for cost_id in cost_ids:
            m,amount,unit,number,cost_time = cost_id
            code, name, name2 = self.search_cost_nameall(conn, m)
            cost_type = self.search_cost_type(conn, code)[0]
            self.env['hospital.invoice.cost'].create({
                'invoice_id': invoice_id.id,
                'cost_type': cost_type.encode('latin-1').decode('gbk'),
                'name': name.encode('latin-1').decode('gbk'),
                'name2': name2 and name2.encode('latin-1').decode('gbk') or '',
                'number': number,
                'unit': unit.encode('latin-1').decode('gbk'),
                'amount': amount,
                'cost_time': cost_time,
                'type': self.type,
            })

    @api.multi
    def crate_hospital_invoice_cost2(self, conn, invoice_id, internal_id):
        cursor = conn.cursor()
        sql = "select BBY01,VAJ38,VAJ35,VAJ25,VAJ46 from VAJ1 WHERE VAK01='%s';"
        cursor.execute(sql % internal_id)
        cost_ids = cursor.fetchall()
        for cost_id in cost_ids:
            print cost_id
            m, amount, unit, number, cost_time = cost_id
            code, name, name2 = self.search_cost_nameall(conn, m)
            cost_type = self.search_cost_type(conn, code)[0]
            self.env['hospital.invoice.cost'].create({
                'invoice_id': invoice_id.id,
                'cost_type': cost_type.encode('latin-1').decode('gbk'),
                'name': name.encode('latin-1').decode('gbk'),
                'name2': name2 and name2.encode('latin-1').decode('gbk') or '',
                'number': number,
                'unit': unit.encode('latin-1').decode('gbk'),
                'amount': amount,
                'cost_time': cost_time,
                'type': self.type,
            })

    @api.multi
    def crate_hospital_invoice_pay(self, conn, invoice_id, internal_id):
        cursor = conn.cursor()
        sql = "select VBL14,VBL13 from VBL1 WHERE vak01='%s';"
        cursor.execute(sql % internal_id)
        pay_ids = cursor.fetchall()
        for pay_id in pay_ids:
            name,amount = pay_id
            self.env['hospital.pay.line'].create({
                'invoice_id': invoice_id.id,
                'name': name.encode('latin-1').decode('gbk'),
                'amount': amount,
                'type': self.type, })
        return True

    # 查询病人信息数据
    @api.multi
    def search_coustorm(self, conn, name_id):
        cursor = conn.cursor()
        sql = "select VAA05,VAA15 from VAA1 WHERE VAA01='%s';"
        cursor.execute(sql % name_id)
        name_code = cursor.fetchone()
        return name_code

    # 查询病人信息数据
    @api.multi
    def search_pay_type(self, conn, name_id):
        cursor = conn.cursor()
        sql = "select BDP02 from VAE1 WHERE VAE01='%s';"
        cursor.execute(sql % name_id)
        name_code = cursor.fetchone()
        return name_code


    # 查询药品信息数据
    @api.multi
    def search_cost_nameall(self, conn, name_id):
        cursor = conn.cursor()
        sql = "select ABF01,BBY05,bby06 from BBY1 WHERE BBY01='%s';"
        cursor.execute(sql % name_id)
        name_code = cursor.fetchone()
        return name_code

    # 查询药品费别数据
    @api.multi
    def search_cost_type(self, conn, name_id):
        cursor = conn.cursor()
        sql = "select ABF02 from ABF1  WHERE ABF01='%s';"
        cursor.execute(sql % name_id)
        name_code = cursor.fetchone()
        return name_code

