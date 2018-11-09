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
import base64
import re
import odoo.addons.decimal_precision as dp
from odoo.exceptions import UserError

import sys
reload(sys)
sys.setdefaultencoding('utf8')
# 字段只读状态
READONLY_STATES = {
        'done': [('readonly', True)],
    }
#增加引出K3销售相关内容
class GoodK3ErpProducting(models.Model):
    _name = 'tax.invoice.producting'

    _order = "name"
    name = fields.Many2one(
        'finance.period',
        u'会计期间',
        ondelete='restrict',
        required=True,
        states=READONLY_STATES)
    line_ids = fields.One2many('tax.invoice.producting.line', 'order_id', u'产成品入库明细',
                               states=READONLY_STATES, copy=False)
    state = fields.Selection([('draft', u'草稿'),
                              ('done', u'已结束')], u'状态', default='draft')
    k3_sql = fields.Many2one('k3.category', u'自方公司', copy=False)
    attachment_number = fields.Integer(compute='_compute_attachment_number', string=u'附件号')

    @api.multi
    def action_get_attachment_view(self):
        res = self.env['ir.actions.act_window'].for_xml_id('base', 'action_attachment')
        res['domain'] = [('res_model', '=', 'tax.invoice.producting'), ('res_id', 'in', self.ids)]
        res['context'] = {'default_res_model': 'tax.invoice.productingt', 'default_res_id': self.id}
        return res

    @api.multi
    def _compute_attachment_number(self):
        attachment_data = self.env['ir.attachment'].read_group(
            [('res_model', '=', 'tax.invoice.producting'), ('res_id', 'in', self.ids)], ['res_id'], ['res_id'])
        attachment = dict((data['res_id'], data['res_id_count']) for data in attachment_data)
        for expense in self:
            expense.attachment_number = attachment.get(expense.id, 0)

    # COPY excel
    @api.multi
    def worksheetcopy(self,worksheet1,worksheet2):
        ncows = worksheet1.nrows
        ncols = worksheet1.ncols
        for i in range(0,ncows):
            row = worksheet1.row_values(i)
            for j in range(0,ncols):
                worksheet2.write(i,j,row[j])

    # 读取excel
    @api.multi
    def readexcel(self,table):
        ncows = table.nrows
        ncols = 0
        colnames = table.row_values(0)
        list = []
        for rownum in range(1,ncows):
            row = table.row_values(rownum)
            if row:
                app = {}
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                list.append(app)
                ncols += 1
        return list,colnames

    #插入物料
    @api.multi
    def createexcel(self, excel, line, worksheet, number, groups_name, max_code, colnames):

        for i in excel:
            # 修改内容。
            i[u'名称'] = line.product_name2 #名称
            i[u'规格型号'] = "" #规格型号
            i[u'计量单位组_FName']=i[u'基本计量单位_FGroupName']=i[u'采购计量单位_FGroupName']=i[u'销售计量单位_FGroupName']=i[u'生产计量单位_FGroupName']=i[u'库存计量单位_FGroupName'] = groups_name #单位组
            i[u'采购计量单位_FName'] = i[u'销售计量单位_FName'] = i[u'生产计量单位_FName'] = i[u'库存计量单位_FName']  =i[u'基本计量单位_FName']=line.unit #单位
            i[u'存货科目代码_FNumber'] = self.k3_sql.stock_code_out #存货科目代码
            i[u'销售收入科目代码_FNumber'] = self.k3_sql.income_code_out #销售收入科目代码
            i[u'销售成本科目代码_FNumber'] = self.k3_sql.cost_code_out #销售成本科目代码
            i[u'代码'] = max_code #物料代码

        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(number,j,i[key])
            j += 1

    # 插入PAGE1
    @api.multi
    def createproducting(self, conn, excel, worksheet, colnames, number):
        dep_code,dep_name = self.search_department(conn)
        user_code,user_name = self.search_user(conn)
        max_code = self.search_max_fbillno(conn)[0]
        t = int(re.findall("\d+", max_code)[0])
        billno = '%s%s' % ('CIN', "%06d" % (t + 1))
        for i in excel:
            # 修改内容。
            i[u'审核日期'] = i[u'日期'] = self.env['finance.period'].get_period_month_date_range(self.name)[
            1]
            i[u'编    号'] = billno
            i[u'交货单位_FNumber'] = dep_code
            i[u'交货单位_FName'] = dep_name
            i[u'制单人_FName'] = i[u'审核人_FName'] = u'宣一敏'
            i[u'验收_FNumber'] = i[u'保管_FNumber'] = user_code
            i[u'验收_FName'] = i[u'保管_FName'] = user_name
        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(number,j,i[key])
            j += 1

    # 插入PAGE2
    @api.multi
    def createproductingline(self, conn, line, excel, worksheet, colnames, number, line_number):
        unit_code = self.search_groups_name(conn, line)
        max_code = self.search_max_fbillno(conn)[0]
        t = int(re.findall("\d+", max_code)[0])
        billno = '%s%s' % ('CIN', "%06d" % (t + 1))
        wearhouse = self.search_wearhouse(conn)
        wearhouse_code, wearhouse_name = wearhouse
        for i in excel:
            # 修改内容。
            i[u'行号'] = str(line_number)
            i[u'单据号_FBillno'] = billno
            i[u'物料编码_FNumber'] = line.product_code
            i[u'物料编码_FName'] = line.product_name
            i[u'实收数量'] = i[u'基本单位实收数量'] = line.number
            i[u'金额'] = i[u'金额(本位币)'] = line.amount
            i[u'单价'] = line.price
            i[u'单位_FName'] = line.unit
            i[u'单位_FNumber'] = unit_code
            i[u'收货仓库_FNumber'] = wearhouse_code
            i[u'收货仓库_FName'] = wearhouse_name

        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(number, j, i[key])
            j += 1

    # 导出K3销售发票
    @api.multi
    def producting_order(self):
        xls_data = xlrd.open_workbook('./excel/producting.xls')
        Page1 = xls_data.sheet_by_name('Page1')
        Page2 = xls_data.sheet_by_name('Page2')
        Page3 = xls_data.sheet_by_name('Page3')
        Page4 = xls_data.sheet_by_name('t_Schema')
        conn = self.createConnection()
        excel1, colnames1 = self.readexcel(Page1)  # 读模版，返回字典及表头数组
        excel2, colnames2 = self.readexcel(Page2)
        workbook = xlwt.Workbook(encoding='utf-8')  # 生成文件
        worksheet = workbook.add_sheet(u'Page1')  # 在文件中创建一个名为Page1的sheet
        worksheet2 = workbook.add_sheet(u'Page2')
        worksheet3 = workbook.add_sheet(u'Page3')
        self.worksheetcopy(Page3, worksheet3)
        worksheet4 = workbook.add_sheet(u't_Schema')
        self.worksheetcopy(Page4, worksheet4)
        i = j = number = number2 =0
        for key in colnames1:
            worksheet.write(0,j,key)
            j += 1
        for key in colnames2:
            worksheet2.write(0,i,key)
            i += 1
        number += 1
        self.createproducting(conn, excel1, worksheet, colnames1, number)
        line_number = 0
        for line in self.line_ids:
            number2 += 1
            line_number += 1
            self.createproductingline(conn, line, excel2, worksheet2, colnames2, number2, line_number)

        workbook.save('producting.xls')
        self.closeConnection(conn)
        # 生成附件
        f = open('producting.xls', 'rb')
        self.env['ir.attachment'].create({
            'datas': base64.b64encode(f.read()),
            'name': u'k3产品入库单导入',
            'datas_fname': u'%sk3产品入库单%s.xls' % (self.k3_sql.name, self.name.name),
            'res_model': 'tax.invoice.producting',
            'res_id': self.id, })

    # 查询入库单最大编号
    @api.multi
    def search_max_fbillno(self, conn):
        cursor = conn.cursor()
        cursor.execute("select max(FBillno) from ICStockBill where ftrantype = '24';")
        FBillno = cursor.fetchone()
        return FBillno

    # 导出K3物料
    @api.multi
    def producting_product(self,order = False):
        xls_data = xlrd.open_workbook('./excel/good.xls')
        Page1 = xls_data.sheet_by_name('Page1')
        Page2 = xls_data.sheet_by_name('Page2')
        Page3 = xls_data.sheet_by_name('Page3')
        Page4 = xls_data.sheet_by_name('t_Schema')
        #连接数据库
        conn = self.createConnection()
        excel,colnames = self.readexcel(Page1) #读模版，返回字典及表头数组
        workbook = xlwt.Workbook(encoding = 'utf-8')   # 生成文件
        worksheet = workbook.add_sheet(u'Page1')# 在文件中创建一个名为Page1的sheet
        worksheet2 = workbook.add_sheet(u'Page2')
        self.worksheetcopy(Page2,worksheet2)
        worksheet3 = workbook.add_sheet(u'Page3')
        self.worksheetcopy(Page3, worksheet3)
        worksheet4 = workbook.add_sheet(u't_Schema')
        self.worksheetcopy(Page4, worksheet4)

        i = j = 0
        good = []
        values = self.k3_sql.stock_code_out
        max_code = self.search_max_code(conn,values)[0]
        for key in colnames:
            worksheet.write(0,j,key)
            j += 1
        for line in self.line_ids:
            good_id = self.search_goods(conn,line)
            if not good_id:
                if line.product_name2 in good:
                    continue
                good.append(line.product_name2)
                groups_name = self.search_groups_name(conn, line)[0]
                i += 1
                code = self.get_new_code(max_code,i)
                self.createexcel(excel, line, worksheet, i, groups_name, code, colnames)

        workbook.save(u'goods.xls')
        self.closeConnection(conn)
        # 生成附件
        f = open('goods.xls', 'rb')
        self.env['ir.attachment'].create({
            'datas': base64.b64encode(f.read()),
            'name': u'K3产成品物料导出',
            'datas_fname': u'%s物料%s.xls' % (self.k3_sql.name, self.name.name),
            'res_model': 'tax.invoice.producting',
            'res_id': self.id, })

    @api.multi
    def get_new_code(self, code, i):
        old_code = code.split('.')
        if len(old_code) == 1:
            a = old_code
            changdu = len(a)
            x = int(a) + i
            changdu2 = len(str(x))
            j = a[0:(changdu - changdu2)] + str(x)
            new_code = '%s' % (j)
        elif len(old_code) == 2:
            a, b = old_code
            changdu = len(b)
            x = int(b) + i
            changdu2 = len(str(x))
            j = b[0:(changdu - changdu2)] + str(x)
            new_code = '%s.%s' % (a, j)
        elif len(old_code) == 3:
            a, b, c = old_code
            changdu = len(c)
            x = int(c) + i
            changdu2 = len(str(x))
            j = c[0:(changdu - changdu2)] + str(x)
            new_code = '%s.%s.%s' % (a, b, j)
        return new_code

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
        conn = pymssql.connect(server=k3_server, user=k3_user, password=k3_password, database=self.k3_sql.code, charset='utf8')
        return conn

    # 关闭数据库连接。
    @api.multi
    def closeConnection(self,conn):
        conn.close()

    # 查询物料数据
    @api.multi
    def search_goods(self, conn, line):
        cursor = conn.cursor()
        sql = "select fnumber,fname,fmodel from t_ICItem WHERE fname='%s' and fmodel='%s';"
        values = (line.product_name2, "")
        cursor.execute(sql%values)
        good_id = cursor.fetchone()
        if good_id:
            return good_id
        else:
            return False

    # 查询仓库
    @api.multi
    def search_wearhouse(self, conn):
        cursor = conn.cursor()
        cursor.execute("select top 1 FNumber,Fname from t_Stock ;")
        wearhouse = cursor.fetchone()
        return wearhouse

    # 查询单位组
    @api.multi
    def search_groups_name(self, conn, line):
        cursor = conn.cursor()
        sql = "select Funitgroupid from t_MeasureUnit WHERE fname='%s';"
        values = (line.unit)
        cursor.execute(sql%values)
        groups_id = cursor.fetchone()
        if groups_id:
            cursor.execute("select fname from t_UnitGroup WHERE Funitgroupid='%s';"%(groups_id))
            groups_name = cursor.fetchone()
        else:
            raise UserError('请到K3系统增加计量单位%s。产品：%s。'% (line.unit,line.product_name))
        return groups_name

    # 查询单位CODE
    @api.multi
    def search_unit_code(self, conn, name):
        cursor = conn.cursor()
        sql = "select fnumber from t_MeasureUnit WHERE fname='%s';"
        values = (name)
        cursor.execute(sql%values)
        unit_code = cursor.fetchone()
        return unit_code

    # 查询物料最大code
    @api.multi
    def search_max_code(self, conn, values):
        cursor = conn.cursor()
        sql = "select max(fnumber) from t_ICItem WHERE FAcctID=(select faccountid from t_Account where fnumber= '%s');"
        cursor.execute(sql % values)
        max_code = cursor.fetchone()
        return max_code

    # 查询单位编码
    @api.multi
    def search_patner_code(self, conn, name):
        cursor = conn.cursor()
        sql = "select FNumber from t_Organization WHERE fname='%s';"
        values = (name)
        cursor.execute(sql % values)
        partner_code = cursor.fetchone()
        if not partner_code:
            raise UserError(u'请在K3系统中增加客户:%s'% name)
        return partner_code

    # 查询部门
    @api.multi
    def search_department(self, conn):
        cursor = conn.cursor()
        cursor.execute("select top 1 FNumber,Fname from t_Department ;")
        department = cursor.fetchone()
        return department

    # 查询员工
    @api.multi
    def search_user(self, conn):
        cursor = conn.cursor()
        cursor.execute("select top 1 FNumber,Fname from t_Emp ;")
        user = cursor.fetchone()
        return user


    @api.multi
    def button_excel(self):
        return {
            'name': u'引入excel',
            'view_mode': 'form',
            'view_type': 'form',
            'res_model': 'create.tax.producting.wizard',
            'type': 'ir.actions.act_window',
            'target': 'new',
        }

class GoodK3ErpProductingLine(models.Model):
    _name = 'tax.invoice.producting.line'

    order_id = fields.Many2one('tax.invoice.producting', u'生产单', index=True, copy=False, readonly=True)
    product_code = fields.Char(u'产品代码')
    product_name = fields.Char(u'产品名称')
    product_name2 = fields.Char(u'规格型号')
    unit= fields.Char(u'单位')
    amount = fields.Float(u'金额')
    number = fields.Float(u'数量')
    price = fields.Float(u'单价')
    material = fields.Float(u'直接材料')
    artificial = fields.Float(u'直接人工')
    manufacturing = fields.Float(u'制造费用')


#导入金穗发票，生成产成品明细
class create_tax_producting_wizard(models.TransientModel):
    _name = 'create.tax.producting.wizard'
    _description = 'Sale Invoice Import'

    excel = fields.Binary(u'导入excel文件',)

    @api.one
    def create_producting(self):
        """
        通过Excel文件导入信息到tax.invoice
        """
        producting = self.env['tax.invoice.producting'].browse(self.env.context.get('active_id'))
        if not producting:
            return {}
        xls_data = xlrd.open_workbook(
                file_contents=base64.decodestring(self.excel))
        table = xls_data.sheets()[0]
        #取得行数
        ncows = table.nrows
        #取得第3行数据
        colnames =  table.row_values(2)
        list =[]
        newcows = 0
        for rownum in range(3,ncows):
            row = table.row_values(rownum)
            if row:
                app = {}
                for i in range(len(colnames)):
                   app[colnames[i]] = row[i]
                #过滤掉不需要的行，详见销货清单的会在清单中再次导入
                if app.get(u'单价') :
                    list.append(app)
                    newcows += 1
        #数据读入。
        invoice_id = False
        for data in range(0,newcows):
            in_xls_data = list[data]
            product_code = in_xls_data.get(u'产品代码')
            product_name = in_xls_data.get(u'产品名称')
            product_name2 = in_xls_data.get(u'规格型号')
            unit = in_xls_data.get(u'单位')
            amount = in_xls_data.get(u'合计')
            number = in_xls_data.get(u'数量')
            price = in_xls_data.get(u'单价')
            material = in_xls_data.get(u'直接材料')
            artificial = in_xls_data.get(u'直接人工')
            manufacturing = in_xls_data.get(u'制造费用')
            #创建销售发票,已有的发票跳过
            if in_xls_data.get(u'单价'):
                self.env['tax.invoice.producting.line'].create({
                    'product_code': product_code,
                    'product_name': product_name,
                    'product_name2': product_name2,
                    'unit':unit,
                    'amount': amount,
                    'number': number,
                    'price': price,
                    'material': material,
                    'artificial': artificial,
                    'manufacturing': manufacturing,
                    'order_id': producting.id,
                })


    def excel_date(self,data):
        #将excel日期改为正常日期
        if type(data) in (int,float):
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(data,0)
            py_date = datetime.datetime(year, month, day, hour, minute, second)
        else:
            py_date = data
        return py_date
