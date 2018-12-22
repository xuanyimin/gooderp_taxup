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
import base64
import xlrd
import xlwt

MATHING_TYPE= [
    ('ok', u'完全匹配'),
    ('handwork', u'需要手工调'),
    ('unknown', u'未匹配')]

class ExportDrawback(models.Model):
    '''海关报关单信息'''
    _name = 'export.drawback'
    _order = "name"

    name = fields.Date(u'出口日期', copy=False, required=True,)
    batch = fields.Char(u'批次',copy=False, required=True,)
    line_ids = fields.One2many('export.drawback.line', 'order_id', u'报关单明细行',
                               copy=False)
    attachment_number = fields.Integer(compute='_compute_attachment_number', string=u'附件号')

    @api.multi
    def action_get_attachment_view(self):
        res = self.env['ir.actions.act_window'].for_xml_id('base', 'action_attachment')
        res['domain'] = [('res_model', '=', 'export.drawback'), ('res_id', 'in', self.ids)]
        res['context'] = {'default_res_model': 'export.drawback', 'default_res_id': self.id}
        return res

    @api.multi
    def _compute_attachment_number(self):
        attachment_data = self.env['ir.attachment'].read_group(
            [('res_model', '=', 'export.drawback'), ('res_id', 'in', self.ids)], ['res_id'], ['res_id'])
        attachment = dict((data['res_id'], data['res_id_count']) for data in attachment_data)
        for expense in self:
            expense.attachment_number = attachment.get(expense.id, 0)

    @api.multi
    def declaration_invoice(self):
        #智能配单
        for line_id in self.line_ids:
            export_line_ids = line_id.export_declaration.line_ids
            invoice_ids = line_id.tax_invoice
            for line in export_line_ids:
                self.export_declaration(line, invoice_ids, line_id)
            for i in line_id.detailed_ids:
                if i.declaration_type != 'ok':
                    line_id.declaration_type = 'handwork'
                else:
                    line_id.declaration_type = 'ok'

    def export_declaration(self, export_line, invoice_ids, line_id):
        #配单过程
        amount = line_id.export_declaration.export_amount
        fob_amount = amount - line_id.export_declaration.bf - line_id.export_declaration.zf
        export_fob_amount = round(export_line.yb_amt/amount*fob_amount,2)
        for invoice in invoice_ids:
            type1 = self.export_unit(invoice, export_line, line_id, export_fob_amount)
            if type1:
                declaration_type = 'ok'
                return declaration_type
            type2 = self.export_unit2(invoice, export_line, line_id, export_fob_amount)
            if type2:
                declaration_type = 'ok'
                return declaration_type


    def export_unit(self, invoice,export_line,line_id,export_fob_amount):
        #数量一致且产品名称包含的
        invoice_line_ids = invoice.line_ids
        for invoice_line in invoice_line_ids:
            if invoice_line.product_name in export_line.cm_name:
                # 完全相等
                for unit_line in export_line.unit_ids:
                    invoice_line_qnt = invoice_line.product_count - invoice_line.is_used
                    if unit_line.qnt == invoice_line_qnt:
                        invoice_line.write({'is_used': invoice_line_qnt})
                        type = 'ok'
                        self.env['export.drawback.detailed'].create({
                            'line_id': line_id.id,
                            'ordinal': export_line.id,
                            'export_qnt': unit_line.qnt,
                            'export_unit': unit_line.unit.id,
                            'invoice_line': invoice_line.id,
                            'invoice_qnt': invoice_line.product_count,
                            'invoice_unit': invoice_line.product_unit,
                            'declaration_type': type,
                            'export_fob_amount': export_fob_amount,
                        })
                        return True
            else:
                raise UserError(u"未在发票找到相对应商品，请修改发票中的商品！")


    def export_unit2(self, invoice,export_line, line_id, export_fob_amount):
        # 名称包含的发票数量小于退税数量的(计量单位一至)
        invoice_line_ids = invoice.line_ids
        for invoice_line in invoice_line_ids:
            if invoice_line.product_name in export_line.cm_name:
                for unit_line in export_line.unit_ids:
                    invoice_line_qnt = invoice_line.product_count - invoice_line.is_used
                    if unit_line.unit.name == invoice_line.product_unit:
                        if unit_line.qnt > invoice_line_qnt:
                            invoice_line.write({'is_used': unit_line.qnt})
                            type = 'low'
                            return (unit_line.qnt, unit_line.unit.id,type)

    def export_unit3(self, invoice_line,export_line):
        # 名称包含的发票数量大于退税数量的(计量单位一至)
        for unit_line in export_line.unit_ids:
            invoice_line_qnt = invoice_line.product_count - invoice_line.is_used
            if unit_line.unit.name == invoice_line.product_unit:
                if unit_line.qnt < invoice_line_qnt:
                    invoice_line.write({'is_used': unit_line.qnt})
                    type = 'low'
                    return (unit_line.qnt, unit_line.unit.id,type)


    @api.multi
    def exp_data(self):
        self.exp_for_update()
        self.invoice_for_update()

    @api.multi
    def exp_for_update(self):
        xls_data = xlrd.open_workbook('./excel/10031.xls')
        Page1 = xls_data.sheet_by_name('sheet1')
        excel,colnames = self.readexcel(Page1) #读模版，返回字典及表头数组
        workbook = xlwt.Workbook(encoding = 'utf-8')   # 生成文件
        worksheet = workbook.add_sheet('sheet1')# 在文件中创建一个名为Page1的sheet

        i = j = n = 0
        for key in colnames:
            worksheet.write(0,j,key)
            j += 1
        for line in self.line_ids:
            n += 1
            if line.declaration_type != 'ok':
                raise UserError(u'有明细未完全匹配]')
            for detailed_id in line.detailed_ids:
                i += 1
                self.createexcel2(excel, detailed_id, worksheet, i, colnames, n)

        workbook.save(u'export2.xls')
        # 生成附件
        f = open('export2.xls', 'rb')
        self.env['ir.attachment'].create({
            'datas': base64.b64encode(f.read()),
            'name': u'外贸企业出口退税进货明细申报表',
            'datas_fname': u'%s外贸企业出口退税进货明细申报表.xls' % (self.name),
            'res_model': 'export.drawback',
            'res_id': self.id, })

    @api.multi
    def createexcel2(self, excel, detailed_id, worksheet, number, colnames, n):
        # 修改内容。
        export_invoice = detailed_id.line_id.export_invoice
        re_number = self.batch + '%04d' % n
        export_number = detailed_id.line_id.export_declaration.name + '%03d' % int(detailed_id.ordinal.spxh)
        export_date = detailed_id.line_id.export_declaration.lj_date
        export_product_cmcode = detailed_id.ordinal.cmcode[:8]
        export_product_cmcode_all = detailed_id.ordinal.cmcode
        export_product_name = detailed_id.ordinal.cm_name
        export_unit = self.env['export.product'].search([('code', '=', export_product_cmcode)], limit=1)
        if not export_unit:
            export_unit = self.env['export.product'].search([('code', '=', export_product_cmcode[:8])], limit=1)
        export_qnt = detailed_id.export_qnt
        yb_bz = detailed_id.ordinal.yb_bz.name
        if yb_bz == 'USD':
            usd_amount = detailed_id.export_fob_amount
            cny_amount = round(detailed_id.export_fob_amount * detailed_id.line_id.usd_rate,2)
        elif yb_bz == 'CNY':
            cny_amount = detailed_id.export_fob_amount
            usd_amount = round(detailed_id.export_fob_amount / detailed_id.line_id.usd_rate,2)
        else:
            cny_amount = round(detailed_id.export_fob_amount * detailed_id.line_id.usd_rate,2)
            usd_amount = round(cny_amount / detailed_id.line_id.usd_rate, 2)
        invoice_amount = detailed_id.invoice_line.product_amount
        export_product_rate = self.env['export.product'].search([('code', '=', export_product_cmcode)], limit=1)
        if not export_product_rate:
            export_product_rate = self.env['export.product'].search([('code', '=', export_product_cmcode[:8])], limit=1)
        invoice_tax_rate = detailed_id.invoice_line.product_tax_rate
        if export_product_rate.drawback_rate > invoice_tax_rate:
            export_tax_rate = invoice_tax_rate
        else:
            export_tax_rate = export_product_rate.drawback_rate
        invoice_tax_amount = round(detailed_id.invoice_line.product_tax * export_tax_rate / invoice_tax_rate,2)

        # invoice_tax_number = detailed_id.invoice_line.order_id.partner_code_in

        for i in excel:
            i[u'序号'] = number  # 序号
            i[u'关联号'] = re_number  # 关联号
            i[u'出口发票号'] = export_invoice  # 出口发票号
            i[u'报关单号'] = export_number  # 进货凭证号
            i[u'代理证明号'] = ''  # 代理证明号
            i[u'出口日期'] = export_date  # 发票开票日期
            i[u'核销单号'] = ''  # 核销单号
            i[u'商品代码'] = export_product_cmcode  # 商品代码
            i[u'商品名称'] = export_product_name  # 商品名称
            i[u'申报商品代码'] = export_product_cmcode_all  # 商品名称
            i[u'单位'] = export_unit.unit.name  # 单位
            i[u'出口数量'] = export_qnt  # 数量
            i[u'美元离岸价'] = usd_amount  # 美元离岸价
            i[u'人民币离岸价'] = cny_amount  # 人民币离岸价
            i[u'出口进货金额'] = invoice_amount  # 出口进货金额
            i[u'退税率'] = export_tax_rate  # 退税率
            i[u'退增值税税额'] = invoice_tax_amount  # 退增值税税额
            i[u'退消费税税额'] = 0  # 退消费税税额
            i[u'单证不齐标志'] = ''  # 单证不齐标志
            i[u'业务类型'] = ''  # 业务类型
            i[u'进料登记册号'] = ''  # 进料登记册号
            i[u'备注'] = ''  # 备注

        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(number, j, i[key])
            j += 1

    @api.multi
    def invoice_for_update(self):
        xls_data = xlrd.open_workbook('./excel/10026.xls')
        Page1 = xls_data.sheet_by_name('sheet1')
        excel,colnames = self.readexcel(Page1) #读模版，返回字典及表头数组
        workbook = xlwt.Workbook(encoding = 'utf-8')   # 生成文件
        worksheet = workbook.add_sheet('sheet1')# 在文件中创建一个名为Page1的sheet

        i = j = n = 0
        for key in colnames:
            worksheet.write(0,j,key)
            j += 1
        for line in self.line_ids:
            n += 1
            if line.declaration_type != 'ok':
                raise UserError(u'有明细未完全匹配]')
            for detailed_id in line.detailed_ids:
                i += 1
                self.createexcel(excel, detailed_id, worksheet, i, colnames, n)

        workbook.save(u'export1.xls')
        # 生成附件
        f = open('export1.xls', 'rb')
        self.env['ir.attachment'].create({
            'datas': base64.b64encode(f.read()),
            'name': u'外贸企业出口退税进货明细申报表',
            'datas_fname': u'%s外贸企业出口退税进货明细申报表.xls' % (self.name),
            'res_model': 'export.drawback',
            'res_id': self.id, })

    @api.multi
    def createexcel(self, excel, detailed_id, worksheet, number, colnames, n):
        # 修改内容。
        if detailed_id.invoice_line.order_id.invoice_type == 'zy':
            invoice_type = '增值税'
        invoice_number = detailed_id.invoice_line.order_id.invoice_code + detailed_id.invoice_line.order_id.name
        invoice_date = detailed_id.invoice_line.order_id.invoice_date
        export_product_type = detailed_id.ordinal.cmcode[:8]
        invoice_name = u'*%s*%s' % (detailed_id.invoice_line.tax_type, detailed_id.invoice_line.product_name)
        invoice_unit = detailed_id.invoice_line.product_unit
        export_number = detailed_id.invoice_qnt
        invoice_tax_rate = detailed_id.invoice_line.product_tax_rate
        # 比较大小，小的就是退税率
        export_product_cmcode = detailed_id.ordinal.cmcode
        export_product_rate = self.env['export.product'].search([('code', '=', export_product_cmcode)], limit=1)
        if not export_product_rate:
            export_product_rate = self.env['export.product'].search([('code', '=', export_product_cmcode[:8])], limit=1)
        if export_product_rate.drawback_rate > invoice_tax_rate:
            export_tax_rate = invoice_tax_rate
        else:
            export_tax_rate = export_product_rate.drawback_rate
        # todo 要是发现分拆开发票时需要改计税金额及税额
        invoice_amount = detailed_id.invoice_line.product_amount
        invoice_product_tax = detailed_id.invoice_line.product_tax
        re_number = self.batch + '%03d' % n
        if detailed_id.ordinal.yb_bz == 'CNY':
            export_type = u'跨境贸易'
        else:
            export_type = ''
        invoice_tax_number = '913300007743880298'
        invoice_tax_amount = round(detailed_id.invoice_line.product_tax * export_tax_rate / invoice_tax_rate,2)
        # invoice_tax_number = detailed_id.invoice_line.order_id.partner_code_in
        for i in excel:
            i[u'序号'] = number  # 序号
            i[u'关联号'] = re_number  # 关联号
            i[u'税种'] =  invoice_type # 税种
            i[u'进货凭证号'] =  invoice_number # 进货凭证号
            i[u'发票开票日期'] = invoice_date  # 发票开票日期
            i[u'商品代码'] = export_product_type  # 商品代码
            i[u'商品名称'] = invoice_name  # 商品名称
            i[u'单位'] = invoice_unit  # 单位
            i[u'数量'] = export_number  # 数量
            i[u'计税金额'] = invoice_amount  # 计税金额
            i[u'法定征税税率'] = invoice_tax_rate  # 法定征税税率
            i[u'税额'] = invoice_product_tax  # 税额
            i[u'退税率'] = export_tax_rate  # 退税率
            i[u'可退税额'] = invoice_tax_amount  # 可退税额
            i[u'业务类型'] = export_type  # 业务类型
            i[u'供货方税号'] = invoice_tax_number  # 供货方税号
            i[u'备注'] = ''  # 备注

        j = 0
        for key in colnames:
            # 写入excel
            worksheet.write(number, j, i[key])
            j += 1

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

class ExportDrawbackLine(models.Model):
    '''海关报关单信息'''
    _name = 'export.drawback.line'

    order_id = fields.Many2one('export.drawback', u'报关编号', index=True,
                               required=True, ondelete='cascade',
                               help=u'关联报关的编号')
    export_declaration = fields.Many2one('export.declaration', u'报关单', domain=[('export_drawback_id', '=', False)])
    tax_invoice = fields.Many2many('cn.account.invoice', string=u'进项发票', domain=[('export_drawback_id', '=', False)])
    export_invoice = fields.Char(u'出口发票',)
    declaration_type = fields.Selection(MATHING_TYPE, u'匹配状态',default='unknown')
    detailed_ids = fields.One2many('export.drawback.detailed', 'line_id', u'报关单明细行',
                               copy=False)
    original_rate = fields.Float(u'原币汇率', store=True, readonly=True,digits=(16, 5),
                        compute='_compute_export_declaration')
    usd_rate = fields.Float(u'美元汇率',store=True, readonly=True,digits=(16, 5),
                        compute='_compute_export_declaration')

    @api.one
    @api.multi
    @api.depends('export_declaration')
    def _compute_export_declaration(self):
        if self.export_declaration:
            date = self.export_declaration.lj_date
            usd = self.env['res.currency'].search([('name','=','USD')],limit=1).id
            original = self.export_declaration.currency_id.id
            self.usd_rate = self.env['money.order'].get_rate_silent(date,usd)
            self.original_rate = self.env['money.order'].get_rate_silent(date,original)

class ExportDrawbackDetailed(models.Model):
    '''海关报关单信息'''
    _name = 'export.drawback.detailed'

    line_id = fields.Many2one('export.drawback.line', u'申报明细', index=True,
                               required=True, ondelete='cascade',
                               help=u'关联申报的编号')

    ordinal = fields.Many2one('export.declaration.line', u'报关单序号')
    export_qnt = fields.Float(u'申报数量')
    export_unit = fields.Many2one('export.unit', u'单位',)
    export_fob_amount = fields.Float(u'原币离岸价')
    invoice_line = fields.Many2many('cn.account.invoice.line', string=u'发票明细行')
    invoice_qnt = fields.Float(u'申报发票数量')
    invoice_unit = fields.Char(u'申报发票单位')
    declaration_type = fields.Selection(MATHING_TYPE, u'匹配状态', default='unknown')

class cn_account_invoice(models.Model):
    _inherit = 'cn.account.invoice'
    _description = u'中国发票'
    _rec_name='name'

    export_drawback_id = fields.Many2one('export.drawback', u'出口退税申报年月', index=True, copy=False, readonly=True)

class cn_account_invoice_line(models.Model):
    _inherit = 'cn.account.invoice.line'
    _description = u'中国发票明细'
    _rec_name='product_name'

    is_used = fields.Float(u'出口退税已使用', default=0)

class ExportDeclarationLine(models.Model):
    '''海关报关单信息'''
    _inherit = 'export.declaration'
    _rec_name='name'

    export_drawback_id = fields.Many2one('export.drawback', u'出口退税申报年月', index=True, copy=False, readonly=True)
