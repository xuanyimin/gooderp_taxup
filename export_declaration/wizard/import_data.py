# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models
import base64
from lxml import etree
import xlrd
import codecs

SELECT = [
    ('excel', u'互联网中的EXCEL'),
    ('xml', u'已解密的报关单xml'),]

class ImportExportDeclaration(models.TransientModel):
    _name = 'import.export.declaration'
    _description = u'导入已解密的报关单xml'

    type = fields.Selection(SELECT, u'类型',)
    ht_xml = fields.Binary(
        "已解密的报关单xml文件", attachment=True, required=True)

    @api.multi
    def import_data(self):
        '''
        读入报关单xml文件，生成报关单信息
        '''

        file = base64.b64decode(self.ht_xml).decode('gb2312').encode('utf-8')
        parser = etree.HTMLParser(strip_cdata = False)
        tree = etree.HTML(file, parser=parser)
        for bbox in tree.xpath('//dec'):
            for corner in bbox.getchildren():
                #跳过重复的报关单
                try:
                    if corner.tag == 'dechead':
                        export_id = self.create_export_declaration(corner)
                    if corner.tag == 'declists':
                        for corner2 in corner.getchildren():
                            self.create_export_declaration_line(corner2,export_id)
                except:
                    continue

    @api.multi
    def create_export_declaration(self,corner):
        app = {}
        list = []
        for export in corner.getchildren():
            app[export.tag] = export.text
        if app and app.get('bgd_no'):
            list.append(app)

        for d in list:
            name = d.get('bgd_no')
            lj_date = d.get('lj_date')
            cj_type = d.get('cj_type')
            yf = d.get('yf')
            bf = d.get('bf')
            my_type = d.get('my_type')
            old_name = self.env['export.declaration'].search([('name','=',name)],limit=1)
            if old_name:
                continue
            else:
                export_declaration_id = self.env['export.declaration'].create({
                    'name':name,
                    'lj_date':lj_date,
                    'cj_type':cj_type,
                    'yf':yf or 0,
                    'bf':bf or 0,
                    'my_type':my_type,
                })
            return export_declaration_id

    @api.multi
    def create_export_declaration_line(self, corner,export_id):
        app = {}
        list = []
        for export in corner.getchildren():
            app[export.tag] = export.text
        if app and app.get('cmcode'):
            list.append(app)
        for d in list:
            spxh = d.get('spxh')
            cmcode = d.get('cmcode')
            cm_name = d.get('cm_name')
            yb_bz = d.get('yb_bz')
            currency = self.env['res.currency'].search([('name','=',yb_bz)],limit=1).id
            yb_amt = d.get('yb_amt')
            fd_unit = d.get('fd_unit')
            fd_qnt = d.get('fd_qnt')
            no2_fd_unit = d.get('no2_fd_unit')
            no2_fd_qnt = d.get('no2_fd_qnt')
            cj_unit = d.get('cj_unit')
            cj_qnt = d.get('cj_qnt')

            line_id = self.env['export.declaration.line'].create({
                'order_id':export_id.id,
                'spxh':spxh,
                'cmcode':cmcode,
                'cm_name':cm_name,
                'yb_bz':currency,
                'yb_amt':float(yb_amt),
            })
            print currency
            export_id.write({'currency_id':currency})

            if fd_unit:
                unit = self.env['export.unit'].search([('code','=',fd_unit)])
                self.env['export.declaration.line.unit'].create({
                    'unit':unit,
                    'qnt': float(fd_qnt),
                    'line_id':line_id.id,
                    'note':'fd_unit',
                })
            if no2_fd_unit:
                unit = self.env['export.unit'].search([('code', '=', no2_fd_unit)])
                self.env['export.declaration.line.unit'].create({
                    'unit':unit,
                    'qnt': float(no2_fd_qnt),
                    'line_id':line_id.id,
                    'note':'no2_fd_unit',
                })
            if cj_unit:
                unit = self.env['export.unit'].search([('code', '=', cj_unit)])
                self.env['export.declaration.line.unit'].create({
                    'unit': unit,
                    'qnt': float(cj_qnt),
                    'line_id': line_id.id,
                    'note': 'cj_unit',
                })

    @api.multi
    def import_data2(self):
        '''
        读入报关单excel文件，生成报关单信息
        '''
        xls_data = xlrd.open_workbook(
            file_contents=base64.decodestring(self.ht_xml))
        table = xls_data.sheets()[0]
        # 取得行数
        ncows = table.nrows
        # 取得第1行数据

        colnames = table.row_values(0)
        list = []
        newcows = 0
        for rownum in range(2, ncows):
            row = table.row_values(rownum)
            if row:
                app = {}
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                # 过滤掉不需要的行，详见销货清单的会在清单中再次导入
                if app.get(u'商品代码') :
                    list.append(app)
                    newcows += 1
        for data in range(0, newcows):
            in_xls_data = list[data]
            export_declaration_id = self.env['export.declaration'].search([('name','=',in_xls_data.get(u'海关报关单号'))],limit=1)
            name = in_xls_data.get(u'海关报关单号')
            lj_date = in_xls_data.get(u'出口日期')
            cj_type = in_xls_data.get(u'成交方式')
            yf = in_xls_data.get(u'保费金额')
            bf = in_xls_data.get(u'运费金额')
            my_type = in_xls_data.get(u'海关贸易方式代码')
            currency = in_xls_data.get(u'币种').split(" ")[0]
            currency_id  = self.env['res.currency'].search([('name', '=', currency)], limit=1).id
            spxh = int(in_xls_data.get(u'序号'))
            cmcode = in_xls_data.get(u'商品代码')
            cm_name = in_xls_data.get(u'商品名称')
            yb_amt = in_xls_data.get(u'成交金额')
            fd_unit= in_xls_data.get(u'计量单位1')
            fd_qnt = in_xls_data.get(u'数量1')
            no2_fd_unit = in_xls_data.get(u'计量单位2')
            no2_fd_qnt = in_xls_data.get(u'数量2')
            cj_unit = in_xls_data.get(u'计量单位3')
            cj_qnt = in_xls_data.get(u'数量3')
            ht_no = in_xls_data.get(u'进出口合同号')
            if not export_declaration_id:
                export_amount = 0
                export_declaration_id = self.env['export.declaration'].create({
                    'name': name,
                    'lj_date': lj_date,
                    'cj_type': cj_type,
                    'yf': yf or 0,
                    'bf': bf or 0,
                    'my_type': my_type,
                    'currency_id':currency_id,
                    'ht_no':ht_no,
                    'export_amount':export_amount
                })

            export_declaration_line = self.env['export.declaration.line'].search(['&',('spxh','=',spxh),('order_id','=',export_declaration_id.id)],limit=1)
            if not export_declaration_line:
                line_id = self.env['export.declaration.line'].create({
                    'order_id': export_declaration_id.id,
                    'spxh': spxh,
                    'cmcode': cmcode,
                    'cm_name': cm_name,
                    'yb_bz': currency_id,
                    'yb_amt': yb_amt,
                })
                export_amount += yb_amt
                export_declaration_id.write({'export_amount':export_amount})

                if fd_unit:
                    unit = self.env['export.unit'].search([('name', '=', fd_unit)],limit=1)
                    self.env['export.declaration.line.unit'].create({
                        'unit': unit.id,
                        'qnt': float(fd_qnt),
                        'line_id': line_id.id,
                        'note': 'fd_unit',
                    })
                else:
                    raise UserWarning(u'请将表头的计量单位改为计量单位1，数量改为数量1')
                if no2_fd_unit:
                    unit = self.env['export.unit'].search([('name', '=', no2_fd_unit)],limit=1)
                    self.env['export.declaration.line.unit'].create({
                        'unit': unit.id,
                        'qnt': float(no2_fd_qnt),
                        'line_id': line_id.id,
                        'note': 'no2_fd_unit',
                    })
                if cj_unit:
                    unit = self.env['export.unit'].search([('name', '=', cj_unit)],limit=1)
                    self.env['export.declaration.line.unit'].create({
                        'unit': unit.id,
                        'qnt': float(cj_qnt),
                        'line_id': line_id.id,
                        'note': 'cj_unit',
                    })