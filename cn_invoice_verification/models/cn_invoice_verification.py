# -*- coding: utf-8 -*-
##############################################################################
#
#    Copyright (C) 2016  德清武康开源软件().
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
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import base64
from PIL import Image,ImageDraw,ImageFont
import requests
from hashlib import md5

driver = False

#定意发票todo add confirm_month
class cn_account_invoice(models.Model):
    _inherit = 'cn.account.invoice'
    _description = u'中国发票'
    _rec_name='name'

    img = fields.Binary(u'图片', attachment=True)
    img_note = fields.Char(u"填写要求")
    img_code = fields.Char(u'验证码')

    def find_firefox_path(self):
        path = None
        if config['firefox_path'] and config['firefox_path'] != 'None':
            path = config['firefox_path']
        try:
            return path
        except IOError:
            raise Exception('Command `%s` not found.' % path)

    # 取验证码
    def get_img_code(self):
        global driver
        if driver:
            driver.close()
        firefox_path = self.find_firefox_path()
        driver = webdriver.Firefox(executable_path=firefox_path)
        driver.maximize_window()
        actions = ActionChains(driver)
        url = "https://inv-veri.chinatax.gov.cn/index.html"
        driver.get(url) #打开网页
        WebDriverWait(driver, 5, 0.5).until(ec.presence_of_all_elements_located((By.ID, "kprq")))
        fpdm = driver.find_element_by_id("fpdm")
        fpdm.send_keys(self.invoice_code)
        time.sleep(1)
        actions.click(fpdm)#提交发票代码
        fphm = driver.find_element_by_id("fphm")
        fphm.send_keys(self.name) #提交发票号码
        date = self.invoice_date.replace('-', '')  # 转换日期格式
        date_in = driver.find_element_by_id("kprq")#提交发票日期
        date_in.clear()
        date_in.send_keys(date)
        if self.invoice_type == 'zy':
            driver.find_element_by_id("kjje").send_keys(str(self.invoice_amount))#提交发票金额
        if self.invoice_type == 'pt':
            driver.find_element_by_id("kjje").send_keys(str(self.invoice_heck_code)[-6:])  # 提交校验码后六位
        if self.invoice_type == 'dz':
            driver.find_element_by_id("kjje").send_keys(str(self.invoice_heck_code)[-6:])  # 提交校验码后六位
        time.sleep(1)
        WebDriverWait(driver, 5, 0.5).until(ec.presence_of_all_elements_located((By.ID, "yzminfo")))
        # 取图片
        try:
            imgelement = driver.find_element_by_id("yzm_img")
        except:
            actions.click(fpdm)
            actions.click(fphm)
            imgelement = driver.find_element_by_id("yzm_img")
        finally:
            img = imgelement.get_attribute('src').split(",")[-1]
        # 取图片说明
        img_note = driver.find_element_by_id("yzminfo").text
        self.write({'img': img, 'img_note': img_note})
        # 处理base64代码不足报错
        missing_padding = 4 - len(img) % 4
        if missing_padding:
            img += b'=' * missing_padding
        imgdata = base64.b64decode(img)
        file = open('./log/out.png', 'wb')
        file.write(imgdata)

    # 传验证码取得明细
    def to_verified(self):
        global driver
        if not driver:
            raise UserError(u'提示！请先取验证码')
        #先前有报错则点击确认
        try:
            popup_ok_button = driver.find_element_by_id("popup_ok")
            popup_ok_button.click()
        except:
            pass

        yzm = driver.find_element_by_id("yzm")
        yzm.clear()
        yzm.send_keys(self.img_code)
        button = driver.find_element_by_id("checkfp")
        button.click()
        time.sleep(6)
        # 如果有JS弹出窗，则由gooderp返回给用户
        if self.invoice_type == 'zy':
            dm = 'zp'
        if self.invoice_type == 'pt':
            dm = 'pp'
        if self.invoice_type == 'dz':
            dm = 'dzfp'
        try:
            WebDriverWait(driver, 5, 0.5).until(ec.presence_of_all_elements_located((By.ID, 'se_%s'%dm)))
        except :
            text = driver.find_element_by_id("popup_message").text or driver.find_element_by_id("cyjg").text
            raise UserError(u'提示！%s' % text)
            return
        else:
            if self.line_ids:
                self.line_ids.unlink() # 删除旧数据
            self.get_invioce_mx()  # 取得明细

    # 传验证码取得明细
    def get_invioce_mx(self):
        global driver
        if self.invoice_type == 'zy':
            dm = 'zp'
        if self.invoice_type == 'pt':
            dm = 'pp'
        if self.invoice_type == 'dz':
            dm = 'dzfp'
        se = 'se_%s'%dm #定位总金额
        je = 'je_%s'%dm #定位总税额
        self.invoice_tax = float(driver.find_element_by_id(se).text[1:])
        if self.invoice_type != 'zy':
            self.invoice_amount = float(driver.find_element_by_id(je).text[1:])
        mc_in = 'xfmc_%s'%dm #定位名称
        sbh_in = 'xfsbh_%s'%dm #定位税号
        dzdh_in = 'xfdzdh_%s'%dm #定位地址电话
        yhzh_in = 'xfyhzh_%s'%dm #定位银行及帐号
        self.partner_name_in = driver.find_element_by_id(mc_in).text
        self.partner_code_in = driver.find_element_by_id(sbh_in).text
        self.partner_address_in = driver.find_element_by_id(dzdh_in).text
        self.partner_bank_number_in = driver.find_element_by_id(yhzh_in).text
        mc_out = 'gfmc_%s'%dm #定位名称
        sbh_out = 'gfsbh_%s'%dm #定位税号
        dzdh_out = 'gfdzdh_%s'%dm #定位地址电话
        yhzh_out = 'gfyhzh_%s'%dm #定位银行及帐号
        self.partner_name_out = driver.find_element_by_id(mc_out).text
        self.partner_code_out = driver.find_element_by_id(sbh_out).text
        self.partner_address_out = driver.find_element_by_id(dzdh_out).text
        self.partner_bank_number_out = driver.find_element_by_id(yhzh_out).text

        bz= 'bz_%s'%dm#定位备注
        self.note = driver.find_element_by_id(bz).text
        #定位明细行
        tab_head = '//tr [@id="tab_head_%s"]/../tr'% dm
        try:
            '''
            todo 红字跟这个有关 style="position:absolute;top:0px;left:0px;display:none;"
            zf = driver.find_element_by_id("icon_zf")
            if zf:
                self.write({'is_deductible': 1, 'is_verified': 1})
                return
            '''
            # 有清单
            button_mx = driver.find_element_by_id("showmx")
            button_mx.click()
            time.sleep(4)
            tr_line = driver.find_elements_by_xpath("//tr [@id='tab_head_mx']/../tr")
            for tr in tr_line[1:-3]:
                # 将每一个tr的数据根据td查询出来，返回结果为list对象,第一行和最后三行不要
                table_td_list = tr.find_elements_by_tag_name("td")
                row_list = []
                for td in table_td_list:  # 遍历每一个td
                    row_list.append(td.text)  # 取出表格的数据，并放入行列表里
                have_type = row_list[1].split('*')
                if len(have_type) > 1:
                    goods_name = row_list[1].split('*')[-1]
                    tax_type = row_list[1].split('*')[1]
                else:
                    goods_name = row_list[1]
                    tax_type = ''
                self.env['cn.account.invoice.line'].create({
                    'order_id': self.id,
                    'product_name': goods_name or '',
                    'product_type': row_list[2] or '',
                    'product_unit': row_list[-6] or '',
                    'product_count': row_list[-5] or '',
                    'product_price': row_list[-4] or '',
                    'product_amount': row_list[-3] or '0',
                    'product_tax_rate': row_list[-2].replace('%', '') or '0',
                    'product_tax': row_list[-1] or '0',
                    'tax_type': tax_type,
                })


        except:
            # 无清单
            tr_line = driver.find_elements_by_xpath(tab_head)
            for tr in tr_line[1:-2]:
                # 将每一个tr的数据根据td查询出来，返回结果为list对象,第一行和最后二行不要
                table_td_list = tr.find_elements_by_tag_name("td")
                row_list = []
                for td in table_td_list:  # 遍历每一个td
                    row_list.append(td.text)  # 取出表格的数据，并放入行列表里
                have_type = row_list[0].split('*')
                if len(have_type) > 1:
                    goods_name = row_list[0].split('*')[-1]
                    tax_type = row_list[0].split('*')[1]
                else:
                    goods_name = row_list[0]
                    tax_type = ''
                self.env['cn.account.invoice.line'].create({
                    'order_id': self.id,
                    'product_name': goods_name or '',
                    'product_type': row_list[1] or '',
                    'product_unit': row_list[-6] or '',
                    'product_count': row_list[-5] or '',
                    'product_price': row_list[-4] or '',
                    'product_amount': row_list[-3] or '0',
                    'product_tax_rate': row_list[-2].replace('%', ''),
                    'product_tax': row_list[-1] or '0',
                    'tax_type': tax_type,
                })
        driver.get_screenshot_as_file("./log/fapiao.png")
        self.write({'is_verified': 1})
        driver.close()
        driver = False
        # 上传附件
        f = open('./log/fapiao.png', 'rb')
        self.env['ir.attachment'].create({
            'datas': base64.b64encode(f.read()),
            'name': u'发票',
            'datas_fname': u'fapiao.png' ,
            'res_model': 'cn.account.invoice',
            'res_id': self.id, })

    def get_code(self):
        #处理图片
        img = Image.open('./log/out.png')
        width = img.size[0]
        height = img.size[1]
        im = Image.new("RGB", (width * 4, height), color=(255, 255, 255))
        box = (0, 0, width, height)
        im.paste(img, box)
        ttfont = ImageFont.truetype("simhei.ttf", 15)
        draw = ImageDraw.Draw(im)
        note = self.img_note
        draw.text((width + 10, 10), note, fill=(0, 0, 0), font=ttfont)
        im.save('./log/out_img.png')
        #上传图片
        username = self.env['ir.values'].get_default('tax.config.settings', 'default_dmpt_name')
        password = self.env['ir.values'].get_default('tax.config.settings', 'default_dmpt_password')
        soft_id = '99787'
        soft_key = '6cff6c2ab82049bb9050e31a5fe2b115'
        base_param = {
            'username': username,
            'password': password,
            'softid': soft_id,
            'softkey': soft_key,
        }
        header = {
            'Connection': 'Keep-Alive',
            'Expect': '100-continue',
            'User-Agent': 'ben',
        }
        im = open('./log/out_img.png', 'rb').read()
        """
                typeid: 难度
                timeout： 超时时间
                im: 图片字节
                im_type: 题目类型
                """
        params = {
            'typeid': '8014',
            'timeout': '60',
        }
        params.update(base_param)
        files = {'image': ('out_img.png', im)}
        r = requests.post('http://api.ruokuai.com/create.json', data=params, files=files, headers=header)
        return r.json()

    @api.one
    @api.multi
    def auto_verification(self):
        self.get_img_code()
        img_code = self.get_code()
        self.img_code = img_code.get(u'Result')
        self.to_verified()

