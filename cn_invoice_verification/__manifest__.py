# -*- coding: utf-8 -*-
{
    'name': "GOODERP 税务模块-发票验证",
    'author': "德清武康开源软件工作室",
    'website': "无",
    'category': 'gooderp',
    "description":
    '''
                        该模块实现中国发票的网上手工验证，为自动化由发票生成采购订单做准备。
                        注：需要增加firefox_path = C:\Program Files\Mozilla Firefox\geckodriver.exe在config文件中。
                            并把geckodriver.exe放在base根目录，geckodriver.exe需与你的FIREFOX版本相匹配。
                        有更好的代码或建议联系QQ:2210864或邮箱freemanxuan@163.com
    ''',
    'version': '11.11',
    'depends': ['base', 'core',  'tax'],
    'data': [
        'view/cn_invoice_verification_view.xml',
        #'view/tree_view_asset.xml',
    ],
    'demo': [
    ],
    'qweb': [
    ],
}
