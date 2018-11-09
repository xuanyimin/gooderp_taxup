# -*- coding: utf-8 -*-
{
    'name': "GOODERP 税务模块-GOODERP生成K3可导入文件",
    'author': "德清武康开源软件工作室",
    'website': "无",
    'category': 'gooderp',
    "description":
    '''
                        该模块实现在月度发票那里可以生成用于导入K3物料与K3出库发票+K3入库单。
                        本模块只读K3数据，不直接写，是生成可导入的EXCEL文件上，还是手工导入的。
                         注：需要增加k3_server = 192.168.0.89:1433
                                    k3_user = sa
                                    k3_password = 123456789　在config文件中（地址，用户跟密码需跟自己的K3SQLSERVER相对应）。
                            有更好的代码或建议联系QQ:2210864或邮箱freemanxuan@163.com
    ''',
    'version': '11.11',
    'depends': ['tax','tax_invoice_out'],
    'data': [
        'view/k3_view_picking.xml',
        'view/k3_view.xml',
        'view/good_k3_view.xml',
        'view/k3_action.xml',
        'view/k3_menu.xml',
        # 'security/ir.model.access.csv',
    ],
    'demo': [
    ],
    'qweb': [
        "static/src/xml/*.xml",
    ],
}
