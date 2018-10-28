# -*- coding: utf-8 -*-
{
    'name': "GOODERP 银行对帐生成K3凭证",
    'author': "德清武康开源软件工作室",
    'website': "无",
    'category': 'gooderp',
    "description":
    '''
                        该模块实现银行对帐单引入，生成K3凭证。
    ''',
    'version': '11.11',
    'depends': ['core', 'finance', 'money', 'tax'],
    'data': [
        'view/bank_form_view.xml',
        'view/bank_form_action.xml',
        'view/bank_form_menu.xml',
    ],
    'demo': [
    ],
    'qweb': [
        "static/src/xml/*.xml",
    ],
}
