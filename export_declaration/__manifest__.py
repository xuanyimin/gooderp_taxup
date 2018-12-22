# -*- coding: utf-8 -*-
{
    'name': "GOODERP 税务模块-报关单管理",
    'author': "德清武康开源软件工作室",
    'website': "无",
    'category': 'gooderp',
    "description":
    '''
                        该模块为税务商易企业出口退税辅助申报
    ''',
    'version': '11.11',
    'depends': ['core', 'finance', 'goods',],
    'data': [
        'view/export_declaration_view.xml',
        'wizard/import_data_view.xml',
        'view/export_declaration_action.xml',
        'view/export_declaration_menu.xml',
    ],
    'demo': [
    ],
    'qweb': [
        "static/src/xml/*.xml",
    ],
}
