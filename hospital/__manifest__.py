# -*- coding: utf-8 -*-
{
    'name': "GOODERP 医院模块",
    'author': "德清武康开源软件工作室",
    'website': "无",
    'category': 'gooderp',
    "description":
    '''
                        该模块为医院发票生成K3可导入文件，医院的那个软件很小众。
                        有更好的代码或建议联系QQ:2210864或邮箱freemanxuan@163.com
    ''',
    'version': '11.11',
    'depends': ['core', 'finance',],
    'data': [
        'view/hospital_view.xml',
        'view/hospital_action.xml',
        'view/hospital_menu.xml',

    ],
    'demo': [
    ],
    'qweb': [
        "static/src/xml/*.xml",
    ],
}
