# -*- coding: utf-8 -*-
from xltmpl.xltmpl import XlTemplate

with XlTemplate('./templates/template.xls', 'output.xls', apply_styles=True) as tmpl:
    data = {'name': '蔡文姬', 'role': '辅助'}
    tmpl.append_dict(1, data, header_row=1, style_row=2)
