# -*- coding: utf-8 -*-
from xltmpl.xltmpl import XlTemplate


with XlTemplate('./templates/template.xlsx', 'output.xlsx', apply_styles=True) as tmpl:
    data = [
        [1, 2],
        [1, 2]
    ]
    tmpl.append_rows(1, data, 1)
