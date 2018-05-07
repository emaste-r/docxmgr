# coding=utf-8

"""
2018年05月07日 星期一
首页
"""

from flask import render_template
from flask.views import MethodView


class IndexHandler(MethodView):
    methods = ['GET']

    def get(self):
        return render_template('index/index.html')
