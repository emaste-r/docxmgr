# coding=utf-8

"""
2018年05月07日 星期一
下载
"""

import os

import logging
from flask import make_response, jsonify
from flask import send_from_directory
from flask.views import MethodView

from common.config import DOC_DIR


class DownloadHandler(MethodView):
    methods = ['GET']

    def get(self, filename):
        if not os.path.isfile(os.path.join(DOC_DIR, filename)):
            logging.error("file=%s, not exist...." % filename)
            return jsonify({"msg": "file no exist...", "code": -1})

        # 需要知道2个参数, 第1个参数是本地目录的path, 第2个参数是文件名(带扩展名)
        return send_from_directory(DOC_DIR, filename, as_attachment=True)