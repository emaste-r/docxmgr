# coding=utf-8
import os

from flask import Flask

from handlers.download_handler import DownloadHandler
from handlers.index_handler import IndexHandler
from handlers.resign_director_handler import ResignDirectorHandler

app = Flask(__name__)
UPLOAD_FOLDER = 'upload'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
basedir = os.path.abspath(os.path.dirname(__file__))
ALLOWED_EXTENSIONS = set(['txt', 'png', 'jpg', 'xls', 'JPG', 'PNG', 'xlsx', 'gif', 'GIF', "docx"])

app.add_url_rule('/', view_func=IndexHandler.as_view("index"))
app.add_url_rule('/doc/resign', view_func=ResignDirectorHandler.as_view("resign_director"))
app.add_url_rule('/download/<filename>', view_func=DownloadHandler.as_view("download"))

if __name__ == '__main__':
    app.run(host="127.0.0.1", port=5000, debug=True)
