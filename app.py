# coding=utf-8
import os

from flask import Flask, render_template, request, jsonify, send_from_directory

app = Flask(__name__)
UPLOAD_FOLDER = 'upload'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
basedir = os.path.abspath(os.path.dirname(__file__))
ALLOWED_EXTENSIONS = set(['txt', 'png', 'jpg', 'xls', 'JPG', 'PNG', 'xlsx', 'gif', 'GIF', "docx"])


@app.route('/')
def hello_world():
    return 'Hello World!'


@app.route("/index")
def index():
    return render_template('index.html')


@app.route("/order", methods=["POST"])
def order():
    customer_name = request.form.get("customer_name", "")
    bookname = request.form.get("bookname", "")
    quantity = int(request.form.get("quantity", 0))
    price = int(request.form.get("price", 0))
    print "customer_name=%s" % customer_name
    print "bookname=%s" % bookname
    print "quantity=%s" % quantity
    print "price=%s" % price

    from docxtpl import DocxTemplate
    tpl = DocxTemplate('order_tpl.docx')
    context = {
        'customer_name': customer_name,
        'items': [
            {'desc': bookname, 'qty': quantity, 'price': price},
        ],
        'in_europe': True,
        'is_paid': False,
        'company_name': u'无敌的公司',
        'total_price': str(price * quantity)
    }

    tpl.render(context)
    tpl.save('upload/order.docx')

    return jsonify({"msg": "ok", "code": 0, "url": "<a href='%s'>order.docx </a>" % "/download/order.docx"})


@app.route("/test")
def test():
    from docxtpl import DocxTemplate
    doc = DocxTemplate("my_word_template.docx")
    context = {'company_name': "World company"}
    doc.render(context)
    doc.save("generated_doc.docx")


@app.route("/download/<filename>")
def download(filename):
    if request.method == "GET":
        if os.path.isfile(os.path.join('upload', filename)):
            return send_from_directory('upload', filename, as_attachment=True)


if __name__ == '__main__':
    app.run(host="127.0.0.1", port=5000)
