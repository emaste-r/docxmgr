# coding=utf-8
import json
import logging
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


@app.route("/doc/resign", methods=["POST"])
def resign():
    data_json = json.loads(request.data)
    print data_json

    from docxtpl import DocxTemplate
    tpl = DocxTemplate('resign_tpl.docx')

    context = data_json

    # 构造标题
    title = data_json['items'][0]['job'] if len(data_json['items']) == 1 else "Directors"

    # 构造第一段
    first_a = ""
    for item in data_json['items']:
        item = data_json["items"][0]
        sex_he = 'he' if item['sex'] == 1 else 'she'
        sex_his = 'his' if item['sex'] == 1 else 'her'
        sex_Mr = 'Mr. %s' % item['lastname'] if item['sex'] == 1 else 'Miss. %s' % item['lastname']
        sex_Mr_long = 'Mr. %s %s' % (item['lastname'], item['firstname']) if item['sex'] == 1 else 'Miss. %s%s' % (
            item['lastname'], item['firstname'])
        first_a += u"a resignation letter from %s (“%s”), informing the Board of %s resignation from the position" \
                  u" as the %s of the Company due to %s " % \
                  (sex_Mr_long, sex_Mr, sex_his, item['position'], item['reason'])
        first_a += ', and '
    first_a = first_a[0:-6]

    first_b = "The resignation takes effect immediately" if data_json['effect_time'] == 1 else \
        "he resignation will take effect upon the election of the new %s of the Company." % data_json['items'][0]['job']

    # 构造第二段
    if len(data_json['items']) == 1:
        item = data_json["items"][0]
        sex_he = 'he' if item['sex'] == 1 else 'she'
        sex_his = 'his' if item['sex'] == 1 else 'her'
        sex_Mr = 'Mr.' + item['lastname'] if item['sex'] == 1 else 'Miss.' + item['lastname']
        second_a = "%s " % (sex_Mr)
        second_b = "%s has" % (sex_his)
        third_a = "%s for %s" %(sex_Mr, sex_his)
        third_b = "%s" %(sex_his)
    else:
        tmp_str = ""
        for item in data_json['items']:
            sex_Mr = 'Mr.' + item['lastname'] if item['sex'] == 1 else 'Miss.' + item['lastname']
            tmp_str += "%s and " % sex_Mr
        tmp_str = tmp_str[0:-4]
        second_a = "each of %s " % (tmp_str)
        second_b = "they have"
        third_a = "their"
        third_b = "their"

    context['in_europe'] = True
    context['is_paid'] = False
    context['title'] = title
    context['first_a'] = first_a
    context['first_b'] = first_b
    context['second_a'] = second_a
    context['second_b'] = second_b
    context['third_a'] = third_a
    context['third_b'] = third_b

    tpl.render(context)
    tpl.save('upload/resign.docx')

    return jsonify({"msg": "ok", "code": 0, "url": "<a href='%s'>resign.docx </a>" % "/download/resign.docx"})


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
