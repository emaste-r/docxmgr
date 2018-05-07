# coding=utf-8
import json
import os

from flask import Flask, render_template, request, jsonify, send_from_directory

from common.constant import EFFECT_TIME_NOW

app = Flask(__name__)
UPLOAD_FOLDER = 'upload'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
basedir = os.path.abspath(os.path.dirname(__file__))
ALLOWED_EXTENSIONS = set(['txt', 'png', 'jpg', 'xls', 'JPG', 'PNG', 'xlsx', 'gif', 'GIF', "docx"])


@app.route('/')
def index():
    return render_template('index/index.html')


@app.route("/doc/resign", methods=["POST"])
def resign():
    data_json = json.loads(request.form.to_dict().keys()[0])

    from docxtpl import DocxTemplate
    tpl = DocxTemplate('resign_tpl.docx')

    # 一些公共的参数
    effect_time = data_json['items'][0]['effect_time']
    date = data_json['items'][0]['date']
    company = data_json['items'][0]['company']
    notice_type_dic = {
        1: u'董事辞任',
        2: u'监事辞任',
        3: u'委任董事',
        4: u'委任监事',
        5: u'委任职工董事',
        6: u'委任职工监事',
        7: u'变更董事',
        8: u'变更监事',
    }
    single_person_flag = True if len(data_json['items']) == 1 else False

    # 构造标题
    title = data_json['items'][0]['job'] if single_person_flag else "Directors"

    # 构造第一段
    first_a = ""
    for item in data_json['items']:
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

    if effect_time == EFFECT_TIME_NOW:
        first_b = "The resignation takes effect immediately"
    else:
        _t = u'、'.join([job_dic[v["job"]] for v in data_json["items"]])
        first_b = u"The resignation will take effect upon the election of the new %s of the Company." % _t

    # 构造第二段
    if len(data_json['items']) == 1:
        item = data_json["items"][0]
        sex_he = 'he' if item['sex'] == 1 else 'she'
        sex_his = 'his' if item['sex'] == 1 else 'her'
        sex_Mr = 'Mr.' + item['lastname'] if item['sex'] == 1 else 'Miss.' + item['lastname']
        second_a = "%s " % (sex_Mr)
        second_b = "%s has" % (sex_his)
        third_a = "%s for %s" % (sex_Mr, sex_his)
        third_b = "%s" % (sex_his)
    else:
        tmp_str = ""
        for item in data_json['items']:
            sex_Mr = 'Mr.' + item['lastname'] if item['sex'] == 1 else 'Miss.' + item['lastname']
            tmp_str += "%s and " % sex_Mr
        tmp_str = tmp_str[0:-4]
        second_a = "Each of %s " % (tmp_str)
        second_b = "they have"
        third_a = "their"
        third_b = "their"

    context = data_json
    context['in_europe'] = True
    context['is_paid'] = False
    context['title'] = title
    context['date'] = date
    context['company'] = company
    context['first_a'] = first_a
    context['first_b'] = first_b
    context['second_a'] = second_a
    context['second_b'] = second_b
    context['third_a'] = third_a
    context['third_b'] = third_b

    tpl.render(context)
    tpl.save('upload/resign.docx')

    return jsonify({"msg": "ok", "code": 0, "url": "/download/resign.docx"})


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
