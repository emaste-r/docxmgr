# coding=utf-8

"""
2018年05月07日 星期一
对应董事辞任...
"""
import datetime
import json
import os

from flask import jsonify
from flask import request
from flask.views import MethodView

from common.config import DOC_DIR, DOC_TEMPLATE_DIR
from common.constant import EFFECT_TIME_NOW


class ResignDirectorHandler(MethodView):
    methods = ['GET', 'POST']

    def post(self):
        if request.data:
            data_json = json.loads(request.data)
        else:
            data_json = json.loads(request.form.to_dict().keys()[0])

        from docxtpl import DocxTemplate
        tpl = DocxTemplate('%s/resign_tpl.docx' % DOC_TEMPLATE_DIR)

        # 一些公共的参数
        effect_time = data_json['items'][0]['effect_time']
        date = data_json['items'][0]['date']
        time_format = datetime.datetime.strptime(date, '%Y-%m-%d')
        date = time_format.strftime('%d %B %Y')

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
        job_dic = {
            1: "Non-Executive Director",
            2: "Supervisory Committee",
            3: "Supervisory Committee",
            4: "Supervisory Committee",
            5: "Supervisory Committee",
            6: "Supervisory Committee",
        }
        single_person_flag = True if len(data_json['items']) == 1 else False

        # 构造标题
        title = job_dic[data_json['items'][0]['job']] if single_person_flag else "Directors"

        # 构造第一段
        first_a = ""
        for item in data_json['items']:
            sex_he = 'he' if item['sex'] == 1 else 'she'
            sex_his = 'his' if item['sex'] == 1 else 'her'
            sex_Mr = 'Mr. %s' % item['lastname'] if item['sex'] == 1 else 'Ms. %s' % item['lastname']
            sex_Mr_long = 'Mr. %s %s' % (item['lastname'], item['firstname']) if item['sex'] == 1 else 'Ms. %s%s' % (
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
            sex_Mr = 'Mr.' + item['lastname'] if item['sex'] == 1 else 'Ms.' + item['lastname']
            second_a = "%s " % (sex_Mr)
            second_b = "%s has" % (sex_he)
            third_a = "%s for %s" % (sex_Mr, sex_his)
            third_b = "%s" % (sex_his)
        else:
            tmp_str = ""
            for item in data_json['items']:
                sex_Mr = 'Mr. ' + item['lastname'] if item['sex'] == 1 else 'Ms. ' + item['lastname']
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

        if not os.path.exists(DOC_DIR):
            os.mkdir(DOC_DIR)

        # 保存文件
        now = datetime.datetime.now().strftime("%Y%m%d:%H%M%S")
        new_filename = 'resign_director%s.docx' % now
        new_filepath = '%s/resign_director%s.docx' % (DOC_DIR, now)
        tpl.save(new_filepath)

        return jsonify({"msg": "ok", "code": 0, "url": "/download/%s" % new_filename})
