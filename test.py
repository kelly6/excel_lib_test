#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import shutil
import pickle
import datetime
from xlwt import *
from xlrd import *

fname_counter = {"01":{"确定负面":0, "疑似负面":0, "有争议":0, "疑似正面":0, "确定正面":0}, \
        "02":{"确定负面":0, "疑似负面":0, "有争议":0, "疑似正面":0, "确定正面":0},\
        "03":{"确定负面":0, "疑似负面":0, "有争议":0, "疑似正面":0, "确定正面":0},\
        "04":{"负面":0, "非负面":0}\
        }

#to_type: 01
#to_neg_type: 确定负面
def copy_file(old_path, from_fpath, new_path, to_type, to_neg_type):
    global fname_counter
    #print from_fpath
    from_fname = os.path.join(old_path, from_fpath)
    fname_counter[to_type][to_neg_type] += 1
    to_fpath = to_type + "/" + to_neg_type + "/" + str(fname_counter[to_type][to_neg_type]) + ".txt"
    to_fname = os.path.join(new_path, to_fpath)
    #print from_fname, to_fpath, new_path, to_fname
    print "from_fname:", from_fname
    print "to_fname:", to_fname
    shutil.copy(from_fname, to_fname)

def classify_article(old_path, new_path, workbook_path):
    neg_type_set = set(["01", "02", "03", "04"])
    neg_dic = {-2:"确定负面", -1:"疑似负面", 0:"有争议", 1:"疑似正面", 2:"确定正面"}
    weibo_dic = {-1:"负面", 1:"非负面"}
    neg_score_dic = {"确定负面":-2, "疑似负面":-1, "有争议":0, "疑似正面":1, "确定正面":2}
    weibo_score_dic = {"负面":-1, "非负面":1}
    
    wb_hd = open_workbook(workbook_path)
    #neg_type : 01
    #neg_score: 确定负面
    for neg_type in neg_type_set:
        score_dic = neg_score_dic
        score_desc_dic = neg_dic
        if neg_type == "04":
            score_dic = weibo_score_dic
            score_desc_dic = weibo_dic
        for neg_score in score_dic:
            table_name = neg_type + "_" + neg_score
            table = wb_hd.sheet_by_name(table_name.decode("utf8"))
            nrows = table.nrows
            for i in range(1, nrows):
                v = table.row_values(i)
                #print v[0], v[1], v[2]
                old_fpath = neg_type + "/" + neg_score + "/" + v[0].encode("utf8")
                try:
                    copy_file(old_path, old_fpath, new_path, neg_type, score_desc_dic[int(v[2])])
                except ValueError:
                    continue

def test_classify_article(old_path, new_path, workbook_path):
    neg_type_set = set(["01", "02", "03", "04"])
    neg_dic = {-2:"确定负面", -1:"疑似负面", 0:"有争议", 1:"疑似正面", 2:"确定正面"}
    weibo_dic = {-1:"负面", 1:"非负面"}
    neg_score_dic = {"确定负面":-2, "疑似负面":-1, "有争议":0, "疑似正面":1, "确定正面":2}
    weibo_score_dic = {"负面":-1, "非负面":1}
    
    v_counter = 0
    empty_counter = 0
    diff_counter = 0
    one_counter = 0
    two_counter = 0
    big_diff_counter = 0
    neg_err_counter = 0
    pos_err_counter = 0
    human_neg = 0
    human_pos = 0
    human_else = 0
    wb_hd = open_workbook(workbook_path)
    #neg_type : 01
    #neg_score: 确定负面
    for neg_type in neg_type_set:
        score_dic = neg_score_dic
        score_desc_dic = neg_dic
        if neg_type == "04":
            score_dic = weibo_score_dic
            score_desc_dic = weibo_dic
        for neg_score in score_dic:
            table_name = neg_type + "_" + neg_score
            table = wb_hd.sheet_by_name(table_name.decode("utf8"))
            nrows = table.nrows
            for i in range(1, nrows):
                v = table.row_values(i)
                #print v[0], v[1], v[2]
                old_fpath = neg_type + "/" + neg_score + "/" + v[0].encode("utf8")
                if v[0] == "":
                    continue
                try:
                    if int(v[2]) == 9:
                        v[2] = "0"
                    if neg_type == "04":
                        if int(v[2]) < 0:
                            v[2] = "-1"
                        else:
                            v[2] = "1"
                    l_v1 = int(v[1])
                    l_v2 = int(v[2])
                    if l_v2 < 0 and l_v1 >= 0:
                        neg_err_counter += 1
                    if l_v2 >= 0 and l_v1 < 0:
                        pos_err_counter += 1
                    if l_v2 < 0:
                        human_neg += 1
                    elif l_v2 >= 0:
                        human_pos += 1
                    else:
                        human_else += 1
                    if l_v2 != l_v1:
                        print "diff:", l_v1, l_v2
                        diff_counter += 1
                        if abs(l_v1 - l_v2) == 1:
                            one_counter += 1
                        if abs(l_v1 - l_v2) == 2:
                            two_counter += 1
                        if abs(l_v1 - l_v2) > 2:
                            big_diff_counter += 1
                    #copy_file(old_path, old_fpath, new_path, neg_type, score_desc_dic[int(v[2])])
                    v_counter += 1
                except ValueError:
                    #print "parse error:", v[2]
                    empty_counter += 1
                    continue
    return v_counter, empty_counter, diff_counter, one_counter, two_counter, big_diff_counter, neg_err_counter, human_neg, human_pos, human_else, pos_err_counter

#return list
#item: ("03_确定负面_1.txt", machine_result, human_result, human_name)
def gen_human_result_list(workbook_path):
    human_result_list = []
    wb_hd = open_workbook(workbook_path)
    sheet_names = wb_hd.sheet_names()
    for sheet_name in sheet_names:
        table = wb_hd.sheet_by_name(sheet_name)
        for i in range(1, table.nrows):
            row_v = table.row_values(i)
            c0 = str(sheet_name.encode("utf-8")) + "_" + str(row_v[0])
            try:
                c1 = int(row_v[1])
            except:
                print sheet_name, row_v, i
                continue
            try:
                c2 = int(row_v[2])
            except:
                continue
            try:
                c3 = str(row_v[3].encode("utf-8"))
            except:
                c3 = "default"
            if c2 == 9:
                c2 = 0
            if sheet_name.encode("utf-8").find("04_") >= 0:
                if c2 < 0:
                    c2 = -1
                else:
                    c2 = 1
            if c0.find("兼容性") >= 0:
                continue
            item = (c0, c1, c2, c3)
            print item
            human_result_list.append(item)
    return human_result_list

def gen_diff_result_list(workbook_path):
    human_result_list = []
    wb_hd = open_workbook(workbook_path)
    sheet_names = wb_hd.sheet_names()
    match_counter = 0
    pos_to_neg_counter = 0
    neg_to_pos_counter = 0
    for sheet_name in sheet_names:
        table = wb_hd.sheet_by_name(sheet_name)
        for i in range(1, table.nrows):
            row_v = table.row_values(i)
            c0 = str(sheet_name.encode("utf-8")) + "_" + str(row_v[0])
            try:
                c1 = int(row_v[1])
            except:
                print sheet_name, row_v, i
                continue
            try:
                c2 = int(row_v[2])
            except:
                continue
            c3 = str(row_v[3].encode("utf-8"))
            if c2 == 9:
                c2 = 0
            if sheet_name.encode("utf-8").find("04_") >= 0:
                if c2 < 0:
                    c2 = -1
                else:
                    c2 = 1
                if c1 < 0:
                    c1 = -1
                else:
                    c1 = 1
            if c1 == c2:
                match_counter += 1
            if c1 >= 0 and c2 < 0:
                neg_to_pos_counter += 1
            if c1 < 0 and c2 >= 0:
                pos_to_neg_counter += 1
    print "match:", match_counter
    print "pos_to_neg:", pos_to_neg_counter
    print "neg_to_pos:", neg_to_pos_counter
    return 0

if 0:
    fd = open("/home/kelly/result/human_result_list_with_res.pickle", "r")
    human_result_list_with_res = pickle.load(fd)
    fd.close()
    less_100_match_counter = 0
    between_100_165_match_counter = 0
    big_165_match_counter = 0
    idx = 0
    match_list = [0, 0, 0]
    neg_to_pos_list = [0, 0, 0]
    pos_to_neg_list = [0, 0, 0]
    match_counter = 0
    for i in human_result_list_with_res:
        if i[4] < 100:
            idx = 0
        elif i[4] >= 100 and i[4] <= 165:
            idx = 1
        else:
            idx = 2
        if i[5] == i[6]:
            match_counter += 1
            match_list[idx] += 1
        elif i[5] >= 0 and i[6] < 0:
            neg_to_pos_list[idx] += 1
        elif i[5] < 0 and i[6] >= 0:
            pos_to_neg_list[idx] += 1
    print match_counter
    print "(0, 100): match:", match_list[0], "pos_to_neg:", pos_to_neg_list[0], "neg_to_pos:", neg_to_pos_list[0]
    print "[100, 165]: match:", match_list[1], "pos_to_neg:", pos_to_neg_list[1], "neg_to_pos:", neg_to_pos_list[1]
    print "(165, max): match:", match_list[2], "pos_to_neg:", pos_to_neg_list[2], "neg_to_pos:", neg_to_pos_list[2]
    #gen_diff_result_list("simple.xls")

if 0:
    #gen_human_result_list
    human_result_list = gen_human_result_list("simple.xls")
    fd = open("/home/kelly/result/human_result_list.pickle", "w")
    pickle.dump(human_result_list, fd)
    fd.close()
    print human_result_list[len(human_result_list) - 1]
    exit()

if 0:
    article_type_list = ["01", "02", "03"]
    neg_type_dic = {"确定负面":-2, "疑似负面":-1, "有争议":0, "疑似正面":1, "确定正面":2}
    
    f = Workbook()
    #print dir(f)
    
    for tp in article_type_list:
        for neg_type in neg_type_dic:
            sheet_name = tp + "_" + neg_type
            print sheet_name.decode("utf8")
            table = f.add_sheet(sheet_name.decode("utf8"))
            table.write(0, 0, "文件名".decode("utf8"))
            table.write(0, 1, "文章得分".decode("utf8"))
            table.write(0, 2, "人工判分".decode("utf8"))
            fpath = tp + "/" + neg_type
            idx = 1
            for fl in os.listdir(fpath):
                fname = fpath + "/" + str(idx) + ".txt"
                table.write(idx, 0, Formula(('HYPERLINK("%s"; "%s")' % (fname, str(idx) + ".txt")).decode("utf8")))
                table.write(idx, 1, str(neg_type_dic[neg_type]).decode("utf8"))
                table.write(idx, 2, "1".decode("utf8"))
                idx += 1
    weibo_neg_type = "04"
    weibo_neg_type_dic = {"负面":-1, "非负面":1}
    for neg_type in weibo_neg_type_dic:
        sheet_name = weibo_neg_type + "_" + neg_type
        print sheet_name.decode("utf8")
        table = f.add_sheet(sheet_name.decode("utf8"))
        table.write(0, 0, "文件名".decode("utf8"))
        table.write(0, 1, "微博得分".decode("utf8"))
        table.write(0, 2, "人工判分".decode("utf8"))
        fpath = weibo_neg_type + "/" + neg_type
        idx = 1
        for fl in os.listdir(fpath):
            fname = fpath + "/" + str(idx) + ".txt"
            table.write(idx, 0, Formula(('HYPERLINK("%s"; "%s")' % (fname, str(idx) + ".txt")).decode("utf8")))
            table.write(idx, 1, str(weibo_neg_type_dic[neg_type]).decode("utf8"))
            table.write(idx, 2, "-1".decode("utf8"))
            idx += 1
    f.save("simple.xls")
    exit()
    
    #ws.write_merge(1, 1, 1, 10, Formula(n + '("http://www.irs.gov/pub/irs-pdf/f1000.pdf";"f1000.pdf")'), h_style)
    table = f.add_sheet("test")
    table.write(1, 1,Formula('HYPERLINK("test/testfile.txt";"testfile")'))
    #table.write(0, 0, "test")
    
    f.save("simple.xls")
    
    exit()

if 0:
    v_counter, empty_counter, diff_counter, one_counter, two_counter, big_diff_counter, neg_err_counter, human_neg, human_pos, human_else, pos_err_counter = test_classify_article("/home/kelly/temp/negative_article", "/home/kelly/temp/negative_article_human_classify", "/home/kelly/code/xlwt_test/simple.xls")
    print "counter", v_counter, "empty_counter:", empty_counter, "combin:", v_counter + empty_counter
    print "有差异:", diff_counter
    print "差距为1:", one_counter
    print "差距为2:", two_counter
    print "差距大于2:", big_diff_counter
    print "人工判负:", human_neg
    print "人工判正:", human_pos
    print "人工结果输入错误:", human_else
    print "人工判负，机器未判负:", neg_err_counter
    print "人工未判负，机器判负:", pos_err_counter

if 0:
    #wb_hd = open_workbook("/home/kelly/test_article/正负面结果统计表.xls")
    wb_hd = open_workbook("/home/kelly/code/xlwt_test/simple.xls")

    print wb_hd
    
    l = wb_hd.sheet_names()
    print "/".join(l)

    table = wb_hd.sheet_by_name(u"02_确定正面")

    print table.hyperlink_list

    print table.row_values(1)
    print table.row_values(1)[2]
    print int(table.row_values(1)[2])
    print table.cell(1, 0).value, ","
    exit()
    row = 0
    col = 0
    
    ctype = 1
    value = 'test value'
    
    xf = 0
    
    print table.cell(0, 0)
    print table.cell(0, 0).value

if 0:
    classify_article("/home/kelly/code/xlwt_test", "/home/kelly/code/xlwt_test/test", "/home/kelly/code/xlwt_test/simple.xls")

if 0:
    #f = Workbook()
    #table = f.add_sheet("test sheet".decode("utf8"))
    #table.write(0, 0, "文件名".decode("utf8"))
    #table.write(1, 0, Formula(('HYPERLINK("%s"; "%s")' % ("/home/kelly/tempfile", "tempfile")).decode("utf8")))
    #f.save("simple.xls")

    wb_hd = open_workbook("simple.xls")
    exit()

    print wb_hd

    table = wb_hd.sheet_by_index(0)

    print table.row_values(0)
    print table.row_values(1)
    print table.hyperlink_list
    link = table.hyperlink_map.get((1, 0))
    print dir(table)

    print link
    exit()

if 0:
    simhash_set = set()
    
    #返回0: 重复
    #返回1: 不重复
    def sim_filter(s):
        import simhash
        global simhash_set
        l_sh = simhash.Simhash(s.decode("utf-8"))
        for sh in simhash_set:
            dis = sh.distance(l_sh)
            if dis < 17:
                return 0
        else:
            simhash_set.add(l_sh)
            return 1

    key_list = ["url", "info_flag", "source", "siteName", "title", "content", "click_count", "comment_count", "redirect_count", "favorite_count", "ctime", "gtime", "dtime", "summary", "reien_level"]

    with open("/home/kelly/result/1000_04.pickle", "r") as fd:
        ret_list = pickle.load(fd)
    print len(ret_list)
    for i in ret_list:
        i[10] = str(datetime.datetime.strptime(i[10], "%Y%m%d%H%M%S"))
        i[11] = str(datetime.datetime.strptime(i[11], "%Y%m%d%H%M%S"))

    f = Workbook()
    table = f.add_sheet("article")
    table.write(0, 0, u"url")
    table.write(0, 1, u"info_flag")
    table.write(0, 2, u"来源")
    table.write(0, 3, u"站点名称")
    table.write(0, 4, u"标题")
    table.write(0, 5, u"内容")
    table.write(0, 6, u"点击数")
    table.write(0, 7, u"评论数")
    table.write(0, 8, u"转发数")
    table.write(0, 9, u"收藏数")
    table.write(0, 10, u"发表时间")
    table.write(0, 11, u"采集时间")
    table.write(0, 12, u"入库时间")
    table.write(0, 13, u"摘要")
    table.write(0, 14, u"正负面")

    line = 1
    neg_dic = {"3":u"非负面", "4":u"负面"}
    for i in ret_list:
        if sim_filter(i[4]):
            for idx in range(len(i)):
                if idx == 14:
                    table.write(line, idx, neg_dic[i[idx]])
                else:
                    table.write(line, idx, i[idx].decode("utf-8"))
            line += 1

    f.save("simple.xls")
    exit()
    article_type_list = ["01", "02", "03"]
    neg_type_dic = {"确定负面":-2, "疑似负面":-1, "有争议":0, "疑似正面":1, "确定正面":2}
    
    f = Workbook()
    #print dir(f)
    
    for tp in article_type_list:
        for neg_type in neg_type_dic:
            sheet_name = tp + "_" + neg_type
            print sheet_name.decode("utf8")
            table = f.add_sheet(sheet_name.decode("utf8"))
            table.write(0, 0, "文件名".decode("utf8"))
            table.write(0, 1, "文章得分".decode("utf8"))
            table.write(0, 2, "人工判分".decode("utf8"))
            fpath = tp + "/" + neg_type
            idx = 1
            for fl in os.listdir(fpath):
                fname = fpath + "/" + str(idx) + ".txt"
                table.write(idx, 0, Formula(('HYPERLINK("%s"; "%s")' % (fname, str(idx) + ".txt")).decode("utf8")))
                table.write(idx, 1, str(neg_type_dic[neg_type]).decode("utf8"))
                table.write(idx, 2, "1".decode("utf8"))
                idx += 1
    weibo_neg_type = "04"
    weibo_neg_type_dic = {"负面":-1, "非负面":1}
    for neg_type in weibo_neg_type_dic:
        sheet_name = weibo_neg_type + "_" + neg_type
        print sheet_name.decode("utf8")
        table = f.add_sheet(sheet_name.decode("utf8"))
        table.write(0, 0, "文件名".decode("utf8"))
        table.write(0, 1, "微博得分".decode("utf8"))
        table.write(0, 2, "人工判分".decode("utf8"))
        fpath = weibo_neg_type + "/" + neg_type
        idx = 1
        for fl in os.listdir(fpath):
            fname = fpath + "/" + str(idx) + ".txt"
            table.write(idx, 0, Formula(('HYPERLINK("%s"; "%s")' % (fname, str(idx) + ".txt")).decode("utf8")))
            table.write(idx, 1, str(weibo_neg_type_dic[neg_type]).decode("utf8"))
            table.write(idx, 2, "-1".decode("utf8"))
            idx += 1
    f.save("simple.xls")
    exit()
    
    #ws.write_merge(1, 1, 1, 10, Formula(n + '("http://www.irs.gov/pub/irs-pdf/f1000.pdf";"f1000.pdf")'), h_style)
    table = f.add_sheet("test")
    table.write(1, 1,Formula('HYPERLINK("test/testfile.txt";"testfile")'))
    #table.write(0, 0, "test")
    
    f.save("simple.xls")
    
    exit()

if 0:
    import datetime
    d = datetime.datetime.strptime("20140903000000", "%Y%m%d%H%M%S")
    print str(d)
