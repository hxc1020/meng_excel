from collections import OrderedDict

from pyexcel_xls import get_data
from pyexcel_xls import save_data
from itertools import groupby
import time
import sys
import os

header = [
    '订单编号',
    '订单状态',
    '商品型号',
    '商品件数',
    '收货人/提货人姓名',
    '收货人/提货人手机号',
    '收货/提货详细地址',
    '下单模板信息',
    '买家留言',
    '卖家备注'
]

result_header = [
    '收货人/提货人手机号',
    '商品型号',
    '订单编号',
    '订单状态',
    '收货人/提货人姓名',
    '收货/提货详细地址',
    '下单模板信息',
    '买家留言',
    '卖家备注'
]


def __trans_item__(item, h):
    return [item[h.index(i)] for i in header]


def read_file():
    for root, dirs, files in os.walk('excel'):
        for file in files:
            print('--------开始转换：' + file)
            data = get_data('excel/' + file)
            data = data['excelReport']
            h = data[0]
            del data[0]
            filtered = [__trans_item__(item, h) for item in data if len(item) > 0]
            g = groupby(filtered, lambda x: x[header.index('收货人/提货人手机号')])
            result = [result_header]
            for key, group in g:
                tmp = [key]
                l = list(group)
                tmp.append(', '.join([i[header.index('商品型号')] + ':' + str(i[header.index('商品件数')]) for i in l]))
                first_row = l[0]
                other = filter(
                    lambda x: first_row.index(x) != header.index('商品型号') and first_row.index(x) != header.index(
                        '商品件数') and first_row.index(x) != header.index('收货人/提货人手机号'),
                    first_row)
                tmp.extend(other)

                result.append(tmp)
            result_data = OrderedDict()
            result_data.update({'sheet1': result})
            save_data('result/result__' + str(time.time()) + '_' + file, result_data)


if __name__ == '__main__':
    read_file()
