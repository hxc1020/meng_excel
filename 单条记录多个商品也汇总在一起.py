from collections import OrderedDict

from pyexcel_xls import get_data
from pyexcel_xls import save_data
import time
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
    '卖家备注',
    '地址合并'
]


def __trans_item__(item, h):
    return [item[h.index(i)] for i in header]


def flatten(target):
    new_list = []
    for x in target:
        for xx in x:
            new_list.append(xx)
    return new_list


def group_by(items, fn):
    result = {}
    for item in items:
        key = fn(item)
        if key in result:
            result[key].append(item)
        else:
            result[key] = [item]
    return result


def read_file():
    for root, dirs, files in os.walk('excel'):
        for file in files:
            print('--------开始转换：' + file)
            data = get_data('excel/' + file)
            data = data['excelReport']
            h = data[0]
            del data[0]
            all_products = {}
            filtered = [__trans_item__(item, h) for item in data if len(item) > 0]
            g = group_by(filtered, lambda x: x[header.index('收货人/提货人手机号')])
            result = [result_header]
            for key in g:
                tmp = [key]
                l = g[key]
                products = flatten([i[header.index('商品型号')].split(';') for i in l])
                product_count = flatten([i[header.index('商品件数')].split(';') for i in l])
                addresses = flatten([i[header.index('收货/提货详细地址')].split(';') for i in l])
                for i, x in enumerate(products):
                    if x in all_products:
                        all_products[x] += int(product_count[i])
                    else:
                        all_products[x] = int(product_count[i])
                    products[i] = x + ':' + str(product_count[i])
                tmp.append(', '.join(products))
                first_row = l[0]
                other = filter(
                    lambda x: first_row.index(x) != header.index('商品型号') and first_row.index(x) != header.index(
                        '商品件数') and first_row.index(x) != header.index('收货人/提货人手机号'),
                    first_row)
                tmp.extend(other)
                tmp.append(', '.join(set(addresses)))
                result.append(tmp)

            for k in all_products:
                result.append([k, all_products[k]])

            result_data = OrderedDict()
            result_data.update({'sheet1': result})
            save_data('result/result__' + str(time.time()) + '_' + file, result_data)


if __name__ == '__main__':
    read_file()
