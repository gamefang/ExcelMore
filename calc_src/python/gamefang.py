# -*- coding: utf-8 -*-

def zpytype(v):
    '''
    返回python中的type名称
    @param v: 任意类型
    @return: string
    '''
    return type(v).__name__
    
def zjoin(v, sep=',', row_first=True, keep_empty=False):
    '''
    连接一个区域各单元格的数值，返回字符串
    @param v: 任意类型
    @param sep: 分隔符
    @param row_first: 循环遍历先横向走row
    @param keep_empty: 是否保留空值
    @return: string
    '''
    if zpytype(v) == 'tuple':
        l = []
        if bool(row_first):
            for row in v:
                for item in row:
                    this_item = _stringfy(item)
                    if this_item == '' and not bool(keep_empty):continue
                    l.append(this_item)
        else:
            for col_num in range(len(v)):
                for row_num in range(len(v[0])):
                    item = v[row_num][col_num]
                    this_item = _stringfy(item)
                    if this_item == '' and not bool(keep_empty):continue
                    l.append(this_item)
        return sep.join(l)
    else:
        return v

def _stringfy(v):
    '''
    将内容字符串化
    '''
    t = zpytype(v)
    if t == 'float':
        if int(v) == v:
            v = int(v)
    elif t == 'NoneType':
        v = ''
    return str(v)