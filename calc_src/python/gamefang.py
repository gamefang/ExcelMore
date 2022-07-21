# -*- coding: utf-8 -*-

def zpytype(v) -> str:
    '''
    返回python中的type名称
    @param v: 任意类型
    @return: string
    '''
    return type(v).__name__
    
def zjoin(v, sep:str = ',', row_first:bool = True, keep_empty:bool = False) -> str:
    '''
    连接一个区域各单元格的数值，返回字符串
    @param v: 任意类型
    @param sep: 分隔符
    @param row_first: 循环遍历先横向走row
    @param keep_empty: 是否保留空值
    @return: string
    '''
    if zpytype(v) == 'tuple':
        l = _tuple2list(v,row_first,keep_empty)
        return sep.join(l)
    else:
        return v

def zfetch(v:str, num:int = 1, sep:str = ','):
    '''
    从一个字符串数组中取出一个数值
    @param v: 字符串数组
    @param num: 取出第几个（从1开始，支持负数倒数）
    @param sep: 字符串数组的分隔符
    @return: string/float/int
    '''
    l = v.split(sep)
    idx = (num, num - 1)[num > 0]
    if idx == 0:return ''
    try:
        str_v = l[idx]
    except IndexError:
        return ''
    try:
        float_v = float(str_v)
    except ValueError:
        return str_v
    if str(float_v) != str_v:
        return int(float_v)
    else:
        return float_v

def zmod(v:str, val, num:int = 1, sep:str = ',') -> str:
    '''
    修改一个字符串数组中的数值
    @param v: 字符串数组
    @param val: 待修改的值
    @param num: 修改第几个（从1开始，支持负数倒数）
    @param sep: 字符串数组的分隔符
    @return: 修改后的字符串数组
    '''
    l = v.split(sep)
    idx = (num, num - 1)[num > 0]
    if idx == 0:return v
    try:
        l[idx] = str(val)
    except IndexError:
        return v
    return sep.join(l)
    
def zlen(v:str, sep:str = ',') -> int:
    '''
    获取一个字符串数组的长度
    @param sep: 字符串数组的分隔符
    @return: 长度数值
    '''
    return len(v.split(sep))
    
def zsum(v, num:int = 1, sep:str = ',', row_first:bool = True):
    '''
    指定一个字符串数组的区域以及列数，返回求和值(字符串为连接)
    @param v: 任意类型
    @param num: 求和第几列数据（从1开始，支持负数倒数）
    @param sep: 字符串数组的分隔符
    @param row_first: 循环遍历先横向走row
    @return: 数值型总和或字符串型连接
    '''
    if zpytype(v) == 'tuple':
        l = _tuple2list(v,row_first,False)  # 并不需要保留空单元格
        cur_item = zfetch(l[0],num,sep)
        if type(cur_item) in [int,float]: # 数值型
            for item in l[1:]:
                try:
                    cur_item += zfetch(item,num,sep)
                except:
                    return '#TypeError'
        else: # 字符串型
            for item in l[1:]:
                try:
                    cur_item = sep.join([cur_item, zfetch(item,num,sep)])
                except:
                    return '#TypeError'
        return cur_item
    else:
        return zfetch(v,num,sep)

def _stringfy(v) -> str:
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

def _tuple2list(v:tuple, row_first:bool = True, keep_empty:bool = False) -> list:
    '''
    将tuple类型的数值转化为list
    @param v: tuple类型(Excel Range，二维)
    @param row_first: 循环遍历先横向走row
    @param keep_empty: 是否保留空值
    @return: list，一维
    '''
    l = []
    if bool(row_first):
        for row in v:
            for item in row:
                this_item = _stringfy(item)
                if this_item == '' and not bool(keep_empty):continue
                l.append(this_item)
    else:
        for col_num in range(len(v[0])):
            for row_num in range(len(v)):
                item = v[row_num][col_num]
                this_item = _stringfy(item)
                if this_item == '' and not bool(keep_empty):continue
                l.append(this_item)
    return l

# 暂不需使用
# def _list_fill(v, fill_value = None):
#     '''
#     将二维列表/元组长度补全
#     @param v: 原始二维列表/元组
#     @param fill_value: 用于补全的数值
#     @return: 长宽相等的二维列表
#     '''
#     result = []
#     len_list = [len(row) for row in v]
#     need_len = max(len_list)
#     for row in v:
#         row = list(row)
#         if len(row) < need_len:
#             for i in range(need_len - len(row)):
#                 row.append(fill_value)
#         result.append(row)
#     return result
