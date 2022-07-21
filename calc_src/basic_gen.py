# -*- coding: utf-8 -*-
# python自动生成basic自定义函数注册接口

py_file = r'python/gamefang.py'

TYPE_DIC = {'str':'String','int':'Integer'}
            
def gen_basic_code(s):
    '''
    生成basic函数接口
    示例语法：
    # BASIC # zjoin@str # v # sep@str@, # row_first@@1 # keep_empty@@0  
    '''
    # 清洗
    l = s.split('#')
    l = [item.strip() for item in l if item]
    # 函数名
    func_l = l[1].split('@')
    func_name = func_l[0]
    # 函数返回类型
    if len(func_l) > 1:
        return_type = func_l[1]
    else:
        return_type = None
    # 参数列表
    param_d = {}
    for item in l[2:]:
        param_l = item.split('@')
        param_name = param_l[0]
        if len(param_l) == 1:
            param_type = None
            param_default = None
        elif len(param_l) == 2:
            param_type = param_l[1]
            param_default = None
        else:
            param_type = param_l[1]
            param_default = param_l[2]
            if param_type == 'str': # 补全Basic格式字符串写法
                param_default = f'"{param_default}"'
        param_d[param_name] = [param_type,param_default]
    # 生成参数字符串
    param_str = '('
    for k,v in param_d.items():
        if not v[1] is None:    # 加可选参数
            param_str += 'Optional '
        param_str += k
        if not v[0] is None:    # 加类型
            type_name = TYPE_DIC.get(v[0])
            if type_name:
                param_str += f' As {type_name}'
        param_str += ', '
        # if v == [None,None]:
            # param_str += f'{k}, '
        # elif v[0] == None or TYPE_DIC.get(v[0]) is None:    # 未填type或type不在预定义范围内
            # param_str += f'Optional {k}, '
        # else:
            # param_str += f'Optional {k} As {TYPE_DIC[v[0]]}, '
    param_str = param_str[:-2] + ')'
    # 输出
    code = f'Function {func_name}{param_str}\n'
    for k,v in param_d.items():
        if not v[1] is None:    # 默认值
            code += f'    If IsMissing({k}) Then\n        {k} = {v[1]}\n    End If\n'
    param_names = ','.join(param_d.keys())
    code += f'    {func_name} = invokePyFunc(py_fn, "{func_name}", Array({param_names}), Array(), Array())\nEnd Function'
    return code
    
def main():
    result = ''
    with open(py_file, 'r', encoding='utf-8') as f:
        for line in f:
            if line.startswith('# BASIC'):
                result += gen_basic_code(line) + '\n\n'
    return result[:-2]
    
if __name__ == '__main__':
    result = main()
    print(result)
