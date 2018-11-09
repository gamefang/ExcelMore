Attribute VB_Name = "原创"
'############################################################################################test
'2017/8/1，zc，仅用于打开图标文件
Sub FileOpen()
    If ActiveCell.Value Like "*.*" Then    '打开图标文件
        Const targetpath = "src\fbclient\resfile\ui\images\" '需要打开文件的目录，上级目录project
        Dim fulpath As String, filename As String
        fullpath = Split(ActiveWorkbook.Path, "design") '以excel所在design目录确定项目路径
        filename = fullpath(0) & targetpath & ActiveCell.Value
        Shell "cmd.exe /c" & filename
    End If
End Sub

'############################################################################################
'2017/8/4，zc，用于配置html语句输出。
'每句使用"|"分隔，强调处使用"{"与"}"包含。
'源字符串内容不可包含"|"、"{"、"}"。
'参数1  str 源字符串
'参数2  before  每句前内容，通常表示加点，默认加"<img src='img://Uiicon_zhuangbeitip_dian.png'>"
'参数3  accent_style    强调内容字体形式，格式"颜色+字号"，可省略字号，默认"#ffcc33+12"
'参数4  normal_style    普通内容字体形式，格式"颜色+字号"，可省略字号，默认"#e5d2ac"
Function html( _
    str As String, _
    Optional before As String = "<img src='img://Uiicon_zhuangbeitip_dian.png'>", _
    Optional accent_style As String = "#ffcc33+12", _
    Optional normal_style As String = "#e5d2ac") As String
    
    Dim pre_nor, suf_nor, pre_acc, suf_acc As String
    Dim content
    Dim i As Byte, line As String
    
    pre_nor = get_pre(normal_style) '得到普通字体html
    suf_nor = "</font>"
    pre_acc = get_pre(accent_style) '得到强调字体html
    suf_acc = "</font>"
    
    content = Split(str, "|")   '字符串组
    For i = 0 To UBound(content)
        line = content(i)   '取字符串组项
        line = Replace(line, "{", pre_acc)  '替换强调前缀
        line = Replace(line, "}", suf_acc)  '替换强调后缀
        line = before & pre_nor & line & suf_nor & "<br>" '连接前缀（点图片）及换行，并扩起普通字体
        html = html & line  '更新返回值
    Next
    
End Function

'由html函数调用，计算html输出，分隔符为"+"。
Private Function get_pre(str As String) As String
    Dim tmp
    tmp = Split(str, "+")
    If UBound(tmp) = 0 Then '仅指定颜色（1个元素）
        get_pre = "<font color='" & tmp(0) & "'>"
    ElseIf UBound(tmp) = 1 Then '指定了颜色和字体大小
        get_pre = "<font color='" & tmp(0) & "' size='" & tmp(1) & "'>"
    Else
        MsgBox UBound(tmp) & "配置错误！"
    End If
End Function

'2017/8/4，zc，html反函数，用于将html语句还原回可视化语句。
'配置pre_acc可指定强调前缀，必须为完整<>内容。强调结束符会自动补齐。
'配置endl可指定换行符，一般为<br>。
Function h2s(str As String, Optional pre_acc As String = "<font color='#ffcc33' size='12'>", Optional endl As String = "<br>") As String

    Dim pass As Boolean, accent As Boolean, i As Integer, c As String
    
    str = Replace(str, pre_acc, "{")    '替换强调前缀
    str = Replace(str, endl, "|")   '分隔行
    
    pass = False
    accent = False
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If c = "<" Then '开启跳过模式
            pass = True
            If accent Then  '正在强调，遇见第一个"<"则结束强调，补充强调关闭符
                h2s = h2s & "}"
                accent = False
            End If
        ElseIf c = ">" Then '关闭跳过模式
            pass = False
        ElseIf Not pass Then    '输出字符
            h2s = h2s & c
            If c = "{" Then accent = True   '打开强调开关
        End If
    Next
    
    If Right(h2s, 1) = "|" Then h2s = Left(h2s, Len(h2s) - 1) '去掉最后一位的分隔符
    
End Function

'############################################################################################
'2017/9/6，将选定区域合并为字符串，2018/10/12可反向连接
'@param sel: range区域
'@param sign: 连接符，默认,
'@param reverse: 是否反向连接，默认False
'@return: 连接后的字符串
Function join(sel As range, Optional ByVal sign As String = ",", Optional ByVal reverse As Boolean = False) As String
    For Each i In sel
        If i <> "" Then  '2018/8/22 添加空单元格不计算的逻辑
            If reverse Then
                join = sign & i & join
            Else
                join = join & i & sign
            End If
        End If
    Next
    If reverse Then
        join = Right(join, Len(join) - Len(sign))
    Else
        join = Left(join, Len(join) - Len(sign))
    End If
End Function

'############################################################################################
'2017/8/5，zc，分离数据的函数。
'仅有一个元素时，填什么都返回本身。
'一维元组时，指向第p1个元素。
'二维元组时，指向第p1个元组的第p2个元素。
Function part(ByVal str As String, Optional ByVal p1 As Byte = 1, Optional ByVal p2 As Byte = 1) As String
    
    If p1 <= 0 Or p2 <= 0 Then part = CVErr(2042) '错误，打死捣乱填负数零的
    
    Dim tmp, tmp2, i As Byte
    
    tmp = Split(str, "|")
    
    If UBound(tmp) = 0 Then '最多一维元组
        tmp2 = Split(tmp(0), "+")
        If p1 <= UBound(tmp2) + 1 Then
            part = tmp2(p1 - 1)   '返回一维元组第n个
        Else
            part = CVErr(2042)  '错误，大于元组上标
        End If
    Else    '最少二维元组
        If p1 <= UBound(tmp) + 1 Then
            tmp2 = Split(tmp(p1 - 1), "+")  '跳入指定元组
            '###以下结构基本同上，略不同
            If p2 <= UBound(tmp2) + 1 Then
                part = tmp2(p2 - 1)   '返回*指定*元组第n个
            Else
                part = CVErr(2042)  '错误，大于元组上标
            End If
            '###以上结构基本同上，略不同
        Else
            part = CVErr(2042)    '错误，大于二维元组上标
        End If
    End If
    
End Function

'2017/8/5，用于将格式化字符串数据转化为二维元组，引用part函数。
Private Function get_data(str As String)
    Dim arr(), tmp, tmp2
    Dim i As Byte, j As Byte
    
    tmp = Split(str, "|")
    tmp2 = Split(tmp(0), "+")
    p1 = UBound(tmp) + 1
    p2 = UBound(tmp2) + 1
    
    ReDim arr(p1, p2)
    For i = 1 To p1
        For j = 1 To p2
            If p1 = 1 Then  '一维元组
                arr(i, j) = part(str, j)
            Else    '多维元组
                arr(i, j) = part(str, i, j)
            End If
        Next
    Next
    get_data = arr  '返回二维元组
End Function

'############################################################################################
'2017/8/7，zc，修改格式化字符串数据，引用get_data及part函数。
'pos表示替换第几个位置的数据。
'method表示修改方法，目前支持+-*/以及sub（替换涵盖替换、增、删）、fix（强制改变值）
'value1为修改方法对应的第一个值，+-*/即对应值，sub表示替换前字符串
'value2为修改方法对应的第二个值，仅sub使用，表示替换后字符串
'2017/9/30,新增fix方法，强制改变值，默认方法为sub
Function change(ByVal str As String, ByVal pos As Byte, Optional ByVal method As String = "sub", Optional ByVal value1 As String = "0", Optional ByVal value2 As String = "") As String
    
    On Error GoTo catch '错误捕捉
    
    Dim data, i As Byte, j As Byte
    data = get_data(str)
    If pos > UBound(data, 2) Or pos = 0 Then GoTo catch '错误，引用位置多于最大量或为0
    
    Select Case method  '修改数据
    Case "+"
        For i = 1 To UBound(data, 1)     '每组第pos个加value1
            data(i, pos) = data(i, pos) + CDbl(value1)
        Next
    Case "-"
        For i = 1 To UBound(data, 1)     '每组第pos个减value1
            data(i, pos) = data(i, pos) - CDbl(value1)
        Next
    Case "*"
        For i = 1 To UBound(data, 1)     '每组第pos个乘value1
            data(i, pos) = data(i, pos) * CDbl(value1)
        Next
    Case "/"
        For i = 1 To UBound(data, 1)     '每组第pos个除value1
            data(i, pos) = data(i, pos) / CDbl(value1)
        Next
    Case "sub"
        For i = 1 To UBound(data, 1)     '每组第pos个value1内容替换为value2
            data(i, pos) = Replace(data(i, pos), value1, value2)
        Next
    Case "fix"
        For i = 1 To UBound(data, 1)     '每组第pos个的内容替换为value1
            data(i, pos) = value1
        Next
    Case Else
        change = CVErr(2042) '错误，不支持的方法
    End Select
    
    For i = 1 To UBound(data, 1)   '组装数据
        For j = 1 To UBound(data, 2)
            change = change & data(i, j)
            If j <> UBound(data, 2) Then change = change & "+"
        Next
        If i <> UBound(data, 1) Then change = change & "|"
    Next
    
    Exit Function
    
catch:
    change = CVErr(2042)

End Function

'2017/10/31 拆分change系列函数
Function cplus(ByVal str As String, ByVal pos As Byte, Optional ByVal value1 As String = "1") As String
    cplus = change(str, pos, "+", value1, value2)
End Function
Function cminus(ByVal str As String, ByVal pos As Byte, Optional ByVal value1 As String = "1") As String
    cminus = change(str, pos, "-", value1, value2)
End Function
Function cmul(ByVal str As String, ByVal pos As Byte, Optional ByVal value1 As String = "2") As String
    cmul = change(str, pos, "*", value1, value2)
End Function
Function cdiv(ByVal str As String, ByVal pos As Byte, Optional ByVal value1 As String = "2") As String
    cdiv = change(str, pos, "/", value1, value2)
End Function
Function csub(ByVal str As String, ByVal pos As Byte, Optional ByVal value1 As String = "0", Optional ByVal value2 As String = "1") As String
    csub = change(str, pos, "sub", value1, value2)
End Function
Function cfix(ByVal str As String, ByVal pos As Byte, Optional ByVal value1 As String = "0") As String
    cfix = change(str, pos, "fix", value1, value2)
End Function

'############################################################################################
'2017/10/30 正则匹配式修改
'默认筛掉所有<>内含内容
'<  起始
'.+?    任意字符，非贪婪
'> 结束
Function resub(s As String, Optional pat As String = "<.+?>", Optional rep As String = "", Optional glo As Boolean = True, Optional ign As Boolean = True) As String
    Dim re As Object
    Set re = CreateObject("Vbscript.Regexp")
    re.Global = glo 'True表示匹配所有, False表示仅匹配第一个符合项
    re.ignorecase = ign  'True表示不区分大小写, False表示区分大小写
    re.Pattern = pat    '匹配模式
    resub = re.Replace(s, rep)
    Set re = Nothing
End Function

'2017/11/9 正则匹配，取出匹配项
'默认取出最后一个\以后的内容
'which：取出第几个匹配项
'[^\\]以\开始
'+后接任意
'$至字符串尾
Function rematch(s As String, Optional pat As String = "[^\\]+$", Optional which As Integer = 0, Optional glo As Boolean = True, Optional ign As Boolean = True) As String
    Dim re As Object, mas As Object
    Set re = CreateObject("Vbscript.Regexp")
    re.Global = glo 'True表示匹配所有, False表示仅匹配第一个符合项
    re.ignorecase = ign  'True表示不区分大小写, False表示区分大小写
    re.Pattern = pat    '匹配模式
    Set mas = re.Execute(s)
    rematch = mas(which)
    Set mas = Nothing
    Set re = Nothing
End Function

'############################################################################################
'2017/10/30-31取得字符串某字符或背景的色值
'rng：字符串所在单元格（不可为多单元格区域）
'letter：提取的单字所在字符串第几个
'isFontColor：True为判断字色，False为判断背景色
Function getcolor(rng As Range, Optional letter As Integer = 1, Optional isFontColor = True) As Long
    If isFontColor Then
        getcolor = rng.Characters(letter, 1).Font.Color
    Else
        getcolor = rng.Interior.Color
    End If
End Function

'2017/10/30取出字符串中指定颜色的所有文字
'rng：字符串所在单元格（不可为多单元格区域）
'col：提取颜色的色值（可用getcolor获得）
Function ctext(rng As Range, Optional col As Long = 255) As String
    Dim l As Integer, s As String
    For l = 1 To Len(rng)
        If rng.Characters(l, 1).Font.Color = col Then
            s = rng.Characters(l, 1).Text
            ctext = ctext & s
        End If
    Next
End Function

'2017/10/30取出字符串中加粗的所有文字
Function btext(rng As Range) As String
    Dim l As Integer, s As String
    For l = 1 To Len(rng)
        If rng.Characters(l, 1).Font.Bold = True Then
            s = rng.Characters(l, 1).Text
            btext = btext & s
        End If
    Next
End Function

'2017/10/30取出字符串中斜体的所有文字
Function itext(rng As Range) As String
    Dim l As Integer, s As String
    For l = 1 To Len(rng)
        If rng.Characters(l, 1).Font.Italic = True Then
            s = rng.Characters(l, 1).Text
            itext = itext & s
        End If
    Next
End Function

'2017/10/30取出字符串中下划线的所有文字
'isSingle：True单下划线，False双下划线
Function utext(rng As Range, Optional isSingle As Boolean = True) As String
    Dim l As Integer, s As String
    For l = 1 To Len(rng)
        If isSingle And rng.Characters(l, 1).Font.Underline = xlUnderlineStyleSingle _
            Or Not isSingle And rng.Characters(l, 1).Font.Underline = xlUnderlineStyleDouble Then
            s = rng.Characters(l, 1).Text
            utext = utext & s
        End If
    Next
End Function

'############################################################################################
'2017/11/8字符串数组变为vba数组（自动判断一维或二维数组）。
'sRow：行分割符
'sCol：列分割符
'二维访问：arr(row,col)；上标row-Ubound(arr,1)、col-Ubound(arr,2)
'所有数组下标为0
Function arr(ByVal str As String, Optional sRow As String = "|", Optional sCol As String = "+")
    Dim tmp, tmp2(), tmp3, i As Byte, j As Byte
    If Right(str, 1) = sRow Or Right(str, 1) = sCol Then str = Left(str, Len(str) - 1)  '若有，去除最后一个多余分割符
    tmp = Split(str, sRow)
    If UBound(tmp) = 0 Then
        arr = Split(str, sCol)
    Else
        tmp3 = Split(tmp(i), sCol)
        ReDim tmp2(0 To UBound(tmp), 0 To UBound(tmp3))
        For i = 0 To UBound(tmp)
            For j = 0 To UBound(tmp3)
                On Error GoTo colerror
                tmp2(i, j) = Split(tmp(i), sCol)(j)
            Next
        Next
        arr = tmp2
    End If
    Exit Function
    
colerror: '列不齐的情况，补位为0
    tmp2(i, j) = 0
    Resume Next
End Function

'vba数组输出至excel区域（自动判断一维或二维数组）
'rng：可以使用Sheets("sheet1").[A1]，或ActiveCell，必须用call调用
Function tablize(arr, rng)
    On Error GoTo dim1
    rng.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1) = arr
    Exit Function

dim1:
    rng.Resize(1, UBound(arr) + 1) = arr
End Function

'############################################################################################
'2018/9/14
'根据指定条件，找出某区域内所有符合条件的单元格值，并以分隔符号分割，返回一个单元格
'@param range: 条件所在的Excel区域
'@param criteria: 条件表达式
'@param join_range: 需要组合的Excel区域
'@param sep: 分隔符号，默认,
'@return: 单元格值,单元格值,单元格值...
Function joinif(range As range, ByRef criteria As String, join_range As range, Optional sep As String = ",")
    Dim r
    Dim result As String
    Dim arr1, arr2
    arr1 = range
    arr2 = join_range
    For r = 1 To Application.Min(UBound(arr1), Sheets(range.Parent.Name).UsedRange.Rows.Count)  '暂时只优化至最小行数
        If arr2(r, 1) <> "" And arr1(r, 1) = criteria Then
            result = result & arr2(r, 1) & sep
        End If
    Next
    If Right(result, 1) = sep Then result = Left(result, Len(result) - 1)
    joinif = result
End Function