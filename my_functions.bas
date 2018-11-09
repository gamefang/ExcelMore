Attribute VB_Name = "ԭ��"
'############################################################################################test
'2017/8/1��zc�������ڴ�ͼ���ļ�
Sub FileOpen()
    If ActiveCell.Value Like "*.*" Then    '��ͼ���ļ�
        Const targetpath = "src\fbclient\resfile\ui\images\" '��Ҫ���ļ���Ŀ¼���ϼ�Ŀ¼project
        Dim fulpath As String, filename As String
        fullpath = Split(ActiveWorkbook.Path, "design") '��excel����designĿ¼ȷ����Ŀ·��
        filename = fullpath(0) & targetpath & ActiveCell.Value
        Shell "cmd.exe /c" & filename
    End If
End Sub

'############################################################################################
'2017/8/4��zc����������html��������
'ÿ��ʹ��"|"�ָ���ǿ����ʹ��"{"��"}"������
'Դ�ַ������ݲ��ɰ���"|"��"{"��"}"��
'����1  str Դ�ַ���
'����2  before  ÿ��ǰ���ݣ�ͨ����ʾ�ӵ㣬Ĭ�ϼ�"<img src='img://Uiicon_zhuangbeitip_dian.png'>"
'����3  accent_style    ǿ������������ʽ����ʽ"��ɫ+�ֺ�"����ʡ���ֺţ�Ĭ��"#ffcc33+12"
'����4  normal_style    ��ͨ����������ʽ����ʽ"��ɫ+�ֺ�"����ʡ���ֺţ�Ĭ��"#e5d2ac"
Function html( _
    str As String, _
    Optional before As String = "<img src='img://Uiicon_zhuangbeitip_dian.png'>", _
    Optional accent_style As String = "#ffcc33+12", _
    Optional normal_style As String = "#e5d2ac") As String
    
    Dim pre_nor, suf_nor, pre_acc, suf_acc As String
    Dim content
    Dim i As Byte, line As String
    
    pre_nor = get_pre(normal_style) '�õ���ͨ����html
    suf_nor = "</font>"
    pre_acc = get_pre(accent_style) '�õ�ǿ������html
    suf_acc = "</font>"
    
    content = Split(str, "|")   '�ַ�����
    For i = 0 To UBound(content)
        line = content(i)   'ȡ�ַ�������
        line = Replace(line, "{", pre_acc)  '�滻ǿ��ǰ׺
        line = Replace(line, "}", suf_acc)  '�滻ǿ����׺
        line = before & pre_nor & line & suf_nor & "<br>" '����ǰ׺����ͼƬ�������У���������ͨ����
        html = html & line  '���·���ֵ
    Next
    
End Function

'��html�������ã�����html������ָ���Ϊ"+"��
Private Function get_pre(str As String) As String
    Dim tmp
    tmp = Split(str, "+")
    If UBound(tmp) = 0 Then '��ָ����ɫ��1��Ԫ�أ�
        get_pre = "<font color='" & tmp(0) & "'>"
    ElseIf UBound(tmp) = 1 Then 'ָ������ɫ�������С
        get_pre = "<font color='" & tmp(0) & "' size='" & tmp(1) & "'>"
    Else
        MsgBox UBound(tmp) & "���ô���"
    End If
End Function

'2017/8/4��zc��html�����������ڽ�html��仹ԭ�ؿ��ӻ���䡣
'����pre_acc��ָ��ǿ��ǰ׺������Ϊ����<>���ݡ�ǿ�����������Զ����롣
'����endl��ָ�����з���һ��Ϊ<br>��
Function h2s(str As String, Optional pre_acc As String = "<font color='#ffcc33' size='12'>", Optional endl As String = "<br>") As String

    Dim pass As Boolean, accent As Boolean, i As Integer, c As String
    
    str = Replace(str, pre_acc, "{")    '�滻ǿ��ǰ׺
    str = Replace(str, endl, "|")   '�ָ���
    
    pass = False
    accent = False
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If c = "<" Then '��������ģʽ
            pass = True
            If accent Then  '����ǿ����������һ��"<"�����ǿ��������ǿ���رշ�
                h2s = h2s & "}"
                accent = False
            End If
        ElseIf c = ">" Then '�ر�����ģʽ
            pass = False
        ElseIf Not pass Then    '����ַ�
            h2s = h2s & c
            If c = "{" Then accent = True   '��ǿ������
        End If
    Next
    
    If Right(h2s, 1) = "|" Then h2s = Left(h2s, Len(h2s) - 1) 'ȥ�����һλ�ķָ���
    
End Function

'############################################################################################
'2017/9/6����ѡ������ϲ�Ϊ�ַ�����2018/10/12�ɷ�������
'@param sel: range����
'@param sign: ���ӷ���Ĭ��,
'@param reverse: �Ƿ������ӣ�Ĭ��False
'@return: ���Ӻ���ַ���
Function join(sel As range, Optional ByVal sign As String = ",", Optional ByVal reverse As Boolean = False) As String
    For Each i In sel
        If i <> "" Then  '2018/8/22 ��ӿյ�Ԫ�񲻼�����߼�
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
'2017/8/5��zc���������ݵĺ�����
'����һ��Ԫ��ʱ����ʲô�����ر���
'һάԪ��ʱ��ָ���p1��Ԫ�ء�
'��άԪ��ʱ��ָ���p1��Ԫ��ĵ�p2��Ԫ�ء�
Function part(ByVal str As String, Optional ByVal p1 As Byte = 1, Optional ByVal p2 As Byte = 1) As String
    
    If p1 <= 0 Or p2 <= 0 Then part = CVErr(2042) '���󣬴�������������
    
    Dim tmp, tmp2, i As Byte
    
    tmp = Split(str, "|")
    
    If UBound(tmp) = 0 Then '���һάԪ��
        tmp2 = Split(tmp(0), "+")
        If p1 <= UBound(tmp2) + 1 Then
            part = tmp2(p1 - 1)   '����һάԪ���n��
        Else
            part = CVErr(2042)  '���󣬴���Ԫ���ϱ�
        End If
    Else    '���ٶ�άԪ��
        If p1 <= UBound(tmp) + 1 Then
            tmp2 = Split(tmp(p1 - 1), "+")  '����ָ��Ԫ��
            '###���½ṹ����ͬ�ϣ��Բ�ͬ
            If p2 <= UBound(tmp2) + 1 Then
                part = tmp2(p2 - 1)   '����*ָ��*Ԫ���n��
            Else
                part = CVErr(2042)  '���󣬴���Ԫ���ϱ�
            End If
            '###���Ͻṹ����ͬ�ϣ��Բ�ͬ
        Else
            part = CVErr(2042)    '���󣬴��ڶ�άԪ���ϱ�
        End If
    End If
    
End Function

'2017/8/5�����ڽ���ʽ���ַ�������ת��Ϊ��άԪ�飬����part������
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
            If p1 = 1 Then  'һάԪ��
                arr(i, j) = part(str, j)
            Else    '��άԪ��
                arr(i, j) = part(str, i, j)
            End If
        Next
    Next
    get_data = arr  '���ض�άԪ��
End Function

'############################################################################################
'2017/8/7��zc���޸ĸ�ʽ���ַ������ݣ�����get_data��part������
'pos��ʾ�滻�ڼ���λ�õ����ݡ�
'method��ʾ�޸ķ�����Ŀǰ֧��+-*/�Լ�sub���滻�����滻������ɾ����fix��ǿ�Ƹı�ֵ��
'value1Ϊ�޸ķ�����Ӧ�ĵ�һ��ֵ��+-*/����Ӧֵ��sub��ʾ�滻ǰ�ַ���
'value2Ϊ�޸ķ�����Ӧ�ĵڶ���ֵ����subʹ�ã���ʾ�滻���ַ���
'2017/9/30,����fix������ǿ�Ƹı�ֵ��Ĭ�Ϸ���Ϊsub
Function change(ByVal str As String, ByVal pos As Byte, Optional ByVal method As String = "sub", Optional ByVal value1 As String = "0", Optional ByVal value2 As String = "") As String
    
    On Error GoTo catch '����׽
    
    Dim data, i As Byte, j As Byte
    data = get_data(str)
    If pos > UBound(data, 2) Or pos = 0 Then GoTo catch '��������λ�ö����������Ϊ0
    
    Select Case method  '�޸�����
    Case "+"
        For i = 1 To UBound(data, 1)     'ÿ���pos����value1
            data(i, pos) = data(i, pos) + CDbl(value1)
        Next
    Case "-"
        For i = 1 To UBound(data, 1)     'ÿ���pos����value1
            data(i, pos) = data(i, pos) - CDbl(value1)
        Next
    Case "*"
        For i = 1 To UBound(data, 1)     'ÿ���pos����value1
            data(i, pos) = data(i, pos) * CDbl(value1)
        Next
    Case "/"
        For i = 1 To UBound(data, 1)     'ÿ���pos����value1
            data(i, pos) = data(i, pos) / CDbl(value1)
        Next
    Case "sub"
        For i = 1 To UBound(data, 1)     'ÿ���pos��value1�����滻Ϊvalue2
            data(i, pos) = Replace(data(i, pos), value1, value2)
        Next
    Case "fix"
        For i = 1 To UBound(data, 1)     'ÿ���pos���������滻Ϊvalue1
            data(i, pos) = value1
        Next
    Case Else
        change = CVErr(2042) '���󣬲�֧�ֵķ���
    End Select
    
    For i = 1 To UBound(data, 1)   '��װ����
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

'2017/10/31 ���changeϵ�к���
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
'2017/10/30 ����ƥ��ʽ�޸�
'Ĭ��ɸ������<>�ں�����
'<  ��ʼ
'.+?    �����ַ�����̰��
'> ����
Function resub(s As String, Optional pat As String = "<.+?>", Optional rep As String = "", Optional glo As Boolean = True, Optional ign As Boolean = True) As String
    Dim re As Object
    Set re = CreateObject("Vbscript.Regexp")
    re.Global = glo 'True��ʾƥ������, False��ʾ��ƥ���һ��������
    re.ignorecase = ign  'True��ʾ�����ִ�Сд, False��ʾ���ִ�Сд
    re.Pattern = pat    'ƥ��ģʽ
    resub = re.Replace(s, rep)
    Set re = Nothing
End Function

'2017/11/9 ����ƥ�䣬ȡ��ƥ����
'Ĭ��ȡ�����һ��\�Ժ������
'which��ȡ���ڼ���ƥ����
'[^\\]��\��ʼ
'+�������
'$���ַ���β
Function rematch(s As String, Optional pat As String = "[^\\]+$", Optional which As Integer = 0, Optional glo As Boolean = True, Optional ign As Boolean = True) As String
    Dim re As Object, mas As Object
    Set re = CreateObject("Vbscript.Regexp")
    re.Global = glo 'True��ʾƥ������, False��ʾ��ƥ���һ��������
    re.ignorecase = ign  'True��ʾ�����ִ�Сд, False��ʾ���ִ�Сд
    re.Pattern = pat    'ƥ��ģʽ
    Set mas = re.Execute(s)
    rematch = mas(which)
    Set mas = Nothing
    Set re = Nothing
End Function

'############################################################################################
'2017/10/30-31ȡ���ַ���ĳ�ַ��򱳾���ɫֵ
'rng���ַ������ڵ�Ԫ�񣨲���Ϊ�൥Ԫ������
'letter����ȡ�ĵ��������ַ����ڼ���
'isFontColor��TrueΪ�ж���ɫ��FalseΪ�жϱ���ɫ
Function getcolor(rng As Range, Optional letter As Integer = 1, Optional isFontColor = True) As Long
    If isFontColor Then
        getcolor = rng.Characters(letter, 1).Font.Color
    Else
        getcolor = rng.Interior.Color
    End If
End Function

'2017/10/30ȡ���ַ�����ָ����ɫ����������
'rng���ַ������ڵ�Ԫ�񣨲���Ϊ�൥Ԫ������
'col����ȡ��ɫ��ɫֵ������getcolor��ã�
Function ctext(rng As Range, Optional col As Long = 255) As String
    Dim l As Integer, s As String
    For l = 1 To Len(rng)
        If rng.Characters(l, 1).Font.Color = col Then
            s = rng.Characters(l, 1).Text
            ctext = ctext & s
        End If
    Next
End Function

'2017/10/30ȡ���ַ����мӴֵ���������
Function btext(rng As Range) As String
    Dim l As Integer, s As String
    For l = 1 To Len(rng)
        If rng.Characters(l, 1).Font.Bold = True Then
            s = rng.Characters(l, 1).Text
            btext = btext & s
        End If
    Next
End Function

'2017/10/30ȡ���ַ�����б�����������
Function itext(rng As Range) As String
    Dim l As Integer, s As String
    For l = 1 To Len(rng)
        If rng.Characters(l, 1).Font.Italic = True Then
            s = rng.Characters(l, 1).Text
            itext = itext & s
        End If
    Next
End Function

'2017/10/30ȡ���ַ������»��ߵ���������
'isSingle��True���»��ߣ�False˫�»���
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
'2017/11/8�ַ��������Ϊvba���飨�Զ��ж�һά���ά���飩��
'sRow���зָ��
'sCol���зָ��
'��ά���ʣ�arr(row,col)���ϱ�row-Ubound(arr,1)��col-Ubound(arr,2)
'���������±�Ϊ0
Function arr(ByVal str As String, Optional sRow As String = "|", Optional sCol As String = "+")
    Dim tmp, tmp2(), tmp3, i As Byte, j As Byte
    If Right(str, 1) = sRow Or Right(str, 1) = sCol Then str = Left(str, Len(str) - 1)  '���У�ȥ�����һ������ָ��
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
    
colerror: '�в�����������λΪ0
    tmp2(i, j) = 0
    Resume Next
End Function

'vba���������excel�����Զ��ж�һά���ά���飩
'rng������ʹ��Sheets("sheet1").[A1]����ActiveCell��������call����
Function tablize(arr, rng)
    On Error GoTo dim1
    rng.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1) = arr
    Exit Function

dim1:
    rng.Resize(1, UBound(arr) + 1) = arr
End Function

'############################################################################################
'2018/9/14
'����ָ���������ҳ�ĳ���������з��������ĵ�Ԫ��ֵ�����Էָ����ŷָ����һ����Ԫ��
'@param range: �������ڵ�Excel����
'@param criteria: �������ʽ
'@param join_range: ��Ҫ��ϵ�Excel����
'@param sep: �ָ����ţ�Ĭ��,
'@return: ��Ԫ��ֵ,��Ԫ��ֵ,��Ԫ��ֵ...
Function joinif(range As range, ByRef criteria As String, join_range As range, Optional sep As String = ",")
    Dim r
    Dim result As String
    Dim arr1, arr2
    arr1 = range
    arr2 = join_range
    For r = 1 To Application.Min(UBound(arr1), Sheets(range.Parent.Name).UsedRange.Rows.Count)  '��ʱֻ�Ż�����С����
        If arr2(r, 1) <> "" And arr1(r, 1) = criteria Then
            result = result & arr2(r, 1) & sep
        End If
    Next
    If Right(result, 1) = sep Then result = Left(result, Len(result) - 1)
    joinif = result
End Function