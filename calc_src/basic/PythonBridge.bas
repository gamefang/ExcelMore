' 用于桥接python脚本的basic语言模块，注册后可实现python自定义函数

REM Keep a global reference to the ScriptProvider, since this stuff may be called many times: 
Global g_MasterScriptProvider as Object
REM Specify location of Python script, providing cell functions: 
Const URL_Main as String = "vnd.sun.star.script:" 
Const URL_Args as String = "?language=Python&location=user" 

Function invokePyFunc(file AS String, func As String, args As Array, outIdxs As Array, outArgs As Array)
   sURL = URL_Main & file & ".py$" & func & URL_Args
   oMSP = getMasterScriptProvider()
   On Local Error GoTo ErrorHandler
      oScript = oMSP.getScript(sURL)
      invokePyFunc = oScript.invoke(args, outIdxs, outArgs)
      Exit Function
   ErrorHandler:
      Dim msg As String, toFix As String
      msg = Error$
      toFix = ""
      If 1 = Err AND InStr(Error$, "an error occurred during file opening") Then
         msg = "Couldn' open the script file."
         toFix = "Make sure the 'python' folder exists in the user's Scripts folder, and that the former contains " & file & ".py."
      End If
      MsgBox msg & chr(13) & toFix, 16, "Error " & Err & " calling " & func
end Function

Function getMasterScriptProvider() 
   if isNull(g_MasterScriptProvider) then 
      oMasterScriptProviderFactory = createUnoService("com.sun.star.script.provider.MasterScriptProviderFactory") 
      g_MasterScriptProvider = oMasterScriptProviderFactory.createScriptProvider("") 
   endif 
   getMasterScriptProvider = g_MasterScriptProvider
End Function


' 配置项：python脚本的名称
' python脚本位置：...Users\XXXX\AppData\Roaming\LibreOffice\4\user\Scripts\python\
Const py_fn as String = "gamefang"

' 注册python函数为basic函数，从而可在自定义函数中使用
Function zpytype(value)
    zpytype = invokePyFunc(py_fn, "zpytype", Array(value), Array(), Array())
End Function

Function zjoin(v, Optional sep As String, Optional row_first, Optional keep_empty)
	If IsMissing(sep) Then
		sep = ","
	End If
	If IsMissing(row_first) Then
		row_first = 1
	End If
	If IsMissing(keep_empty) Then
		keep_empty = 0
	End If
    zjoin = invokePyFunc(py_fn, "zjoin", Array(v,sep,row_first,keep_empty), Array(), Array())
End Function

Function zfetch(v, Optional num As Integer, Optional sep As String)
	If IsMissing(num) Then
		num = 1
	End If
   If IsMissing(sep) Then
		sep = ","
	End If
   zfetch = invokePyFunc(py_fn, "zfetch", Array(v,num,sep), Array(), Array())
End Function

Function zmod(v, val, Optional num As Integer, Optional sep As String)
	If IsMissing(num) Then
		num = 1
	End If
   If IsMissing(sep) Then
		sep = ","
	End If
   zmod = invokePyFunc(py_fn, "zmod", Array(v,val,num,sep), Array(), Array())
End Function