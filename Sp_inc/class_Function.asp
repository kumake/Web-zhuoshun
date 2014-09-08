<%
'*******************************************************************
'	程序名：常用函数类
'	工程名：通用类
'	文件名：Class_PublicFunction.asp
'	版  本：V2.3
'	作  者：一支笔（317340145）
'	更新日期:2006-02-21 13:41:52
'*******************************************************************
%>
<%
	'过滤常量f打头
	Const fNo = 0
	Const fInt = 1
	Const fReal = 8
	Const fString = 2
	Const fLCase = 3
	Const fUCase = 4
	Const fNumberAndString = 5
	Const fAll = 6
	Const fForSQL = 7
	
	'文件打开方式
	Const ForReading = 1, ForWriting = 2

	'默认基本对象
	'在二代中，各个通用类自己负责变量声明，这个类对象名，都是默认了的，不许修改
	Dim F
	Set F = New HotCMS_Class_PublicFunction

Class HotCMS_Class_PublicFunction
	Public Ver,ClassName	'类版本,名称
	Public Debug	'是否调试
	Private vFSO, vDic, vReg
	Private vAct, vSrc
	Public SelfName
	
	'类初始化
	Private Sub Class_Initialize()
		Ver = "2.3"
		ClassName = "HotCMS_Class_PublicFunction"
		Debug = false 	'禁用调试
		
		SelfName = Request.ServerVariables("PATH_TRANSLATED")
		SelfName = Mid(SelfName, InstrRev(SelfName,"\") + 1, Len(SelfName))
	End Sub

	'类结束时执行的过程
	Private Sub Class_Terminate()
		If IsObject(vFSO) Then Set vFSO = Nothing
		If IsObject(vDic) Then Set vDic = Nothing
		If IsObject(vReg) Then Set vReg = Nothing
	End Sub

'**********************常用属性*****************
	'动作
	Public Property Get Act()
		If IsNull(vAct) Or IsEmpty(vAct) Then vAct = R("Act", fLCase)
		Act = vAct
	End Property

	'来源
	Public Property Get Src()
		If IsNull(vSrc) Or IsEmpty(vSrc) Then vSrc = R("Src", fLCase)
		Src = vSrc
	End Property
'**********************常用属性结束*****************

'**********************常用对象*****************
	'获取一个文件系统对象
	Public Property Get FSO()
		If Not IsObject(vFSO) Then Set vFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set FSO = vFSO
	End Property

	'获取一个字典对象
	Public Property Get Dic()
		If Not IsObject(vDic) Then Set vDic = Server.CreateObject("Scripting.Dictionary")
		Set Dic = vDic
	End Property

	'初始化一个正则表达式对象
	Private Sub InitReg()
		If Not IsObject(vReg) Then
			Set vReg = New RegExp
		End If
		vReg.Global = True
		vReg.IgnoreCase = True
	End Sub

	'获取一个正则表达式对象
	Public Property Get Reg()
		If Not IsObject(vReg) Then
			Call InitReg()
		End If
		Set Reg = vReg
	End Property
'********************常用对象结束***************

'************************测试用过程开始******************************************
	'基本输出过程
	Public Sub O(ByVal OutString)
		Response.Write OutString
	End Sub
	
	'重定向
	Public Sub Redirect(str)
		Response.Redirect(str)
	End Sub
	
	'红色换行输出。适用于运行时调试
	Public Sub B(OutString)
		O vbNewLine & "<BR><font color = red>" & OutString & "</font>"
	End Sub

	'调试专用(绿色换行输出)
	Public Sub T(ByVal OutString)
		If Not Debug Then Exit Sub

		'20070430 增加特殊数据类型的支持，使得调试输出更友好
		If IsNull(OutString) Then
			O "<BR><font color = red>{</font><font color = green>&lt;<font color = red>Null</font>&gt;</font><font color = red>}</font>" & vbNewLine
			Exit Sub
		End If
		If IsEmpty(OutString) Then
			O "<BR><font color = red>{</font><font color = green>&lt;<font color = red>Empty</font>&gt;</font><font color = red>}</font>" & vbNewLine
			Exit Sub
		End If
		If IsObject(OutString) Then
			O "<BR><font color = red>{</font><font color = green>&lt;<font color = red>Object</font>&gt;" & TypeName(OutString) & "</font><font color = red>}</font>" & vbNewLine
			Exit Sub
		End If
		
		Dim THead
		Thead = ""
		'20070430 把OutString强制转为字符串
		OutString = OutString & ""
		While Left(OutString, 1) = chr(32) Or Left(OutString, 1) = chr(9)
			If Left(OutString, 1) = chr(32) Then Thead = Thead & "&nbsp;"
			If Left(OutString, 1) = chr(9) Then Thead = Thead & "&nbsp;&nbsp;&nbsp;&nbsp;"
			OutString = Right(OutString, Len(OutString) - 1)
		Wend
		O "<BR><font color = red>{</font><font color = green>" & Thead & Server.HTMLEncode(OutString) & "</font><font color = red>}</font>" & vbNewLine
	End Sub

	'程序执行步骤
	Public Sub DebugStep(ByVal OutString)
		If Not Debug Then Exit Sub
		
		Dim THead
		Thead = ""
		While Left(OutString, 1) = chr(32) Or Left(OutString, 1) = chr(9)
			If Left(OutString, 1) = chr(32) Then Thead = Thead & "&nbsp;"
			If Left(OutString, 1) = chr(9) Then Thead = Thead & "&nbsp;&nbsp;&nbsp;&nbsp;"
			OutString = Right(OutString, Len(OutString) - 1)
		Wend
		O "<BR><font color = red>当前步骤：</font><font color = blue>" & Thead & OutString & "</font>" & vbNewLine
	End Sub
'************************测试用过程结束******************************************

'*******************************接收类函数****************************	
	'标准数据接收
	Public Function R(ByVal Qstr, ByVal FilterType)
		T "F.R(" & Qstr & "," & FilterType & ")"
		R = RR(Qstr)
		Select Case FilterType
			Case fNo	'0	'直接返回
				R = R & ""	'防止空数据
			Case fInt	'1	'数字
				If R = "" Or Not IsNumeric(R) Then
					R = 0
				Else
					R = CLng("0" & R)
				End If
			Case fString	'2	'字母
				'R = FL(R, fString)
				If Not Test(R, "^[a-zA-Z]+$") Then R=""
			Case fLCase	'3'小写
				R = LCase(R)
			Case fUCase	'4	'大写
				R = UCase(R)
			Case fNumberAndString	'5	'数字/大小写字母
				'R = FL(R, fNumberAndString)
				If Not Test(R, "^\w+$") Then R=""
			Case fAll	'6	'数字/大小写字母/符号
				'R = FL(R, fAll)
			Case fForSQL	'7	'防SQL
				'R = FL(R, fForSQL)
				If Test(R, "['#%;]") Then R=""
			Case fReal	'8	'浮点数
				If Not Test(R, "^-?\d+(\.\d*)?(e-?\d*)?$") Then R=0.0
			Case Else		'外部正则表达式
				'R = FL(R, FilterType)
				If Not Test(R, FilterType) Then R=""
		End Select
		If Debug Then O " Out={<font color=red>" & Server.HTMLEncode(Left(R,20)) & "</font>} Type=<font color=red>" & TypeName(R) & "</font>"
	End Function
	
	'基本数据接收。如果不需要类型判断，可直接调用这里
	Public Function RR(ByVal Qstr)
		If Qstr = "" Then Exit Function
		If Request.QueryString(Qstr)<>"" Then
			RR = Request.QueryString(Qstr)
		ElseIf F.ReplaceFormText(Qstr)<>"" Then
			RR = F.ReplaceFormText(Qstr)
		ElseIf Request.Cookies(Qstr)<>"" Then
			RR = Request.Cookies(Qstr)
		Else
			Dim temp
			If Instr(1, Qstr, "$$$", 1)>0 Then temp = Left(Qstr, Instr(1, Qstr, "$$$", 1)-1)
			If temp<>"" Then
				If Request.Cookies(temp)(Replace(Qstr,temp & "$$$",""))<>"" Then
					RR = Request.Cookies(temp)(Replace(Qstr,temp & "$$$",""))
				Else
					RR = Request(Qstr)
				End If
			Else
				RR = Request(Qstr)
			End If
		End If
	End Function
'*******************************接收类函数结束****************************	

'*******************************过滤类函数开始****************************	
	Public Function FL(Str, FilterType)
		Select Case FilterType
			Case fForSQL	'0	'常用过滤,防SQL注入
				FL = FL2(Str,"['#%;]")
			Case fInt	'1	'数字
				FL = FL1(Str, "^-?\d+$")
				If IsNull(FL) Or IsEmpty(FL) Or FL = "" Then FL = 0
				FL = CLng(FL)
			Case fString	'2	'字母
'				FL = FL1(Str, "[a-zA-Z]+")
				FL = FL2(Str, "[^a-zA-Z]")
			Case fLCase		'3	'小写字母
				FL = FL2(Str, "[^a-z]")
'				FL = FL1(Str, "[a-z]+")
			Case fUCase		'4	'大写字母
				FL = FL2(Str, "[^A-Z]")
'				FL = FL1(Str, "[A-Z]+")
			Case fNumberAndString	'5	'数字/大小写字母/下划线，这是标准变量
				FL = FL1(Str, "\w+")
			Case fAll	'6	'数字/大小写字母/可显示符号
				FL = Str
			Case fReal	'8	'浮点数
				FL = FL1(Str, "^-?\d+(\.\d*)?(e-?\d*)?$")
			Case Else
				FL = FL1(Str, FilterType)
		End Select
	End Function

	'只返回允许的字符
	Public Function FL1(Str,AllowStr)
		If Not IsObject(vReg) Then Call InitReg()
		vReg.Pattern = AllowStr
		If vReg.Test(Str) Then FL1 = Str
	End Function

	'删除不允许的字符
	Public Function FL2(Str, NotAllowStr)
		If Not IsObject(vReg) Then Call InitReg()
		vReg.Pattern = NotAllowStr
		FL2 = vReg.Replace(Str, "")
	End Function
'*******************************过滤类函数结束****************************

	'抛出异常信息
	Public Sub Raise(msg)
		O "<BR>错误：" & "<font color = red>" & msg & "</font>"
		Finish()
		'Response.End()
		'Err.Raise vbObjectError + 1, "通用类", msg
	End Sub

	'截去左边的Leng个字符，Leng也可以是字符串，此时截去左边第一个出现的Leng
	Public Function LCut(Byval str, Byval Leng)
		If Leng & "" = "" Then Exit Function
		LCut = ""
		If IsNumeric(Leng) Then	'截去左边的Leng个字符
			If str = "" Or Leng<0 Then Exit Function
			LCut = Right(str, Len(str)-Leng)
		Else	'截去左边第一个出现的Leng
			If Instr(str, Leng)>0 Then
				LCut = Left(str, Instr(str, Leng)-1)&Right(str, Len(str)-Instr(str, Leng)-Len(Leng)+1)
			End If
		End If
	End Function

	Public Function RCut(Byval str, Byval Leng)
		If Leng & "" = "" Then Exit Function
		RCut = ""
		If IsNumeric(Leng) Then
			If str = "" Or Leng<0 Then Exit Function
			RCut = Left(str, Len(str)-Leng)
		Else
			If Instr(str, Leng)>0 Then
				RCut = Left(str, InstrRev(str, Leng)-1)&Right(str, Len(str)-InstrRev(str, Leng)-Len(Leng)+1)
			End If
		End If
	End Function
	
	'Str是否以s开头
	Public Function StartWith(ByVal Str, ByVal s)
		StartWith = False
		If Left(Str, Len(s)) = s Then StartWith = True
	End Function
	
	'是否为Null或Empty或""
	Public Function IsNullOrEmpty(ByVal value)
		IsNullOrEmpty = False
		If IsNull(value) Then IsNullOrEmpty = True
		If IsEmpty(value) Then IsNullOrEmpty = True
		If Trim(value & "") = "" Then IsNullOrEmpty = True
	End Function

	Public Property Get GetIP()
		GetIP = Request.ServerVariables("REMOTE_HOST")
	End Property

	Public Property Get ScriptName()
		ScriptName = Request.ServerVariables("SCRIPT_NAME")
	End Property

	'取得带端口的URL，推荐使用
	Public Property Get Url()
		If request.Servervariables("Server_PORT") = "80" Then
			Url = "http://" & request.Servervariables("Server_name")&Replace(LCase(request.Servervariables("script_name")),ScriptName,"")
		Else
			Url = "http://" & request.Servervariables("Server_name") & ":" & request.Servervariables("Server_PORT")&Replace(LCase(request.Servervariables("script_name")),ScriptName,"")
		End If
	End Property

	'计算字符串的字节长度
	Public Function BLen(ByVal str)
		BLen = 0
		If str = "" Then Exit Function
		BLen = Len(str)
		Dim i
		for i = 1 to Len(str)
			if Asc(Mid(str,i,1))<0 then BLen = BLen + 1	'中文的ASC为负数
		next
	End Function
	
	'返回一定长度的随机产生的字符串
	Public Function RndStr(ByVal StrLen)
		Randomize()
		Dim i, r, ti
		r =  LowerStr & UpperStr & "_" & Left(NumStr, 10)
		For i = 1 To StrLen
			ti = 0
			while ti = 0
				ti = CInt(Rnd() * 100)
				If Not(ti>=1 And ti<=Len(r)) Then ti = 0
				If i = 1 And ti > 52 Then ti = 0	'大于52表明不是字母开头了
			wend
			RndStr = RndStr & Mid(r, ti, 1)
		Next
	End Function
	
	'如果是时间,则返回时间格式,否则返回MoRen
	Public Function FDate(ByVal inDate, MoRen)
		FDate = MoRen & ""
		If FDate = "" Then FDate = Now()
		If IsDate(FDate) Then FDate = CDate(FDate)
		If IsEmpty(inDate) Or IsNull(inDate) Then Exit Function
		If IsDate(inDate) Then FDate = CDate(inDate)
	End Function
	
	'正则表达式测试。判断是否匹配
	Public Function Test(Str, regstr)
		Call InitReg()
		vReg.Pattern = regstr
		Test = vReg.Test(Str)
	End Function
	
	'正则表达式替换，类似Replace
	Public Function Rep(vS, vF, vD)
		Call InitReg()
		vReg.Pattern = vF
		Rep = vReg.Replace(vS, vD)
	End Function
	
	'正则表达式取匹配位置，类似Instr
	Public Function Ins(vS, vF)
		Ins = -1
		If vS = "" Or vF = "" Then Exit Function
		Call InitReg()
		vReg.Pattern = vF
		Dim TemRegm
		Set TemRegm = vReg.Execute(vS)
		If TemRegm.Count>0 Then Ins = TemRegm(0).FirstIndex + 1	'它会从0算起
	End Function

	'读文本文件
	Function ReadTextFile(fPath)
		Dim Fo
		ReadTextFile = ""
		If Mid(fPath, 2, 1) <> ":" Then fPath = Server.MapPath(fPath)
		If Not FSO.FileExists(fPath) Then Response.Write("<BR>找不到文件：" & fPath) : Exit Function
		Set Fo = FSO.OpenTextFile(fPath, 1)
		If Not Fo.AtEndOfStream Then ReadTextFile = Fo.ReadAll
		Fo.Close
	End Function
	
	'写文本文件
	Function WriteTextFile(Byval fPath,Byval Content)
		'response.End()
		Dim Fo
		If Mid(fPath, 2, 1) <> ":" Then fPath = Server.MapPath(fPath)
		Dim fp
		fp = Left(fPath, InstrRev(fPath, "\")-1)
		If Not FSO.FolderExists(fp) Then FSO.CreateFolder fp
		Set Fo = FSO.CreateTextFile(fPath, true)
		Fo.Write Content & ""
		Fo.Close
	End Function
	
	'写文本文件
	Function WriteToFile(Byval fPath,Byval Content)
		'response.Write fPath&"\"&filename
		'response.End()
		Const ForReading = 1, ForWriting = 2
		Dim fso, f
		If Mid(fPath, 2, 1) <> ":" Then fPath = Server.MapPath(fPath)
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(fPath, ForWriting, True)
		f.Write Content
	End Function
	
	Function DeleteFile(Byval fPath) 
		Dim msg:msg = false
		If Mid(fPath, 2, 1) <> ":" Then fPath = Server.MapPath(fPath)
		Set Sys = Server.CreateObject("Scripting.FileSystemObject") 
		If Sys.FileExists(fPath) Then 
			Sys.DeleteFile(fPath) 
			msg =true
		Else 
			msg =false
		End If 
		Set Sys = Nothing 
		DeleteFile = msg 
	End Function
	
	''''''
	'''验证用户输入代码
	Function ReplaceFormText(Byval finput) 
		dim tempinput:tempinput = request.Form(finput)
		if tempinput<>"" then
			tempinput = replace(tempinput,"'","")
			ReplaceFormText = tempinput
		else
			ReplaceFormText = ""
		end if
	End Function	
	
	
	'跟踪
	Function Trace()
		Dim traceHead, traceEnd
		traceHead = "<table border=""0"" cellpadding=""1"" cellspacing=""1"" class=""TraceView"" style=""border-collapse:collapse;""><thead>"
		traceHead = traceHead & "<th width=""100"" class=""elm"">变量名</th><th class=""elmvalue"">值</th><th>长度</th><th>类型ID</th><th>类型名</th><th>备注</th>"
		traceHead = traceHead & "</thead>"
		traceEnd = "</table>"
		
		O "Application:"
		O traceHead
		TraceApplication
		O traceEnd

		O "Session:"
		O "&nbsp;&nbsp;<font color=green>SessionID</font> = <font color=red>" & Session.SessionID & "</font>"
		O "&nbsp;&nbsp;<font color=green>CodePage</font> = <font color=red>" & Session.CodePage & "</font>"
		O "&nbsp;&nbsp;<font color=green>LCID</font> = <font color=red>" & Session.LCID & "</font>"
		O traceHead
		TraceSession
		O traceEnd

		O "F.ReplaceFormText:"
		O traceHead
		TraceForm
		O traceEnd

		O "Request.QueryString:"
		O traceHead
		TraceQueryString
		O traceEnd

		O "Request.Cookies:"
		O traceHead
		TraceCookies
		O traceEnd
	End Function
	
	'跟踪Application
	Function TraceApplication()
		Dim elm, i
		For Each elm In Application.Contents
			O "<tr><td class=""elm""><font color=blue>" & elm & "</font></td>"
			TraceShow Application.Contents.Item(elm), ""
			O "</tr>"
		Next
	End Function
	
	'跟踪Session
	Function TraceSession()
		Dim elm, i
		For Each elm In Session.Contents
			O "<tr><td class=""elm""><font color=blue>" & elm & "</font></td>"
			TraceShow Session.Contents.Item(elm), ""
			O "</tr>"
		Next
	End Function
	
	'跟踪Form
	Function TraceForm()
		Dim elm, i
		For Each elm In F.ReplaceFormText
			O "<tr><td class=""elm""><font color=blue>" & elm & "</font></td>"
			TraceShow F.ReplaceFormText(elm), ""
			O "</tr>"
		Next
	End Function
	
	'跟踪QueryString
	Function TraceQueryString()
		Dim elm, i
		For Each elm In Request.QueryString
			O "<tr><td class=""elm""><font color=blue>" & elm & "</font></td>"
			TraceShow Request.QueryString(elm), ""
			O "</tr>"
		Next
	End Function
	
	'跟踪Cookies
	Function TraceCookies()
		Dim elm, i
		For Each elm In Request.Cookies
			O "<tr><td class=""elm""><font color=blue>" & elm & "</font></td>"
			TraceShow Request.Cookies(elm), ""
			O "</tr>"
		Next
	End Function
	
	Function GetVarStr(ID, vValue)
		If IsEmpty(vValue) Then
			GetVarStr = "<td class=""elmvalue"">&lt;<font color=red>Empty</font>&gt;</td>"
			GetVarStr = GetVarStr & "<td>0</td>"
		ElseIf IsNull(vValue) Then
			GetVarStr = "<td class=""elmvalue"">&lt;<font color=red>Null</font>&gt;</td>"
			GetVarStr = GetVarStr & "<td>0</td>"
		ElseIf IsObject(vValue) Then
			GetVarStr = "<td class=""elmvalue"">&lt;<font color=red>Object</font>&gt;</td>"
			GetVarStr = GetVarStr & "<td>0</td>"
		ElseIf vValue & "" = "" Then
			GetVarStr = "<td class=""elmvalue"">&nbsp;</td>"
			GetVarStr = GetVarStr & "<td>0</td>"
		Else
			GetVarStr = "<td class=""elmvalue""><font color=red>" & Server.HTMLEncode(Left(vValue, 200)) & "</font></td>"
			GetVarStr = GetVarStr & "<td>" & Len(ID) & "</td>"
		End If
		GetVarStr = GetVarStr & "<td>" & VarType(ID) & "</td>"
		GetVarStr = GetVarStr & "<td>" & TypeName(ID) & "</td>"
		GetVarStr = GetVarStr & "<td>" & GetType(VarType(ID)) & "</td>"
	End Function
	
	Function GetType(iType)
		Select Case iType
		Case 0
			GetType = "Empty（未初始化）"
		Case 1
			GetType = "Null（无有效数据）"
		Case 2
			GetType = "整数"
		Case 3
			GetType = "长整数"
		Case 4
			GetType = "单精度浮点数"
		Case 5
			GetType = "双精度浮点数"
		Case 6
			GetType = "货币"
		Case 7
			GetType = "日期"
		Case 8
			GetType = "字符串"
		Case 9
			GetType = "Automation 对象"
		Case 10
			GetType = "错误"
		Case 11
			GetType = "Boolean"
		Case 12
			GetType = "Variant"	'（只和变量数组一起使用）
		Case 13
			GetType = "数据访问对象"
		Case 17
			GetType = "字节"
		Case Else
			If iType > 8192 Then
				GetType = GetType(iType - 8192) & " 数组"
			Else
				GetType = ""
			End If
		End Select
	End Function
	
	Function TraceShow(ID, strHead)
		Select Case TypeName(ID)
		Case "Byte" '字节值
			O GetVarStr(ID, ID)
		Case "String" ' 字符串值 
			O GetVarStr(ID, ID)
		Case "Boolean" ' Boolean 值；True 或 False 
			If ID Then
				O GetVarStr(ID, "True")
			Else
				O GetVarStr(ID, "False")
			End If
		Case "Long"	'整数
			O GetVarStr(ID, ID)
		Case "Empty" ' 未初始化 
			O GetVarStr(ID, "(Empty) 未初始化")
		Case "Null" ' 无有效数据 
			O GetVarStr(ID, "(Null) 无有效数据")
		Case "<object type>" ' 实际对象类型名 
			O GetVarStr(ID, ID)
		Case "Object" ' 一般对象 
			O GetVarStr(ID, ID)
		Case "Unknown" ' 未知对象类型 
			O GetVarStr(ID, "")
		Case "Nothing" ' 还未引用对象实例的对象变量 
			O GetVarStr(ID, "")
		Case "Error" ' 错误 
			O GetVarStr(ID, "")
		Case "IReadCookie" ' Cookie 
			O GetVarStr(ID, ID & "")
		Case "Variant()" ' 数组 
			ShowArray ID, strHead
		Case Else 
			O GetVarStr(ID, ID)
		End Select
	End Function
	
	Function ShowArray(arr, strHead)
		If Not IsArray(arr) Then Exit Function

		O "<td class=""elmvalue"">&nbsp;</td><td>" & Ubound(arr) - Lbound(arr) + 1 & "</td><td>&nbsp;</td><td>&nbsp;</td><td>" & GetType(VarType(arr)) & "</td>"
		Dim k
		For k = Lbound(arr) To Ubound(arr)
			O "<tr><td align=Right class=""elm"">" & strHead & "(<font color=blue>" & k & "</font>)</td>"
			TraceShow arr(k), strHead & "(" & k & ")"
			O "</tr>"
		Next
	End Function
End Class
%>