<%
'*******************************************************************
'	�����������ú�����
'	��������ͨ����
'	�ļ�����Class_PublicFunction.asp
'	��  ����V2.3
'	��  �ߣ�һ֧�ʣ�317340145��
'	��������:2006-02-21 13:41:52
'*******************************************************************
%>
<%
	'���˳���f��ͷ
	Const fNo = 0
	Const fInt = 1
	Const fReal = 8
	Const fString = 2
	Const fLCase = 3
	Const fUCase = 4
	Const fNumberAndString = 5
	Const fAll = 6
	Const fForSQL = 7
	
	'�ļ��򿪷�ʽ
	Const ForReading = 1, ForWriting = 2

	'Ĭ�ϻ�������
	'�ڶ����У�����ͨ�����Լ������������������������������Ĭ���˵ģ������޸�
	Dim F
	Set F = New HotCMS_Class_PublicFunction

Class HotCMS_Class_PublicFunction
	Public Ver,ClassName	'��汾,����
	Public Debug	'�Ƿ����
	Private vFSO, vDic, vReg
	Private vAct, vSrc
	Public SelfName
	
	'���ʼ��
	Private Sub Class_Initialize()
		Ver = "2.3"
		ClassName = "HotCMS_Class_PublicFunction"
		Debug = false 	'���õ���
		
		SelfName = Request.ServerVariables("PATH_TRANSLATED")
		SelfName = Mid(SelfName, InstrRev(SelfName,"\") + 1, Len(SelfName))
	End Sub

	'�����ʱִ�еĹ���
	Private Sub Class_Terminate()
		If IsObject(vFSO) Then Set vFSO = Nothing
		If IsObject(vDic) Then Set vDic = Nothing
		If IsObject(vReg) Then Set vReg = Nothing
	End Sub

'**********************��������*****************
	'����
	Public Property Get Act()
		If IsNull(vAct) Or IsEmpty(vAct) Then vAct = R("Act", fLCase)
		Act = vAct
	End Property

	'��Դ
	Public Property Get Src()
		If IsNull(vSrc) Or IsEmpty(vSrc) Then vSrc = R("Src", fLCase)
		Src = vSrc
	End Property
'**********************�������Խ���*****************

'**********************���ö���*****************
	'��ȡһ���ļ�ϵͳ����
	Public Property Get FSO()
		If Not IsObject(vFSO) Then Set vFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set FSO = vFSO
	End Property

	'��ȡһ���ֵ����
	Public Property Get Dic()
		If Not IsObject(vDic) Then Set vDic = Server.CreateObject("Scripting.Dictionary")
		Set Dic = vDic
	End Property

	'��ʼ��һ��������ʽ����
	Private Sub InitReg()
		If Not IsObject(vReg) Then
			Set vReg = New RegExp
		End If
		vReg.Global = True
		vReg.IgnoreCase = True
	End Sub

	'��ȡһ��������ʽ����
	Public Property Get Reg()
		If Not IsObject(vReg) Then
			Call InitReg()
		End If
		Set Reg = vReg
	End Property
'********************���ö������***************

'************************�����ù��̿�ʼ******************************************
	'�����������
	Public Sub O(ByVal OutString)
		Response.Write OutString
	End Sub
	
	'�ض���
	Public Sub Redirect(str)
		Response.Redirect(str)
	End Sub
	
	'��ɫ�������������������ʱ����
	Public Sub B(OutString)
		O vbNewLine & "<BR><font color = red>" & OutString & "</font>"
	End Sub

	'����ר��(��ɫ�������)
	Public Sub T(ByVal OutString)
		If Not Debug Then Exit Sub

		'20070430 ���������������͵�֧�֣�ʹ�õ���������Ѻ�
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
		'20070430 ��OutStringǿ��תΪ�ַ���
		OutString = OutString & ""
		While Left(OutString, 1) = chr(32) Or Left(OutString, 1) = chr(9)
			If Left(OutString, 1) = chr(32) Then Thead = Thead & "&nbsp;"
			If Left(OutString, 1) = chr(9) Then Thead = Thead & "&nbsp;&nbsp;&nbsp;&nbsp;"
			OutString = Right(OutString, Len(OutString) - 1)
		Wend
		O "<BR><font color = red>{</font><font color = green>" & Thead & Server.HTMLEncode(OutString) & "</font><font color = red>}</font>" & vbNewLine
	End Sub

	'����ִ�в���
	Public Sub DebugStep(ByVal OutString)
		If Not Debug Then Exit Sub
		
		Dim THead
		Thead = ""
		While Left(OutString, 1) = chr(32) Or Left(OutString, 1) = chr(9)
			If Left(OutString, 1) = chr(32) Then Thead = Thead & "&nbsp;"
			If Left(OutString, 1) = chr(9) Then Thead = Thead & "&nbsp;&nbsp;&nbsp;&nbsp;"
			OutString = Right(OutString, Len(OutString) - 1)
		Wend
		O "<BR><font color = red>��ǰ���裺</font><font color = blue>" & Thead & OutString & "</font>" & vbNewLine
	End Sub
'************************�����ù��̽���******************************************

'*******************************�����ຯ��****************************	
	'��׼���ݽ���
	Public Function R(ByVal Qstr, ByVal FilterType)
		T "F.R(" & Qstr & "," & FilterType & ")"
		R = RR(Qstr)
		Select Case FilterType
			Case fNo	'0	'ֱ�ӷ���
				R = R & ""	'��ֹ������
			Case fInt	'1	'����
				If R = "" Or Not IsNumeric(R) Then
					R = 0
				Else
					R = CLng("0" & R)
				End If
			Case fString	'2	'��ĸ
				'R = FL(R, fString)
				If Not Test(R, "^[a-zA-Z]+$") Then R=""
			Case fLCase	'3'Сд
				R = LCase(R)
			Case fUCase	'4	'��д
				R = UCase(R)
			Case fNumberAndString	'5	'����/��Сд��ĸ
				'R = FL(R, fNumberAndString)
				If Not Test(R, "^\w+$") Then R=""
			Case fAll	'6	'����/��Сд��ĸ/����
				'R = FL(R, fAll)
			Case fForSQL	'7	'��SQL
				'R = FL(R, fForSQL)
				If Test(R, "['#%;]") Then R=""
			Case fReal	'8	'������
				If Not Test(R, "^-?\d+(\.\d*)?(e-?\d*)?$") Then R=0.0
			Case Else		'�ⲿ������ʽ
				'R = FL(R, FilterType)
				If Not Test(R, FilterType) Then R=""
		End Select
		If Debug Then O " Out={<font color=red>" & Server.HTMLEncode(Left(R,20)) & "</font>} Type=<font color=red>" & TypeName(R) & "</font>"
	End Function
	
	'�������ݽ��ա��������Ҫ�����жϣ���ֱ�ӵ�������
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
'*******************************�����ຯ������****************************	

'*******************************�����ຯ����ʼ****************************	
	Public Function FL(Str, FilterType)
		Select Case FilterType
			Case fForSQL	'0	'���ù���,��SQLע��
				FL = FL2(Str,"['#%;]")
			Case fInt	'1	'����
				FL = FL1(Str, "^-?\d+$")
				If IsNull(FL) Or IsEmpty(FL) Or FL = "" Then FL = 0
				FL = CLng(FL)
			Case fString	'2	'��ĸ
'				FL = FL1(Str, "[a-zA-Z]+")
				FL = FL2(Str, "[^a-zA-Z]")
			Case fLCase		'3	'Сд��ĸ
				FL = FL2(Str, "[^a-z]")
'				FL = FL1(Str, "[a-z]+")
			Case fUCase		'4	'��д��ĸ
				FL = FL2(Str, "[^A-Z]")
'				FL = FL1(Str, "[A-Z]+")
			Case fNumberAndString	'5	'����/��Сд��ĸ/�»��ߣ����Ǳ�׼����
				FL = FL1(Str, "\w+")
			Case fAll	'6	'����/��Сд��ĸ/����ʾ����
				FL = Str
			Case fReal	'8	'������
				FL = FL1(Str, "^-?\d+(\.\d*)?(e-?\d*)?$")
			Case Else
				FL = FL1(Str, FilterType)
		End Select
	End Function

	'ֻ����������ַ�
	Public Function FL1(Str,AllowStr)
		If Not IsObject(vReg) Then Call InitReg()
		vReg.Pattern = AllowStr
		If vReg.Test(Str) Then FL1 = Str
	End Function

	'ɾ����������ַ�
	Public Function FL2(Str, NotAllowStr)
		If Not IsObject(vReg) Then Call InitReg()
		vReg.Pattern = NotAllowStr
		FL2 = vReg.Replace(Str, "")
	End Function
'*******************************�����ຯ������****************************

	'�׳��쳣��Ϣ
	Public Sub Raise(msg)
		O "<BR>����" & "<font color = red>" & msg & "</font>"
		Finish()
		'Response.End()
		'Err.Raise vbObjectError + 1, "ͨ����", msg
	End Sub

	'��ȥ��ߵ�Leng���ַ���LengҲ�������ַ�������ʱ��ȥ��ߵ�һ�����ֵ�Leng
	Public Function LCut(Byval str, Byval Leng)
		If Leng & "" = "" Then Exit Function
		LCut = ""
		If IsNumeric(Leng) Then	'��ȥ��ߵ�Leng���ַ�
			If str = "" Or Leng<0 Then Exit Function
			LCut = Right(str, Len(str)-Leng)
		Else	'��ȥ��ߵ�һ�����ֵ�Leng
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
	
	'Str�Ƿ���s��ͷ
	Public Function StartWith(ByVal Str, ByVal s)
		StartWith = False
		If Left(Str, Len(s)) = s Then StartWith = True
	End Function
	
	'�Ƿ�ΪNull��Empty��""
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

	'ȡ�ô��˿ڵ�URL���Ƽ�ʹ��
	Public Property Get Url()
		If request.Servervariables("Server_PORT") = "80" Then
			Url = "http://" & request.Servervariables("Server_name")&Replace(LCase(request.Servervariables("script_name")),ScriptName,"")
		Else
			Url = "http://" & request.Servervariables("Server_name") & ":" & request.Servervariables("Server_PORT")&Replace(LCase(request.Servervariables("script_name")),ScriptName,"")
		End If
	End Property

	'�����ַ������ֽڳ���
	Public Function BLen(ByVal str)
		BLen = 0
		If str = "" Then Exit Function
		BLen = Len(str)
		Dim i
		for i = 1 to Len(str)
			if Asc(Mid(str,i,1))<0 then BLen = BLen + 1	'���ĵ�ASCΪ����
		next
	End Function
	
	'����һ�����ȵ�����������ַ���
	Public Function RndStr(ByVal StrLen)
		Randomize()
		Dim i, r, ti
		r =  LowerStr & UpperStr & "_" & Left(NumStr, 10)
		For i = 1 To StrLen
			ti = 0
			while ti = 0
				ti = CInt(Rnd() * 100)
				If Not(ti>=1 And ti<=Len(r)) Then ti = 0
				If i = 1 And ti > 52 Then ti = 0	'����52����������ĸ��ͷ��
			wend
			RndStr = RndStr & Mid(r, ti, 1)
		Next
	End Function
	
	'�����ʱ��,�򷵻�ʱ���ʽ,���򷵻�MoRen
	Public Function FDate(ByVal inDate, MoRen)
		FDate = MoRen & ""
		If FDate = "" Then FDate = Now()
		If IsDate(FDate) Then FDate = CDate(FDate)
		If IsEmpty(inDate) Or IsNull(inDate) Then Exit Function
		If IsDate(inDate) Then FDate = CDate(inDate)
	End Function
	
	'������ʽ���ԡ��ж��Ƿ�ƥ��
	Public Function Test(Str, regstr)
		Call InitReg()
		vReg.Pattern = regstr
		Test = vReg.Test(Str)
	End Function
	
	'������ʽ�滻������Replace
	Public Function Rep(vS, vF, vD)
		Call InitReg()
		vReg.Pattern = vF
		Rep = vReg.Replace(vS, vD)
	End Function
	
	'������ʽȡƥ��λ�ã�����Instr
	Public Function Ins(vS, vF)
		Ins = -1
		If vS = "" Or vF = "" Then Exit Function
		Call InitReg()
		vReg.Pattern = vF
		Dim TemRegm
		Set TemRegm = vReg.Execute(vS)
		If TemRegm.Count>0 Then Ins = TemRegm(0).FirstIndex + 1	'�����0����
	End Function

	'���ı��ļ�
	Function ReadTextFile(fPath)
		Dim Fo
		ReadTextFile = ""
		If Mid(fPath, 2, 1) <> ":" Then fPath = Server.MapPath(fPath)
		If Not FSO.FileExists(fPath) Then Response.Write("<BR>�Ҳ����ļ���" & fPath) : Exit Function
		Set Fo = FSO.OpenTextFile(fPath, 1)
		If Not Fo.AtEndOfStream Then ReadTextFile = Fo.ReadAll
		Fo.Close
	End Function
	
	'д�ı��ļ�
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
	
	'д�ı��ļ�
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
	'''��֤�û��������
	Function ReplaceFormText(Byval finput) 
		dim tempinput:tempinput = request.Form(finput)
		if tempinput<>"" then
			tempinput = replace(tempinput,"'","")
			ReplaceFormText = tempinput
		else
			ReplaceFormText = ""
		end if
	End Function	
	
	
	'����
	Function Trace()
		Dim traceHead, traceEnd
		traceHead = "<table border=""0"" cellpadding=""1"" cellspacing=""1"" class=""TraceView"" style=""border-collapse:collapse;""><thead>"
		traceHead = traceHead & "<th width=""100"" class=""elm"">������</th><th class=""elmvalue"">ֵ</th><th>����</th><th>����ID</th><th>������</th><th>��ע</th>"
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
	
	'����Application
	Function TraceApplication()
		Dim elm, i
		For Each elm In Application.Contents
			O "<tr><td class=""elm""><font color=blue>" & elm & "</font></td>"
			TraceShow Application.Contents.Item(elm), ""
			O "</tr>"
		Next
	End Function
	
	'����Session
	Function TraceSession()
		Dim elm, i
		For Each elm In Session.Contents
			O "<tr><td class=""elm""><font color=blue>" & elm & "</font></td>"
			TraceShow Session.Contents.Item(elm), ""
			O "</tr>"
		Next
	End Function
	
	'����Form
	Function TraceForm()
		Dim elm, i
		For Each elm In F.ReplaceFormText
			O "<tr><td class=""elm""><font color=blue>" & elm & "</font></td>"
			TraceShow F.ReplaceFormText(elm), ""
			O "</tr>"
		Next
	End Function
	
	'����QueryString
	Function TraceQueryString()
		Dim elm, i
		For Each elm In Request.QueryString
			O "<tr><td class=""elm""><font color=blue>" & elm & "</font></td>"
			TraceShow Request.QueryString(elm), ""
			O "</tr>"
		Next
	End Function
	
	'����Cookies
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
			GetType = "Empty��δ��ʼ����"
		Case 1
			GetType = "Null������Ч���ݣ�"
		Case 2
			GetType = "����"
		Case 3
			GetType = "������"
		Case 4
			GetType = "�����ȸ�����"
		Case 5
			GetType = "˫���ȸ�����"
		Case 6
			GetType = "����"
		Case 7
			GetType = "����"
		Case 8
			GetType = "�ַ���"
		Case 9
			GetType = "Automation ����"
		Case 10
			GetType = "����"
		Case 11
			GetType = "Boolean"
		Case 12
			GetType = "Variant"	'��ֻ�ͱ�������һ��ʹ�ã�
		Case 13
			GetType = "���ݷ��ʶ���"
		Case 17
			GetType = "�ֽ�"
		Case Else
			If iType > 8192 Then
				GetType = GetType(iType - 8192) & " ����"
			Else
				GetType = ""
			End If
		End Select
	End Function
	
	Function TraceShow(ID, strHead)
		Select Case TypeName(ID)
		Case "Byte" '�ֽ�ֵ
			O GetVarStr(ID, ID)
		Case "String" ' �ַ���ֵ 
			O GetVarStr(ID, ID)
		Case "Boolean" ' Boolean ֵ��True �� False 
			If ID Then
				O GetVarStr(ID, "True")
			Else
				O GetVarStr(ID, "False")
			End If
		Case "Long"	'����
			O GetVarStr(ID, ID)
		Case "Empty" ' δ��ʼ�� 
			O GetVarStr(ID, "(Empty) δ��ʼ��")
		Case "Null" ' ����Ч���� 
			O GetVarStr(ID, "(Null) ����Ч����")
		Case "<object type>" ' ʵ�ʶ��������� 
			O GetVarStr(ID, ID)
		Case "Object" ' һ����� 
			O GetVarStr(ID, ID)
		Case "Unknown" ' δ֪�������� 
			O GetVarStr(ID, "")
		Case "Nothing" ' ��δ���ö���ʵ���Ķ������ 
			O GetVarStr(ID, "")
		Case "Error" ' ���� 
			O GetVarStr(ID, "")
		Case "IReadCookie" ' Cookie 
			O GetVarStr(ID, ID & "")
		Case "Variant()" ' ���� 
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