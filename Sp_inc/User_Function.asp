<%
	'''''����IDѡ��ҳ��������
	function UserSinglePageByID(SinglePagetype,SinglePageid)
		set rs = UIcon.Query("select U.Instr,U.content,U.picname from user_Singlepage U,Sp_SinglePage S where U.singlepageID = S.ID and U.ID="&SinglePageid&"")
		if rs.recordcount<>"" then
			if SinglePagetype="instr" then
				UserSinglePageInstr = rs("Instr")
			elseif SinglePagetype="content" then
				UserSinglePageInstr = rs("content")
			elseif SinglePagetype="pic" then
				UserSinglePageInstr = rs("picname")
			end if
		else
			UserSinglePageInstr = "�������Ϣ!"
		end if
	end function

	'''''''''''
	''''ȡ��ҳ���ݼ��
	function UserSinglePageInstr(singlepage)
		set rs = UIcon.Query("select U.Instr from user_Singlepage U,Sp_SinglePage S where U.singlepageID = S.ID and S.singlePagename='"&singlepage&"'")
		if rs.recordcount<>"" then
			if rs("Instr")<>"" then
				UserSinglePageInstr = rs("Instr")
			else
				UserSinglePageInstr = ""
			end if
		else
			UserSinglePageInstr = "�������Ϣ!"
		end if
	end function
	''''ȡ��ҳ����
	function UserSinglePage(singlepage)
		set rs = UIcon.Query("select U.content from user_Singlepage U,Sp_SinglePage S where U.singlepageID = S.ID and S.singlePagename='"&singlepage&"'")
		if rs.recordcount<>"" then
			if rs("content")<>"" then
				UserSinglePage = rs("content")
			else
				UserSinglePage = ""
			end if
		else
			UserSinglePage = "�������Ϣ!"
		end if
	end function
	'''''''''''
	''''ȡ��ҳͼƬ
	function UserSinglePagePic(singlepage)
		set rs = UIcon.Query("select U.picname from user_Singlepage U,Sp_SinglePage S where U.singlepageID = S.ID and S.singlePagename='"&singlepage&"'")
		if rs.recordcount<>"" then
			if rs("picname")<>"" then
				UserSinglePagePic = rs("picname")
			else
				UserSinglePagePic = ""
			end if
		else
			UserSinglePagePic = "�������Ϣ!"
		end if
	end function
	''''
	''''ϵͳ�ռ�
	function WriteLog(username,action,content)
		Con.Execute("insert into Sp_log (username,[action],content) values ('"&username&"','"&action&"','"&content&"')")	
	end function
	''ID,categoryID,categoryname,Pathstr,ParentID
	''��������˵���ʼ
	function fillDropDownList(ModelArray,parentID,defaultvalue)
		Dim spaceLen:spaceLen = 0
		for arraystep=0 to ubound(ModelArray,2) step 1
			'spaceLen = len(replace(ModelArray(3,arraystep),",",""))
			spaceLen = Ubound(split(ModelArray(3,arraystep),","))
			'if spaceLen<>0 then spaceLen = spaceLen - 1
			if ModelArray(4,arraystep)=parentID then
				response.Write " <option value="&ModelArray(0,arraystep)&""
				'if parentID=ModelArray(0,arraystep) then
				'response.Write " style='background-color:#FF0000;' "
				'end if
				if defaultvalue<>"" then
				 if Cint(defaultvalue) = Cint(ModelArray(0,arraystep)) then
					response.Write " selected"
				 end if
				end if
				response.Write ">"&string(spaceLen,"��")&ModelArray(2,arraystep)&"</option>"
				fillDropDownList ModelArray,ModelArray(0,arraystep),defaultvalue
			end if
		next
	end function
	
	''''�б��ѯɸѡ��Ϣ����
	function SearchDropDownList(ModelArray,parentID,defaultvalue)
		Dim spaceLen:spaceLen = 0
		for arraystep=0 to ubound(ModelArray,2) step 1
			'spaceLen = len(replace(ModelArray(3,arraystep),",",""))
			spaceLen = Ubound(split(ModelArray(3,arraystep),","))
			'if spaceLen<>0 then spaceLen = spaceLen - 1
			if ModelArray(4,arraystep)=parentID then
				response.Write " <option value='"&ModelArray(3,arraystep)&"'"
				if defaultvalue<>"" and Cstr(defaultvalue) = Cstr(ModelArray(3,arraystep)) then
					response.Write " selected"
				end if
				response.Write ">"&string(spaceLen,"��")&ModelArray(2,arraystep)&"</option>"
				SearchDropDownList ModelArray,ModelArray(0,arraystep),defaultvalue
			end if
		next
	end function
	''''''''''''''''''''''
	function filldictionaryDropDownList(dictionaryArray,parentID,defaultvalue)
		Dim spaceLen:spaceLen = 0
		for arraystep=0 to ubound(dictionaryArray,2) step 1
			'spaceLen = len(replace(dictionaryArray(2,arraystep),",",""))
			'if spaceLen<>0 then spaceLen = spaceLen - 1
			spaceLen = Ubound(split(dictionaryArray(2,arraystep),","))
			if dictionaryArray(3,arraystep)=parentID then
				response.Write " <option value="&dictionaryArray(0,arraystep)&""
				if defaultvalue<>"" then
				 if Cint(defaultvalue) = Cint(dictionaryArray(0,arraystep)) then
					response.Write " selected"
				 end if
				end if
				response.Write ">"&string(spaceLen,"��")&dictionaryArray(1,arraystep)&"</option>"
				filldictionaryDropDownList dictionaryArray,dictionaryArray(0,arraystep),defaultvalue
			end if
		next
	end function

	''''''
	function searchDictionaryDropDownList(dictionaryArray,parentID,defaultvalue)
		Dim spaceLen:spaceLen = 0
		for arraystep=0 to ubound(dictionaryArray,2) step 1
			spaceLen = Ubound(split(dictionaryArray(2,arraystep),","))
			if dictionaryArray(3,arraystep)=parentID then
				response.Write " <option value="&dictionaryArray(2,arraystep)&""
				if defaultvalue<>"" then
				 if Cstr(defaultvalue) = Cstr(dictionaryArray(2,arraystep)) then
					response.Write " selected"
				 end if
				end if
				response.Write ">"&string(spaceLen,"��")&dictionaryArray(1,arraystep)&"</option>"
				filldictionaryDropDownList dictionaryArray,dictionaryArray(0,arraystep),defaultvalue
			end if
		next
	end function

	''ȡ��β˵���ʼ
	function Listlevelmenu(ModelID,TempCategoryArray,StrPath)
		if StrPath<>"" then
			Dim strPathArray:strPathArray = Split(StrPath,",")
			for tempMenuStep = 0 to Ubound(strPathArray) step 1
				for templevelStep = 0 to Ubound(TempCategoryArray,2) step 1
					if Cint(TempCategoryArray(0,templevelStep))=Cint(strPathArray(tempMenuStep)) then 
						response.Write "<a href='?ModelID="&ModelID&"&categoryID="&TempCategoryArray(3,templevelStep)&"'>"&TempCategoryArray(2,templevelStep)&"</a>"
						if tempMenuStep<>Ubound(strPathArray) then response.Write ">>"
						exit for
					end if
				next
				'response.Write strPathArray(tempMenuStep)
			next
		end if
	end function
	
	'''''''''
	'''''''''
	function ListDictionaryLeav(TempCategoryArray,StrPath)
		if StrPath<>"" then
			Dim strPathArray:strPathArray = Split(StrPath,",")
			for tempMenuStep = 0 to Ubound(strPathArray) step 1
				for templevelStep = 0 to Ubound(TempCategoryArray,2) step 1
					if Cint(TempCategoryArray(0,templevelStep))=Cint(strPathArray(tempMenuStep)) then 
						response.Write "<a href='?pathstr="&TempCategoryArray(3,templevelStep)&"'>"&TempCategoryArray(2,templevelStep)&"</a>"
						if tempMenuStep<>Ubound(strPathArray) then response.Write ">>"
						exit for
					end if
				next
				'response.Write strPathArray(tempMenuStep)
			next
		end if
	end function
	
	'ȡURL·��
	Function GetpageItemUrl() 
		'response.Write total
		Dim ScriptAddress, M_ItemUrl, M_item
		ScriptAddress = CStr(Request.ServerVariables("SCRIPT_NAME"))&"?"'ȡ�õ�ǰ��ַ 
		If (Request.QueryString <> "") Then 
			M_ItemUrl = "" 
		For Each M_item In Request.QueryString 
			If InStr("page",M_Item)=0 and InStr("total",M_Item)=0 Then 
				M_ItemUrl = M_ItemUrl & M_Item &"="& Server.URLEncode(Request.QueryString(""&M_Item&"")) & "&" 
			End If 
		Next 
		ScriptAddress = ScriptAddress & M_ItemUrl'ȡ�ô�������ַ 
		End If 
		GetpageItemUrl = ScriptAddress & "page=" 
	End Function 

	''������Ϣ
	function Alert(msg,url)
		response.Write "<script language='javascript'>alert('Sp_CMSϵͳ��ʾ:\r\n\n"&msg&"');"
		if url<>"" then
			response.Write "location.href='"&url&"';"
		else
			response.Write "history.go(-1);"
		end if
		response.Write "</script>"
	end function
	
	function HtmlAlert(msg)
		Dim msgarray
		msgarray = split(msg,"#")
		response.Write "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd'>"&vbcrlf
		response.Write "<LINK href='images/style.css' type='text/css' rel='stylesheet'>"&vbcrlf
		response.Write "<DIV id=wrap>"&vbcrlf
		response.Write "	<UL id=tags>"&vbcrlf
		response.Write "	  <LI class='selectTag' style='LEFT: 0px; TOP: 0px'><A>ϵͳ����</A></LI>"&vbcrlf
		response.Write "	</UL>"&vbcrlf
		response.Write "	<DIV id='tagContent'>"&vbcrlf
		response.Write "	<DIV class='tagContent selectTag content' id=tagContent1>"&vbcrlf
		response.Write "	  <DIV>"&vbcrlf
		
		response.Write "	<TABLE cellSpacing='1' width='100%'>"&vbcrlf
		response.Write "	<TR class=content-td1>"&vbcrlf
		response.Write "		<TH align='left'>��Ϣ��ʾ��</TH>"&vbcrlf
		response.Write "	</TR>"&vbcrlf
		for i=0 to ubound(msgarray)
		response.Write "	<TR class=content-td1>"&vbcrlf
		response.Write "		<TD style='color:#FF0000;'>"&msgarray(i)&"</TD>"&vbcrlf
		response.Write "	</TR>"&vbcrlf
		next
		response.Write "	<TR class=content-td1>"&vbcrlf
		response.Write "		<TD><a href='javascript:history.go(-1);'>����</a></TD>"&vbcrlf
		response.Write "	</TR>"&vbcrlf
		response.Write "	</TABLE>"&vbcrlf
		response.Write "	  </DIV>"&vbcrlf
		response.Write "	</DIV>"&vbcrlf
		response.Write "	</DIV>"&vbcrlf
		response.Write "</DIV>"&vbcrlf
		response.End()
	end function


	''''��֤Url�����ĺϷ���
	function VerificationUrlParam(ByVal param,ByVal typed,ByVal defaultvalue)
		Dim QueryStr:QueryStr = request.QueryString(param)
		QueryStr = replace(QueryStr,"'","")
		select case typed
			case "int"
				if QueryStr<>"" and isnumeric(QueryStr) then
					VerificationUrlParam = QueryStr
				else
					VerificationUrlParam = defaultvalue
				end if
			case "datatime"
				if QueryStr<>"" and isdate(QueryStr) then
					VerificationUrlParam = request.QueryString(param)
				else
					VerificationUrlParam = defaultvalue
				end if
			case "array"
				if QueryStr<>"" and isarray(QueryStr) then
					VerificationUrlParam = QueryStr
				else
					VerificationUrlParam = defaultvalue
				end if
			case else
				if QueryStr<>"" then
					VerificationUrlParam = QueryStr
				else
					VerificationUrlParam = defaultvalue
				end if			
		end select
	end function
	
	''''''��¼��������
	Sub robot()
		Dim robots:robots="Baiduspider+@Baidu|Googlebot@Google|ia_archiver@Alexa|IAArchiver@Alexa|ASPSeek@ASPSeek|YahooSeeker@Yahoo|SogouBot@sogou|help.yahoo.com/help/us/ysearch/slurp@Yahoo|sohu-search@SOHU|MSNBOT@MSN"
		dim I1,I2,l1,l2,l3,i,rs
		l2=false
		l1=request.servervariables("http_user_agent")
		'response.Write l1
		F1=request.ServerVariables("SCRIPT_NAME")
		I1=split(robots,chr(124))
		for i=0 to ubound(I1)
				I2=split(I1(i),"@")
				'response.Write lcase(l1)&"<br>"&lcase(I2(0))&"<hr>"
				if instr(lcase(l1),lcase(I2(0)))>0 then
					l2=true:l3=I2(1):exit for
				end if
		next
		'response.End()
		if l2 and len(l3)>0 then'�����֩��,�͸���֩����Ϣ
				FilePath = Server.Mappath("robots/"&l3&"_robots.txt")
				'��¼֩������
				Set Fso = Server.CreateObject("Scripting.FileSystemObject")
				if not Fso.FileExists(FilePath) then
					'�����ı�
					Fso.CreateTextFile(FilePath)
				end if
				Set Fout = Fso.OpenTextFile(FilePath,8,True)
								Fout.WriteLine "����ҳ�棺"&F1
								Fout.WriteLine "֩�룺"&l3&chr(32)&chr(32)&"����ʱ�䣺"&Now()
								Fout.WriteLine "-----------------------------------------------"
								Fout.Close
				Set Fout = Nothing
				Set Fso = Nothing
		end if
	end Sub
	
	''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''
	'ת���ַ���
	function HtmlenCode(str) 
	if isNULL(str) then
	htmlencode2=""
	exit function
	end if
	dim result
	result=str
	result=replace(result,chr(13),"<br>")
	result=replace(result,"    ",chr(9))
	result=replace (result,chr(9),"����")
	htmlencode2=result
	end function

	Function outHTML(str)
		Dim sTemp
		sTemp = str
		outHTML = ""
		If IsNull(sTemp) = True Then
			Exit Function
		End If
		sTemp = Replace(sTemp, "&", "&amp;")
		sTemp = Replace(sTemp, "<", "&lt;")
		sTemp = Replace(sTemp, ">", "&gt;")
		sTemp = Replace(sTemp, Chr(34), "&quot;")
		sTemp = Replace(sTemp, Chr(10), "<br>")
		outHTML = sTemp
	End Function

	function Myleft(cString,i)
		dim tempLen,j
		tempLen = 0
	
		if not isnull(cString) then
		for j = 1 to Len(cString)
			if len(Hex(Asc(Mid(cString,j,1)))) > 2 then
				tempLen = tempLen + 2
			else
				tempLen = tempLen + 1
			end if
	
			if tempLen > i then
				Myleft = Left(cString,j-1)&".."
				exit for
			end if
		next
		
		if tempLen <= i then
			Myleft = cString
		end if
		end if
	end function	
	
	' ============================================
	' ȥ��Html��ʽ�����ڴ����ݿ���ȡ��ֵ���������ʱ
	' ע�⣺value="?"���һ��Ҫ��˫����
	' ============================================
	Function inHTML(str)
		Dim sTemp
		sTemp = str
		inHTML = ""
		If IsNull(sTemp) = True Then
			Exit Function
		End If
		sTemp = Replace(sTemp, "&", "&amp;")
		sTemp = Replace(sTemp, "<", "&lt;")
		sTemp = Replace(sTemp, ">", "&gt;")
		sTemp = Replace(sTemp, Chr(34), "&quot;")
		inHTML = sTemp
	End Function
	'''''���ܺ���
	function encryptCode(str)
		For stepdomain=1 to len(str)
			if encryptCode<>"" then
				encryptCode=encryptCode&"-"&(Asc(Mid(str,stepdomain,1))+stepdomain)
			else
				encryptCode=(Asc(Mid(str,stepdomain,1))+stepdomain)
			end if
		next
	end function
	
	''''''���ܺ���
	function DecryptCode(str)
		Dim TempArray:TempArray = split(str,"-")
		For stepdomain=0 to ubound(TempArray)
			if DecryptCode<>"" then
				DecryptCode=DecryptCode&(chr(TempArray(stepdomain)-(stepdomain+1)))
			else
				DecryptCode=chr(TempArray(stepdomain)-(stepdomain+1))
			end if
		next
	end function
	
	'''''''��֤�ļ�
	function ValidationCode(str)
		'''''���/admin/�ļ����µ���Ȩ�ļ��Ƿ����
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(server.MapPath("copyright.inc"))) Then
			ValidationCode = "true"
			Dim tempIdentifyingCode:tempIdentifyingCode = F.ReadTextFile ("copyright.inc")
			if tempIdentifyingCode<>"" then
				if tempIdentifyingCode=str then
					ValidationCode = "true"
				else
					ValidationCode = "false"
				end if
			else
				ValidationCode = "false"
			end if
		Else
			ValidationCode = "false"
		End If
	end function
	
	function VerAccessType(var)
		select case var
		case 3:
		VerAccessType = "int"
		case 7:
		VerAccessType = "datetime"
		case 202:
		VerAccessType = "nvarchar"
		case 203:
		VerAccessType = "ntext"
		case else
		VerAccessType = "ntext"
		end select
	end function
	
	'���ڸ�ʽ��1
	Function FormatDate1(DT,tp)
		dim Y,M,D
		'Y=right(Year(DT),2)
		Y=Year(DT)
		M=month(DT)
		D=Day(DT)
		if M<10 then M="0"&M
		if D<10 then D="0"&D
		select case tp
		case 1 FormatDate1=Y&"-"&M&"-"&D
		case 2 FormatDate1=Y&"��"&M&"��"&D&"��"
		end select
	End Function
	
	
	Function GetPage(url) 
	On Error Resume Next
		Dim Retrieval
		Set Retrieval = CreateObject("Msxml2.XMLHTTP") 
		'Set Retrieval = CreateObject("Microsoft.XMLHTTP") 
		With Retrieval 
		.Open "Get", url, False, "", "" 
		.Send
		if .readyState = 4 then
			if len(.ResponseBody)=0 then
				Response.Write("<h1>����������~��~~~~</h1>")
			else
				GetPage = BytesToBstr(.ResponseBody)
			end if
		end if
		End With 
		Set Retrieval = Nothing
	On Error GoTo 0
	End Function
	
	Function BytesToBstr(body)
		dim objstream
		set objstream = Server.CreateObject("adodb.stream")
		objstream.Type = 1
		objstream.Mode =3
		objstream.Open
		objstream.Write body
		objstream.Position = 0
		objstream.Type = 2
		objstream.Charset = "gb2312"
		'objstream.Charset = "UTF-8"
		BytesToBstr = objstream.ReadText 
		objstream.Close
		set objstream = nothing
	End Function
	
	Function GetContent(str,start,last,n)
	
		If Instr(lcase(str),lcase(start))>0 and Instr(lcase(str),lcase(last))>0 then
			select case n
			case 0	'���Ҷ���ȡ����ȡǰ�棩��ȥ���ؼ��֣�
			GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))-Len(start)+1) 
			GetContent=Left(GetContent,Instr(lcase(GetContent),lcase(last))-1)
			case 1	'���Ҷ���ȡ����ȡǰ�棩�������ؼ��֣�
			GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))+1)
			GetContent=Left(GetContent,Instr(lcase(GetContent),lcase(last))+Len(last)-1)
			case 2	'ֻ���ҽ�ȡ��ȡǰ��ģ���ȥ���ؼ��֣�
			GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))-Len(start)+1)
			case 3	'ֻ���ҽ�ȡ��ȡǰ��ģ��������ؼ��֣�
			GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))+1)
			case 4	'ֻ�����ȡ��ȡ����ģ��������ؼ��֣�
			GetContent=Left(str,InstrRev(lcase(str),lcase(start))+Len(start)-1)
			case 5	'ֻ�����ȡ��ȡ����ģ���ȥ���ؼ��֣�
			GetContent=Left(str,InstrRev(lcase(str),lcase(start))-1)
			case 6	'ֻ�����ȡ��ȡǰ��ģ��������ؼ��֣�
			GetContent=Left(str,Instr(lcase(str),lcase(start))+Len(start)-1)
			case 7	'ֻ���ҽ�ȡ��ȡ����ģ��������ؼ��֣�
			GetContent=Right(str,Len(str)-InstrRev(lcase(str),lcase(start))+1)
			case 8	'ֻ�����ȡ��ȡǰ��ģ���ȥ���ؼ��֣�
			GetContent=Left(str,Instr(lcase(str),lcase(start))-1)
			case 9	'ֻ���ҽ�ȡ��ȡ����ģ��������ؼ��֣�
			GetContent=Right(str,Len(str)-InstrRev(lcase(str),lcase(start)))
			end select
		Else
			GetContent=""
		End if
		
	End function
	'���˿ո� �س� �Ʊ��
	Function filtrate(str)
		str=replace(str,chr(13),"")
		str=replace(str,chr(10),"")
		str=replace(str,chr(9),"")
		filtrate=str
	End Function
	
	
	Function toUTF8(szInput)
	Dim wch, uch, szRet
	Dim x
	Dim nAsc, nAsc2, nAsc3
	
	If szInput = "" Then
	toUTF8 = szInput
	Exit Function
	End If
	For x = 1 To Len(szInput)
	  wch = Mid(szInput, x, 1)
	  nAsc = AscW(wch)
	  If nAsc < 0 Then nAsc = nAsc + 65536
	
	  If (nAsc And &HFF80) = 0 Then
		 szRet = szRet & wch
	  Else
		  If (nAsc And &HF000) = 0 Then
			 uch = "%" & Hex(((nAsc \ 2 ^ 6)) or &HC0) & Hex(nAsc And &H3F or &H80)
			 szRet = szRet & uch
		   Else
			  uch = "%" & Hex((nAsc \ 2 ^ 12) or &HE0) & "%" & _
			  Hex((nAsc \ 2 ^ 6) And &H3F or &H80) & "%" & _
			  Hex(nAsc And &H3F or &H80)
			  szRet = szRet & uch
		   End If
	  End If
	Next
	
	toUTF8 = szRet
	End Function
%>