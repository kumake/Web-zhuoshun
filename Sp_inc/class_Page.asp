<%
'ȡURL·��
Function GetUrl(total) 
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
	GetUrl = ScriptAddress & "total="&total&"&page=" 
End Function 


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''��ҳ����
''2008-3-14
''Author:һ֧��
''total_record  �ܼ�¼��
''current_page  ��ǰҳ
''PCount		   ҳ������
''pagesize	   ÿҳ��¼��
''showpageNum   �Ƿ���ʾѭ��ҳ��
''showpagetotal	�Ƿ���ʾ�ܼ�¼
''IsEnglish		�Ƿ���ʾӢ�ķ�ҳ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub PageIndexUrl(total_record,current_page,PCount,pagesize,showpageNum,showpagetotal,IsEnglish)
	'''''''''''�������ֱ��'''''''''''''''
	Dim Total_Text
	Dim First_Text 
	Dim Next_Text
	Dim Pre_Text
	Dim Last_Text
	Dim Record_Unit
	Dim Page_Unit
	
	if IsEnglish then:Total_Text = "Total Record:":else:Total_Text = "�ܼ�¼:":end if
	if IsEnglish then:First_Text= "Frist":else:First_Text= "��ҳ":end if
	if IsEnglish then:Next_Text= "next":else:Next_Text= "��һҳ":end if
	if IsEnglish then:Pre_Text= "pre":else:Pre_Text= "��һҳ":end if
	if IsEnglish then:Last_Text= "Last":else:Last_Text= "ĩҳ":end if
	if IsEnglish then:Record_Unit= "record":else:Record_Unit= "��":end if
	if IsEnglish then:Page_Unit= "page":else:Page_Unit= "ҳ":end if
	''''''''''''''''''''''''''
	Dim PageIndex,pagestart,PageEnd
	Dim total_page:total_page = total_record \ pagesize   '��ҳ��
	if total_record mod pagesize <>0 then
	  total_page = total_page + 1
	end if
%>
<div id="Divpage">
<%
if showpagetotal then
%>
	<li><%=Total_Text%><%=total_record%><%=Record_Unit%></li>
	<li><%=current_page%>/<%=total_page%><%=Page_Unit%></li>
<%
end if
%>
<li><%if current_page>1 then:Response.Write "<a href='"& GetUrl(total_record) &"1'>"&First_Text&"</a>":else:Response.Write ""&First_Text&"":end if%></li>
<li><%if current_page<=1 then:response.Write ""&Pre_Text&"":else:Response.Write "<a href='"& GetUrl(total_record) & (current_page-1) &"'>"&Pre_Text&"</a>":end if%></li>
<%
if showpageNum then
	Dim Temp_X:Temp_X = (current_page \ PCount)  'ȡ����
	Dim Temp_Y:Temp_Y = (current_page mod PCount)  'ȡ����
	if Temp_X<1 then
		pagestart = 0
		PageEnd = PCount+1
	elseif Temp_X=1 and Temp_Y=0 then
		pagestart = 0
		PageEnd = PCount+1
	elseif Temp_X=0 and Temp_Y=1 then	
		pagestart = 1
		PageEnd = 2*PCount+1
	elseif Temp_X>1 and Temp_Y=0 then
		pagestart = (Temp_X)*PCount-1
		PageEnd = (Temp_X+1)*PCount
	else
		pagestart = (Temp_X)*PCount
		PageEnd = (Temp_X+1)*PCount
	end if
	if PageEnd>total_page then:PageEnd = total_page:end if
	if pagestart=0 then:pagestart = 1:end if
	for PageIndex=pagestart to PageEnd
		Response.Write "<li><a href="& GetUrl(total_record) & PageIndex &">"
		if Cint(current_page) = Cint(PageIndex) then
			Response.Write "<font color='red'><U><b>"& PageIndex &"</b></U></font>"
		else
			Response.Write PageIndex
		end if
		Response.Write "</a></li>"&chr(13)
	next
end if
%>
<li><%if Cint(current_page) < Cint(total_page) then:Response.Write "<a href='"& GetUrl(total_record) & (current_page+1) &"'>"&Next_Text&"</a>":else:Response.Write ""&Next_Text&"":end if%></li>
<li><%if Cint(current_page) < Cint(total_page) then:Response.Write "<a href='"& GetUrl(total_record) & total_page &"'>"&Last_Text&"</a>":else:Response.Write ""&Last_Text&"":end if%></li>
</div>
<%
End sub 
'��ҳ��������
%> 