<%
'取URL路径
Function GetUrl(total) 
	'response.Write total
	Dim ScriptAddress, M_ItemUrl, M_item
	ScriptAddress = CStr(Request.ServerVariables("SCRIPT_NAME"))&"?"'取得当前地址 
	If (Request.QueryString <> "") Then 
		M_ItemUrl = "" 
	For Each M_item In Request.QueryString 
		If InStr("page",M_Item)=0 and InStr("total",M_Item)=0 Then 
			M_ItemUrl = M_ItemUrl & M_Item &"="& Server.URLEncode(Request.QueryString(""&M_Item&"")) & "&" 
		End If 
	Next 
	ScriptAddress = ScriptAddress & M_ItemUrl'取得带参数地址 
	End If 
	GetUrl = ScriptAddress & "total="&total&"&page=" 
End Function 


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''分页函数
''2008-3-14
''Author:一支笔
''total_record  总记录数
''current_page  当前页
''PCount		   页码排列
''pagesize	   每页记录数
''showpageNum   是否显示循环页码
''showpagetotal	是否显示总记录
''IsEnglish		是否显示英文分页
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub PageIndexUrl(total_record,current_page,PCount,pagesize,showpageNum,showpagetotal,IsEnglish)
	'''''''''''定义文字表达'''''''''''''''
	Dim Total_Text
	Dim First_Text 
	Dim Next_Text
	Dim Pre_Text
	Dim Last_Text
	Dim Record_Unit
	Dim Page_Unit
	
	if IsEnglish then:Total_Text = "Total Record:":else:Total_Text = "总记录:":end if
	if IsEnglish then:First_Text= "Frist":else:First_Text= "首页":end if
	if IsEnglish then:Next_Text= "next":else:Next_Text= "下一页":end if
	if IsEnglish then:Pre_Text= "pre":else:Pre_Text= "上一页":end if
	if IsEnglish then:Last_Text= "Last":else:Last_Text= "末页":end if
	if IsEnglish then:Record_Unit= "record":else:Record_Unit= "条":end if
	if IsEnglish then:Page_Unit= "page":else:Page_Unit= "页":end if
	''''''''''''''''''''''''''
	Dim PageIndex,pagestart,PageEnd
	Dim total_page:total_page = total_record \ pagesize   '总页数
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
	Dim Temp_X:Temp_X = (current_page \ PCount)  '取整数
	Dim Temp_Y:Temp_Y = (current_page mod PCount)  '取余数
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
'分页函数结束
%> 