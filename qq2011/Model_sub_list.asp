<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	'action
	Dim action:action = VerificationUrlParam("action","string","")
	'��ģ����Ϣ��ʶID
	Dim ModelItemID:ModelItemID = VerificationUrlParam("ModelItemID","int","0")
	'��Ϣ��ʶID
	Dim ItemID:ItemID = VerificationUrlParam("ID","int","0")
	'��ģ��ID
	Dim ParentModelID:ParentModelID = VerificationUrlParam("ParentModelID","int","0")
	'ģ��ID
	Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","0")
	'cityID
	Dim cityID:cityID = VerificationUrlParam("cityID","int","0")
	'�������ɸѡ
	Dim categoryID:categoryID = VerificationUrlParam("categoryID","int","0")
	'���Źؼ���ɸѡ
	Dim keyword:keyword = VerificationUrlParam("keyword","string","")
	'��Ÿ�ģ�����ݵ�һά����
	Dim Modelarray
	'ģ������
	Dim ModelName
	'��Ÿ�ģ���������ݱ�
	Dim ModelTableName
	'���ģ������ID
	Dim ModelCategoryID
	'������Ķ�ά����
	Dim ModelCategoryArray
	'��ŵ����Ķ�ά����
	Dim CityCategoryArray
	''�Ƿ���������
	Dim IsAllowComment
	''����ģ��ID
	Dim SubModelID
	'''''''''''''''''''''''
	if ModelID=0 then
		response.Write "ģ�ͽ��ͳ��ִ���!"
	else
		'ȡ��Ӧģ�͵����ݴ�����ݱ�
		Modelarray = Con.QueryRow("select ID,modelname,modeltable,modelCategoryID,ISAllowComment,SubModelID from Sp_Model where Id="&ModelID&"",0) 
		ModelName = Modelarray(1)		
		ModelTableName = Modelarray(2)		
		ModelCategoryID = Modelarray(3)
		IsAllowComment = Modelarray(4) 
		SubModelID = Modelarray(5) 
		'''ISdel=1ɾ����Ϣ
		if action<>"" and action="del" then
			Con.Execute("update "&ModelTableName&" set IsDel=1 where ID="&ItemID&"")
			'''''''
			WriteLog CK("username"),"action","����վģ�� "&ModelName&" ����"
			'''''''
			Alert "��Ϣ�������վ�ɹ�","Model_list.asp?ModelID="&ModelID&""
		elseif action="remove" then 
			Con.Execute("delete * from "&ModelTableName&" where ID="&ItemID&"")
			'''''''
			WriteLog CK("username"),"action","ɾ��ģ�� "&ModelName&" ����"
			'''''''
			Alert "ɾ����Ϣ�ɹ�","Model_list.asp?ModelID="&ModelID&""
		elseif action="restore" then 
			Con.Execute("update "&ModelTableName&" set IsDel=0 where ID="&ItemID&"")
			'''''''
			WriteLog CK("username"),"action","��ԭģ�� "&ModelName&" ����"
			'''''''
			Alert "��Ϣ��ԭ�ɹ�","Model_list.asp?ModelID="&ModelID&""
		elseif action="BatchDel" then
		 	Dim BatchremoveItemID:BatchremoveItemID = F.ReplaceFormText("ItemUID")
			Con.Execute("delete * from "&ModelTableName&" where ID in ("&BatchremoveItemID&")")
			'''''''
			WriteLog CK("username"),"action","����ɾ��ģ�� "&ModelName&" ����"
			'''''''
			Alert "����ɾ���ɹ�","Model_list.asp?ModelID="&ModelID&""
		'elseif action="BatchDel" then 
			'Dim BatchDelItemID:BatchDelItemID = F.ReplaceFormText("ItemUID")
			'Con.Execute("update "&ModelTableName&" set IsDel=1 where ID in ("&BatchDelItemID&")")
			'Alert "����ɾ����Ϣ�ɹ�","Model_list.asp?ModelID="&ModelID&""
		elseif action="batchRecycle" then
			Dim BatchrestoreItemID:BatchrestoreItemID = F.ReplaceFormText("ItemUID")
			Con.Execute("update "&ModelTableName&" set IsDel=1 where ID in ("&BatchrestoreItemID&")")
			'''''''
			WriteLog CK("username"),"action","��������վģ�� "&ModelName&" ����"
			'''''''
			Alert "������Ϣ����վ�ɹ�","Model_list.asp?ModelID="&ModelID&"&action=Recycle"
		end if
		'''
		'''ȡ���ݴ�Ž�ModelCategoryArray
		ModelCategoryArray = Con.QueryData("select D.ID,D.categoryID,D.categoryname,D.Pathstr,D.ParentID from Sp_dictionaryCategory C,Sp_dictionary D where C.id=D.categoryID and D.categoryID="&ModelCategoryID&" order by D.id desc")
		CityCategoryArray = Con.QueryData("select ID,Areaname,pathstr,parentID,Depth from Sp_city")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript">
		function ModelSearch()
		{		
			var obj = document.getElementById("txt_CityID").value;
			var obj1 = document.getElementById("sel_category").value;
			var obj2 = document.getElementById("txt_keyword").value;
			var url = "?ModelID=<%=ModelID%>&cityid="+obj;
			if(obj1!="")
			{
				url = url + "&categoryID="+obj1;
			}
			if(obj2!="")
			{
				url = url + "&keyword="+obj2;
			}
			location.href=url;
		}
		//////����������Ϣ
		function SelectALLCheck(obj)
		{
			for (var i=0;i<document.forms[0].elements.length;i++) 
			{ 
				var e = document.forms[0].elements[i]; 
				if (e.name != obj.name && e.disabled==false && e.type=="checkbox") 
				{
					e.checked = obj.checked;
				}				
			} 			
		}
		/////
		function checkedItem()
		{
			var Item = 0;
			for (var i=0;i<document.forms[0].elements.length;i++) 
			{ 
				var e = document.forms[0].elements[i]; 
				if (e.name == "ItemUID" && e.disabled==false && e.type=="checkbox") 
				{
					if(e.checked)
					{
						Item = Item + 1
					}
				}				
			} 
			return Item;
		}
		//////
		function batchAction()
		{
			if(checkedItem()==0)
			{
				alert("Sp_CMSϵͳ��ʾ:\r\n\n��ѡ��Ҫ������������Ϣ!");
				return false;
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?ModelID=<%=ModelID%>&action=<%if action="Recycle" then:response.Write "BatchDel":else:response.Write "batchRecycle":end if%>" name="ModelItemList" onSubmit="javascript:return batchAction();" method="post">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)"><%=ModelName%>����</A></LI>
	  <%if ParentModelID<>0 then%>
	  <LI><A href="model_list.asp?ModelID=<%=ParentModelID%>">����</A></LI>
	  <%end if%>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<div class="divpadding">
		��Ϣ������
	  <%
	  ''''�Ƿ���ʾ����
	  if config_isShowCity=1 then
	  %>
		<select name="txt_CityID" class="select">
		<%searchdictionaryDropDownList CityCategoryArray,0,cityID%>
		</select>		
	  <%
	  '''''Ϊ������һ����
	  else	    
	  %>		
	  <input type="hidden" name="txt_CityID" value="1">
	  <%end if%>
		<select name="sel_category" class="select">
		<%
		''��������˵�����
		SearchDropDownList ModelCategoryArray,0,categoryID
		%>
		</select>
		&nbsp;
		<select name="txt_searchField">
		<option value="title" <%if searchField="title" then response.Write "selected"%>>����</option>
		<%
		set fieldrs = con.query("select fieldname,FieldDescription from Sp_ModelField where IsSearchForm=1 and ModelID="&ModelID&"")
		if fieldrs.recordcount<>0 then
		do while not fieldrs.eof
			response.Write "<option value='"&fieldrs("fieldname")&"'"
			if fieldrs("fieldname")=searchField then response.Write " selected"
			response.Write ">"&fieldrs("FieldDescription")&"</option>"
		fieldrs.movenext
		loop
		end if
		%>		
		</select>
		<input name="txt_keyword" type="text" class="input" value="<%=keyword%>">
		<input name="btnserch" type="button" onClick="javascript:ModelSearch();" value="������Ϣ" class="button">
		</div>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="2%" align=left><input type="checkbox" name="allCheck" onclick="javascript:SelectALLCheck(this);"></TH>
			    <TH align=left>����</TH>
			    <TH align=left>����</TH>
			    <TH width="8%" align=left>������</TH>
			    <TH width="10%" align=left>����ʱ��</TH>
			    <TH width="4%" align=center></TH>
				<TH width="8%" align=left></TH>
		      </TR>
			  <%
				Dim total_record   	'�ܼ�¼��
				Dim current_page	'��ǰҳ��
				Dim PCount			'ѭ��ҳ��ʾ��Ŀ
				Dim pagesize		'ÿҳ��ʾ��¼��
				Dim showpageNum		'�Ƿ���ʾ����ѭ��ҳ
				Dim strWhere		'��ѯ����
				strWhere = " D.ID=MT.categoryID and MT.ModelID = "&ModelID&""
				'response.Write strWhere
				'response.End()
				''''��ѯ����
				if action = "Recycle" then 
					strWhere = strWhere &" and MT.IsDel=1"
				else
					strWhere = strWhere &" and MT.IsDel=0"
				end if
				''''���ɸѡ
				if categoryID<>0 then strWhere = strWhere &"  and D.Pathstr like '"&categoryID&"%'"
				''''�ؼ���ɸѡ
				'if keyword<>"" then strWhere = strWhere &"  and MT.Title like '%"&keyword&"%'"
				''''�ֶ�˧ѡ  �ؼ���ɸѡ
				if searchField<>"" and keyword<>"" then strWhere = strWhere &"  and MT."&searchField&" like '%"&keyword&"%'"
				'''''''''''''''''''''''''''''''''''''''''''''
				''�ù�����ҪΪ����ģ�ͷ���
				''���磺�̼�--->��Ʒ��Ϣ
				'''''''''''''''''''''''''''''''''''''''''''''
				if ParentModelID<>0 then strWhere = strWhere &"  and MT.SubModelID="&ParentModelID&""
				if ModelItemID<>0 then strWhere = strWhere &"  and MT.SubModelItemID="&ModelItemID&""
				'''''''''''''''''''''''''''''''''''''''''''''
				''���Ϲ��ܽ���
				'''''''''''''''''''''''''''''''''''''''''''''
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from "&ModelTableName&" MT, Sp_dictionary D where "&strWhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = Con.QueryPageByNum("MT.ID,MT.IndexID,MT.IsJumpUrl,MT.infoType,MT.title,MT.username,MT.IsCommand,MT.Istop,MT.IsHot,MT.LastUpdateTime,MT.CommentCount,D.Pathstr",""&ModelTableName&" MT, Sp_dictionary D", ""&strWhere&"", "MT.ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemUID" value="<%=rs("id")%>"></TD>
				<TD><%Listlevelmenu ModelID,ModelCategoryArray,rs("Pathstr")%></TD>
				<TD><a href="Model_Edit.asp?ModelID=<%=ModelID%>&id=<%=rs("id")%>"><%=rs("title")%></a>&nbsp;<%if rs("infoType")=1 then response.Write "<span class='huitext'>ͼ</span>"%>&nbsp;<%if rs("IsJumpUrl")=1 then response.Write "<span class='huitext'>ת</span>"%>&nbsp;<% if rs("IsCommand")=1 then response.Write "<span class='redtext'>��</span>"%>&nbsp;<% if rs("istop")=1 then response.Write "<span class='redtext'>��</span>"%>&nbsp;<% if rs("Ishot")=1 then response.Write "<span class='redtext'>��</span>"%></TD>
				<TD><%=rs("username")%></TD>
				<TD><%=Year(rs("LastUpdateTime"))%>-<%=month(rs("LastUpdateTime"))%>-<%=day(rs("LastUpdateTime"))%></TD>
				<TD align="center">
				<!--���۹���//-->
				<%
				if IsAllowComment=1 then
					response.Write "<a title='�Ѿ������� "&rs("CommentCount")&" ��' href='Model_comment_list.asp?modelID="&ModelID&"&ModelItemID="&rs("id")&"'>"
					if rs("CommentCount")=0 then 
					response.Write "<img src='images/comment_2.gif' align='absmiddle' border='0'>"
					else
					response.Write "<img src='images/comment_1.gif' align='absmiddle' border='0'>"
					end if
					response.Write "</a>"
				end if
				%>
				</TD>
				<TD align="center"><a href="Model_Edit.asp?ParentModelID=<%=ParentModelID%>&ModelItemID=<%=ModelItemID%>&ModelID=<%=ModelID%>&id=<%=rs("id")%>">�޸�</a>&nbsp;<%if action="Recycle" then response.Write "<a href='?ModelID="&ModelID&"&action=restore&ID="&rs("id")&"'>��ԭ</a>"%>&nbsp;<a href="javascript:if(confirm('ȷ��ɾ��?ɾ������Ϣ���ɻָ�!')) location.href='?ModelID=<%=ModelID%>&action=<%if action="Recycle" then:response.Write "remove":else:response.Write "del":end if%>&id=<%=rs("id")%>';">ɾ��</a></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<%if action="Recycle" then%>
			<input name="btnrestore" value="����ɾ��" class="button" type="submit">		
			<%else%>
			<input name="btndel" value="��������վ" class="button" type="submit">
			<%end if%>
		</div>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,true,flase%>
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<%
	end if
%>
