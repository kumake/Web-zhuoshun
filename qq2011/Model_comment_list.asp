<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	'action
	Dim action:action = VerificationUrlParam("action","string","")
	'ģ��ID
	Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","0")
	'��Ϣ��ʶID
	Dim ModelItemID:ModelItemID = VerificationUrlParam("ModelItemID","int","0")
	'������Ϣ��ʶID
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	'��Ÿ�ģ���������ݱ�
	Dim ModelTableName
	'ģ������
	Dim ModelName
	if action="del" and ItemID<>0 then
		Con.Execute("delete * from Sp_Comment where ID="&ItemID&"")
		Alert "ɾ���ɹ�","Model_comment_list.asp?modelID="&ModelID&"&ModelItemID="&ModelItemID&""
	elseif action="batchdel" then
		Dim BatchrestoreItemID:BatchrestoreItemID = F.ReplaceFormText("ItemUID")
		Con.Execute("delete * from Sp_Comment where ID in ("&BatchrestoreItemID&")")
		'''''''
		WriteLog CK("username"),"action","��������վģ�� "&ModelName&" ����"
		'''''''
		Alert "����ɾ���ɹ�","Model_comment_list.asp?modelID="&ModelID&"&ModelItemID="&ModelItemID&""
	end if
	if ModelID=0 then
		Alter "ģ�ͽ��ͳ��ִ���!",""
	else
		'ȡ��Ӧģ�͵����ݴ�����ݱ�
		Modelarray = Con.QueryRow("select ID,modelname,modeltable from Sp_Model where Id="&ModelID&"",0) 
		ModelName = Modelarray(1)		
		ModelTableName = Modelarray(2)		
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript">
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
<form action="?ModelID=<%=ModelID%>&ModelItemID=<%=ModelItemID%>&action=batchdel" name="ModelItemList" onSubmit="javascript:return batchAction();" method="post">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">���۹���</A></LI>
	  <LI><A href="model_list.asp?ModelID=<%=ModelID%>">����<%=ModelName%></A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="2%" align=left><input type="checkbox" name="allCheck" onclick="javascript:SelectALLCheck(this);"></TH>
			    <TH align=left width="20%" >����</TH>
			    <TH align=left width="10%" >����</TH>
			    <TH align=left width="10%" >����/��ϵ��</TH>
			    <TH align=left width="30%" >��������/��ϵ��ʽ</TH>
			    <TH width="8%" align=left>������</TH>
			    <TH width="10%" align=left>����ʱ��</TH>
			    <TH width="5%" align=left></TH>
		      </TR>
			  <%
				Dim total_record   	'�ܼ�¼��
				Dim current_page	'��ǰҳ��
				Dim PCount			'ѭ��ҳ��ʾ��Ŀ
				Dim pagesize		'ÿҳ��ʾ��¼��
				Dim showpageNum		'�Ƿ���ʾ����ѭ��ҳ
				Dim strWhere		'��ѯ����
				strWhere = " C.ModelID = "&ModelID&" and C.ModelItemID = MT.ID"
				if ModelItemID<>0 then strWhere = strWhere &" and C.ModelItemID="&ModelItemID&" "
				'response.Write strWhere
				'response.End()
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from "&ModelTableName&" MT, Sp_Comment C where "&strWhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = Con.QueryPageByNum("C.ID as ItemID,C.content,C.title as Ctitle,C.username,C.Ip,C.posttime,C.CommentTYpe,Mt.ID,Mt.title",""&ModelTableName&" MT, Sp_Comment C", ""&strWhere&"", "C.ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemUID" value="<%=rs("ItemID")%>"></TD>
				<TD><a href="Model_Edit.asp?ModelID=<%=ModelID%>&id=<%=rs("id")%>"><%=rs("title")%></a></TD>
				<TD><%if rs("CommentTYpe")="C" then:response.Write "����":else:response.Write "��ְӦƸ":end if%></TD>
				<TD><%=rs("Ctitle")%></TD>
				<TD><%=rs("content")%></TD>
				<TD><%=rs("username")%></TD>
				<TD><%=Year(rs("postTime"))%>-<%=month(rs("postTime"))%>-<%=day(rs("postTime"))%></TD>
				<TD align="center"><a href="javascript:if(confirm('ȷ��ɾ��?ɾ������Ϣ���ɻָ�!')) location.href='Model_comment_list.asp?modelID=<%=ModelID%>&ModelItemID=<%=ModelItemID%>&ItemID=<%=rs("ItemID")%>&action=del';">ɾ��</a></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<input name="btnrestore" value="����ɾ��" class="button" type="submit">		
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
