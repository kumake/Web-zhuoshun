<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	'action
	Dim action:action = VerificationUrlParam("action","string","")
	'��Ϣ��ʶID
	Dim ItemID:ItemID = VerificationUrlParam("ID","int","0")
	'����ɾ��ʱ�����Ϣ
	Dim BatchItemID
	'ģ��ID
	Dim FormID:FormID = VerificationUrlParam("FormID","int","0")
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
	'''''''''''''''''''''''
	if FormID=0 then
		response.Write "ģ�ͽ��ͳ��ִ���!"
	else
		'ȡ��Ӧģ�͵����ݴ�����ݱ�
		Modelarray = Con.QueryRow("select * from Sp_Form where Id="&FormID&"",0) 
		ModelName = Modelarray(1)		
		ModelTableName = Modelarray(2)		
		ModelCategoryID = Modelarray(3)
		'''ISdel=1ɾ����Ϣ
		if action<>"" and action="remove" then
			Con.Execute("delete * from "&ModelTableName&" where ID="&ItemID&"")
			Alert "ɾ����Ϣ�ɹ�","pulg.selfform.info.asp?FormID="&FormID&""
		elseif action="batchDel" then
		''''����ɾ����Ϣ
			BatchItemID = F.ReplaceFormText("ItemUID")
			Con.Execute("delete * from "&ModelTableName&" where ID in ("&BatchItemID&")")
			Alert "ɾ����Ϣ�ɹ�","pulg.selfform.info.asp?FormID="&FormID&""
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
		function ModelSearch()
		{
			var obj2 = document.getElementById("txt_keyword").value;
			var url = "?FormID=<%=FormID%>";
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
		function batchReMove()
		{
			if(checkedItem()==0)
			{
				alert("Sp_CMSϵͳ��ʾ:\r\n\n��ѡ��Ҫ������������Ϣ!");
			}
			else
			{
				document.forms[0].action = "?FormID=<%=FormID%>&action=batchDel";
				document.forms[0].submit();
			}		
		}
	</script>
</HEAD>
<BODY>
<form action="?action=batchDel&FormID=<%=FormID%>" method="post">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onFocus="this.blur()" onClick="selectTag('tagContent0',this)" href="javascript:void(0)"><%=ModelName%>����</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<div class="divpadding">
		��Ϣ������		
		<input name="txt_keyword" type="text" class="input" value="<%=keyword%>">
		<input name="btnserch" type="button" onClick="javascript:ModelSearch();" value="������Ϣ" class="button">
		</div>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="4%" align=left><input type="checkbox" name="allCheck" onClick="javascript:SelectALLCheck(this);"></TH>
				<TH width="62%" align=left>����</TH>
			    <TH width="12%" align=left>����ʱ��</TH>
			    <TH width="10%" align=left></TH>
		      </TR>
			  <%
				Dim total_record   	'�ܼ�¼��
				Dim current_page	'��ǰҳ��
				Dim PCount			'ѭ��ҳ��ʾ��Ŀ
				Dim pagesize		'ÿҳ��ʾ��¼��
				Dim showpageNum		'�Ƿ���ʾ����ѭ��ҳ
				Dim strWhere		'��ѯ����
				
				strWhere = "1=1"
				''''��ѯ����
				if keyword<>"" then strWhere = strWhere &" and title like '%"&keyword&"%'"

				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from "&ModelTableName&" where "&strWhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = Con.QueryPageByNum("*",""&ModelTableName&"", ""&strWhere&"", "ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD><input type="checkbox" name="ItemUID" value="<%=rs("id")%>"></TD>
				<TD><a href="plug.selfform.info.view.asp?FormID=<%=FormID%>&id=<%=rs("id")%>"><%=rs("Field_Tile")%></a><a href="plug.selfform.info.view.asp?FormID=<%=FormID%>&id=<%=rs("id")%>"></a></TD>
				<TD><%=Year(rs("PostTime"))%>-<%=month(rs("PostTime"))%>-<%=day(rs("PostTime"))%></TD>
				<TD align="center"><a href="javascript:if(confirm('ȷ��ɾ��?ɾ������Ϣ���ɻָ�!')) location.href='?FormID=<%=FormID%>&action=remove&id=<%=rs("id")%>';">ɾ��</a></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<input name="btnsearch" onClick="javascript:batchReMove();" value="����ɾ��" class="button" type="button">
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
