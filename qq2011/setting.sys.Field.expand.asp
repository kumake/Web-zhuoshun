<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim tableID:tableID = VerificationUrlParam("tableID","int","")
	Dim ID:ID = VerificationUrlParam("ID","int","")
	
	if action<>"" and action="del" then
		Dim Fieldarray:Fieldarray = Con.QueryRow("select T.tablename,F.fieldname from Sp_table T,Sp_tableField F where T.id=F.tableID and F.ID="&ID&"",0)
		'ɾ�����¼
		Con.Execute("delete * from Sp_TableField where ID="&ID&"")
		'ɾ�������ṹ 
		Con.Execute("alter table "&Fieldarray(0)&" DROP COLUMN "&Fieldarray(1)&"")
		Alert "ɾ���ֶγɹ�","setting.sys.Field.expand.asp?tableID="&tableID&""
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">ϵͳ���ֶ���չ</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<div class="divpadding">
		ϵͳ�������
		<select name="GroupID" style="width:100px;" class="select" onChange="javascript:if(this.value!='') location.href='?tableID='+this.value;">
		<option value="">ѡ��ϵͳ��</option>
		<%
			set Rs1 = Con.Query("select id,tablename from Sp_Table")
			if rs1.recordcount<>0 then
			do while not rs1.eof
			response.Write "<option value="&rs1("id")&""
			if tableid<>"" then
			if Cint(tableid)=Cint(rs1("ID")) then 
			response.Write " selected"
			end if
			end if
			response.Write ">"&rs1("tablename")&"</option>"
			rs1.movenext
			loop
			end if
		%>
		</select>&nbsp;&nbsp;&nbsp;<a href="setting.sys.Field.expand.add.asp">������ֶ�</a>
		</div>
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH align="left">������</TH>
			    <TH align="left">�ֶ�����</TH>
			    <TH align="left">�ֶ�����</TH>
			    <TH align="left">ҳ����ʾ��ʽ</TH>
			    <TH align="left">�ֶ�ҳ������</TH>
			    <TH align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'�ܼ�¼��
				Dim current_page	'��ǰҳ��
				Dim PCount			'ѭ��ҳ��ʾ��Ŀ
				Dim pagesize		'ÿҳ��ʾ��¼��
				Dim showpageNum		'�Ƿ���ʾ����ѭ��ҳ
				Dim strwhere:strwhere = "T.id=F.tableID"
				if tableID<>"" then strwhere = strwhere & " and T.ID="&tableID&"" 
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Table T,Sp_TableField F where "&strwhere&"")
				    total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = Con.QueryPageByNum("F.*,T.tablename","Sp_Table T,Sp_TableField F", ""&strwhere&"", "F.ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD>&nbsp;</TD>
				<TD><%=rs("tablename")%></TD>
				<TD><%=rs("fieldname")%></TD>
				<TD><%=rs("fieldtype")%></TD>
				<TD><%=rs("fieldUI")%></TD>
				<TD><%=rs("FieldDescription")%></TD>
				<TD><a href="setting.sys.Field.expand.Edit.asp?tableid=<%=rs("tableid")%>&ID=<%=rs("ID")%>">�޸�</a>&nbsp;<a href="javascript:if(confirm('Sp_CMSϵͳ��ʾ:\r\n\nȷ��ɾ��?ɾ���󲻿ɻָ�!'))location.href='?action=del&tablename=<%=rs("tablename")%>&ID=<%=rs("ID")%>';">ɾ��</a></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<input name="btnsearch" value="����ɾ��" class="button" type="button">
		</div>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,true,flase%>
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
