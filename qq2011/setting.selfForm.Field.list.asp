<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim FormID:FormID = VerificationUrlParam("FormID","int","")
	Dim ID:ID = VerificationUrlParam("ID","int","")
	if action<>"" and action="del" then
		Dim Fieldarray:Fieldarray = Con.QueryRow("select M.formTable,F.fieldname from Sp_form M,Sp_formField F where F.formID=M.ID and F.ID="&ID&"",0)
		if ubound(Fieldarray)>0 then
			'ɾ�����¼
			Con.Execute("delete * from Sp_formField where ID="&ID&"")
			'ɾ�������ṹ 
			Con.Execute("alter table "&Fieldarray(0)&" DROP COLUMN "&Fieldarray(1)&"")
			Alert "ɾ���ֶγɹ�","setting.selfform.Field.list.asp?FormID="&FormID&""
		else
			Alert "���ݲ�ѯʧ��",""
		end if		
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�Զ�����ֶ��б�</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<div class="divpadding">
		��������
		<select name="ModelID" style="width:100px;" class="select" onChange="javascript:if(this.value!='') location.href='?FormID='+this.value;">
		<option value="">ѡ���</option>
		<%
			set Rs1 = Con.Query("select id,formname from SP_Form")
			if rs1.recordcount<>0 then
			do while not rs1.eof
			response.Write "<option value="&rs1("id")&""
			if FormID<>"" then
			if Cint(FormID)=Cint(rs1("ID")) then 
			response.Write " selected"
			end if
			end if
			response.Write ">"&rs1("formname")&"</option>"
			rs1.movenext
			loop
			end if
		%>
		</select>
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
				Dim strwhere:strwhere = "F.FormID=M.ID"
				if FormID<>"" then strwhere = strwhere & " and M.ID="&FormID&"" 
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Form M,Sp_FormField F where "&strwhere&"")
				    total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = Con.QueryPageByNum("F.*,M.formname","Sp_Form M,Sp_FormField F", ""&strwhere&"", "F.ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD>&nbsp;</TD>
				<TD><%=rs("formName")%></TD>
				<TD><%=rs("fieldname")%></TD>
				<TD><%=rs("fieldtype")%></TD>
				<TD><%=rs("fieldUI")%></TD>
				<TD><%=rs("FieldDescription")%></TD>
				<TD><a href="setting.selfform.Field.Edit.asp?FormID=<%=rs("formID")%>&ID=<%=rs("ID")%>">�޸�</a>&nbsp;<a href="javascript:if(confirm('Sp_CMSϵͳ��ʾ:\r\n\nȷ��ɾ��?ɾ���󲻿ɻָ�!'))location.href='?action=del&FormID=<%=rs("formID")%>&ID=<%=rs("ID")%>';">ɾ��</a></TD>
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
