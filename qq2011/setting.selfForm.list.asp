<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	response.Expires=0
	Dim action:action = VerificationUrlParam("action","string","")
	Dim FormID:FormID = VerificationUrlParam("FormID","int","0")
	if action<>"" and action="del" then
		''ȡģ�Ͷ�Ӧ����ϸ��Ϣ
		Dim FormArray:FormArray = Con.QueryRow("select ID,FormName,FormTable from Sp_Form where ID="&FormID&"",0)
		Dim FormTable:FormTable = FormArray(2)
		'response.Write ModelTable
		'response.End()
		''ɾ����Ϣ
		Con.Execute("delete * from Sp_Form where ID="&FormID&"")
		Con.Execute("delete * from Sp_FormField where FormID="&FormID&"")
		''ɾ�������ṹ		
		Con.Execute("drop table "&FormTable&"")
		ALert "ɾ���Զ�����ɹ�","setting.SelfForm.List.asp"
	end if	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript" src="Scripts/common_management.js"></script>
	<script language="javascript" type="text/javascript">
		function showopen(url)
		{
			window.showModalDialog(url,"","dialogTop:250px,dialogLeft:500px;dialogWidth:500px,DialogHeight:250px,status:no'");
		}
	</script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�Զ����</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id="tagContent0">
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH width="5%" align="left"><input type="checkbox" name="checkbox" value="checkbox"></TH>
			    <TH width="20%" align="left">������</TH>
			    <TH width="10%" align="left">��������</TH>
			    <TH width="40%" align="left">������</TH>
			    <TH width="25%" align="left"></TH>
			  </TR>
			  <%
				Dim total_record   	'�ܼ�¼��
				Dim current_page	'��ǰҳ��
				Dim PCount			'ѭ��ҳ��ʾ��Ŀ
				Dim pagesize		'ÿҳ��ʾ��¼��
				Dim showpageNum		'�Ƿ���ʾ����ѭ��ҳ
				if total = 0 then 
					Dim Tarray:Tarray = Con.QueryData("select count(*) as total from Sp_Form")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6
				pagesize = 11
				showpageNum = true
				
				'ѡ���ֶΣ����������ؼ��֣��Ƿ����򣬿�ʼ�У�1�ǵ�һ�У����������
				Dim sql:sql = Con.QueryPageByNum("*","Sp_Form", "", "ID", false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = Con.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
			  %>
			  <TR class="content-td1">
				<TD>&nbsp;</TD>
				<TD><%=rs("formname")%></TD>
				<TD><%=rs("FormTable")%></TD>
				<TD><%=rs("memo")%></TD>
				<TD><a href="javascript:showopen('setting.SelfForm.show.asp?action=html&FormID=<%=rs("ID")%>&rnd=<%=Int((99999 - 11111 + 1) * Rnd + 11111)%>');">Html����</a>&nbsp;<a href="javascript:showopen('setting.SelfForm.show.asp?action=javascript&FormID=<%=rs("ID")%>&rnd=<%=Rnd(10000)%>');">Js����</a>&nbsp;<a href="setting.selfForm.Field.List.asp?FormID=<%=rs("ID")%>">�ֶ�</a>&nbsp;<a href="setting.selfForm.edit.asp?FormID=<%=rs("ID")%>">�޸�</a>&nbsp;<a href="javascript:if(confirm('Sp_CMSϵͳ��ʾ:\r\n\nȷ��ɾ��?ɾ���󲻿ɻָ�!'))location.href='?action=del&FormID=<%=rs("ID")%>';">ɾ��</a></TD>
			  </TR>
			  <%
			   rs.movenext
			   loop
			   end if
			  %>
			</TABLE>
		</DIV>
		<div class="divpadding">
			<%PageIndexUrl total_record,current_page,PCount,pagesize,true,true,flase%>
		</div>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
