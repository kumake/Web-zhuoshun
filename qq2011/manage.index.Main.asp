<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Md5.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	''''Ȩ����֤'''
	''''CheckUserAuthority "",""
	''''Ȩ����֤'''
	''response.Write MD5("admin")
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="change" then 
		dim userpwd:userpwd = F.ReplaceFormText("txt_userpwd")
		dim olduserpwd:olduserpwd = F.ReplaceFormText("txt_olduserpwd")
		'response.Write rolepower
		'response.End()
		set rs = con.query("select * from Sp_Manage where Username='"&CK("username")&"' and userpwd='"&Md5(olduserpwd)&"'")
		if rs.recordcount<>0 then
			Dim sql:sql = "update Sp_Manage set userpwd='"&md5(userpwd)&"' where iD="&CK("userid")&""
			Con.execute(sql)
			Alert "�޸�����ɹ�","manage.index.main.asp"
		else
			Alert "�������������",""
		end if
	end if
	'response.Write CK("usermenu")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>SPCMS��վasp���̨����Ա����̨--�Ͼ�һ������ </TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_username").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�û���������д!");
				return false;			
			}
			if(document.getElementById("txt_olduserpwd").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�����������д!");
				return false;			
			}			
			if(document.getElementById("txt_userpwd").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n���������д!");
				return false;			
			}			
			if(document.getElementById("txt_reuserpwd").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�ظ����������д!");
				return false;			
			}		
				
			if(document.getElementById("txt_reuserpwd").value!=document.getElementById("txt_userpwd").value)
			{
				alert("Sp_CMS��ʾ\r\n\n���������ظ����벻һ��!");
				return false;			
			}			
			return true;
		}
	</script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onFocus="this.blur()" onClick="selectTag('tagContent0',this)" href="javascript:void(0)">��̨��ҳ</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
		<form action="?action=change" method="post" onSubmit="javascript:return checkform();">
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH align=left colSpan="2">���������޸� </TH>
			  </TR>
			  <TR class="content-td1">
				<TD width="9%">�û�����</TD>
				<TD width="91%"><input type="text" name="txt_username" value="<%=CK("username")%>" class="input" disabled></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�����룺</TD>
				<TD><input type="password" name="txt_olduserpwd" class="input"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�����룺</TD>
				<TD><input type="password" name="txt_userpwd" class="input"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�����룺</TD>
				<TD><input type="password" name="txt_reuserpwd" class="input"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD colspan="2"><input type="submit" name="btnchangePwd" class="button" value="�޸�����"></TD>
			  </TR>
			</TABLE>
		</form>
		</DIV>
		<BR>
		<DIV>
			<TABLE cellSpacing=1 width="100%">
			<TR class=content-td1>
				<TH scope=col align=left colspan="2">�����Ŷ� </TH>
			</TR>
			<TR class=content-td1>
				<TD width="9%">�����д</TD>
				<TD>�Ͼ�һ������ <span id="copyright" class="huitext"> </TD>
			</TR>
			</TABLE>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
