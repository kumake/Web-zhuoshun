<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Md5.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	''''权限验证'''
	CheckUserAuthority "",""
	''''权限验证'''
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="do" then 
		dim sql:sql = F.ReplaceFormText("txt_sql")
		'On Error resume Next
		Con.Query(sql)
		If Err.number<>0 Then
			Alert "执行错误,请检查sql语句正确与否!",""
		else
			Alert "执行成功",""
		end if
	end if
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_sql").value=="")
			{
				alert("Sp_CMS提示\r\n\n请输入要执行的sql!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">sql高级管理</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
		<form action="?action=do" method="post" onSubmit="javascript:return checkform();">
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="9%">sql语句：</TD>
				<TD width="91%"><textarea name="txt_sql" cols="80" rows="5" class="input" type="text"></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD colspan="2"><input type="submit" name="btnchangePwd" class="button" value="执行sql"></TD>
			  </TR>
			</TABLE>
		</form>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
