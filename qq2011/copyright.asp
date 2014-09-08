<!--#include file="../Sp_inc/conn.asp"-->
<%
response.Write encryptCode(request.ServerVariables("HTTP_HOST"))
Dim action:action = VerificationUrlParam("action","string","")
if action<>"" and action="change" then
 	dim Domain:Domain = F.ReplaceFormText("txt_Domain")
	F.WriteTextFile "copyright.inc",Domain
	Alert "设置成功","index.asp"
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
			if(document.getElementById("txt_Domain").value=="")
			{
				alert("Sp_CMS提示\r\n\n验证码必须填写!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">网站版权管理</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
		<form action="?action=change" method="post" onSubmit="javascript:return checkform();">
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH align=left colSpan="2">&nbsp;</TH>
			  </TR>
			  <TR class="content-td1">
				<TD width="14%">验证码写入：</TD>
				<TD width="86%"><input name="txt_Domain" type="text"  class="input" size="80"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD colspan="2"><input type="submit" name="btnchangePwd" class="button" value="写入验证码"></TD>
			  </TR>
			</TABLE>
		</form>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
