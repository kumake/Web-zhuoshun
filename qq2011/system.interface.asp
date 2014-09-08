<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">网站接口</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<form action="?action=change" method="post" onSubmit="javascript:return checkform();">
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH align=left colSpan="2">&nbsp;</TH>
			  </TR>
			  <TR class="content-td1">
				<TD width="9%">论坛接口：</TD>
				<TD width="91%"><textarea name="txt_username" cols="80" rows="2" class="input"></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>博客接口：</TD>
				<TD><textarea name="txt_username" cols="80" rows="2" class="input"></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>备注字段：</TD>
				<TD><textarea name="txt_username" cols="80" rows="2" class="input"></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD colspan="2"><input type="button" onClick="javascript:alert('功能正在开发中....');" name="btnchangePwd" class="button" value="接口设置"></TD>
			  </TR>
			</TABLE>
		</form>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
