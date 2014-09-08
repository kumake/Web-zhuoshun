<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
		Dim Labname:Labname = F.ReplaceFormText("txt_name")
		Dim Labcontent:Labcontent = F.ReplaceFormText("txt_content")
		'''''''增加记录
		Con.execute("insert into sp_Label (name,code) values ('"&Labname&"','"&Labcontent&"')")
		Alert "增加标签成功","setting.Model.label.asp"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/calender.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_name").value=="")
			{
				alert("Sp_CMS提示\r\n\n名称必须填写!");
				return false;			
			}
			if(document.getElementById("txt_content").value=="")
			{
				alert("Sp_CMS提示\r\n\n必须填写!");
				return false;			
			}

			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=save" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">增加标签</A></LI>
	  <LI><A href="setting.Model.label.asp">标签管理</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="10%">标签名称：</TD>
				<TD><input name="txt_name" type="text" class="input"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>内容：</TD>
				<TD><textarea name="txt_content" cols="150" rows="20" class="input"></textarea>
				<span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="增加标签" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
