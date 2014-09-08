<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim email:email = VerificationUrlParam("email","string","")
	if action<>"" and action="send" then 
		email = F.ReplaceFormText("txt_Email")
		dim title:title = F.ReplaceFormText("txt_title")
		dim content:content = F.ReplaceFormText("txt_content")
		Dim emailarray:emailarray = split(email,vbCrlf)
		Set msg = Server.CreateObject("JMail.Message") 
		for i=0 to ubound(emailarray) step 1
			msg.silent = true 
			msg.Logging = true 
			msg.Charset = "gb2312" 
			msg.ContentType = "text/html" 
			msg.ContentTransferEncoding = "base64" 
			msg.ISOEncodeHeaders = false 
			msg.MailServerUserName = config_EmailAccountUserName '邮箱用户名 
			msg.MailServerPassword = config_EmailAccountUserPWD '邮箱密码 
			msg.From = config_EmailAccount '发送邮件的地址 
			msg.FromName = emailarray(i) 
			msg.AddRecipient(emailarray(i)) '收件人邮箱地址 
			msg.Subject = title
			msg.Body=content 
			msg.Send (config_MailSMtp) 'stmp服务器 
		next
		msg.close 
		set msg = nothing 
		Alert "发送成功",""
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/HtmlEditor.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_Email").value=="")
			{
				alert("Sp_CMS提示\r\n\n邮件地址必须填写!");
				return false;			
			}			
			if(document.getElementById("txt_title").value=="")
			{
				alert("Sp_CMS提示\r\n\n邮件标题必须填写!");
				return false;			
			}
			if(document.getElementById("txt_content").value=="")
			{
				alert("Sp_CMS提示\r\n\n邮件内容必须填写!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=send&email=<%=email%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">群发邮件</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="13%">邮件地址：</TD>
				<TD>
				<textarea name="txt_Email" class="input" cols="60" rows="10"><%=email%></textarea>
				<span class="huitext">&nbsp;必填(一行一个email地址)</span></TD>
			  </TR>
			  
			  <TR class="content-td1">
				<TD>邮件标题：</TD>
				<TD><input name="txt_title" type="text" class="input" size="60">
				<span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>邮件内容：</TD>
				<TD><textarea name="txt_content" cols="60" rows="10" class="input"></textarea>
				<span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="发送邮件" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
