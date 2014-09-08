<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Md5.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	''''权限验证'''
	Dim filename:filename = VerificationUrlParam("filename","string","")
	Dim action:action = VerificationUrlParam("action","string","")
	dim templatePath:templatePath = server.MapPath("/Template/"&filename&"")
	Dim Templatestream
	'response.Write templatePath
	if action<>"" and action="save" then 
		Templatestream = F.ReplaceFormText("txtcontent")
		filename = F.ReplaceFormText("txtfilename")
		templatePath = server.MapPath("/Template/"&filename&"")
		F.WriteTextFile templatePath,Templatestream
		Alert "编辑成功","setting.Model.template.asp"
	elseif action<>"" and action="del" then 
		F.DeleteFile templatePath 
		Alert "删除成功","setting.Model.template.asp"
	end if
	if filename<>"" then Templatestream = F.ReadTextFile(templatePath)
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
			if(document.getElementById("txtcontent").value=="")
			{
				alert("Sp_CMS提示\r\n\n必须填写!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">模板编辑</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
		<form action="?action=save&filename=<%=filename%>" method="post" onSubmit="javascript:return checkform();">
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td2">
				<TH align=left colSpan="2">&nbsp;</TH>
			  </TR>
			  <TR class="content-td1">
				<TD width="9%">模板名称：</TD>
				<TD width="91%"><input type="text" class="input" name="txtfilename" value="<%=filename%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="9%">模板内容：</TD>
				<TD width="91%"><textarea cols="150" rows="30" name="txtcontent"><%=Templatestream%></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD colspan="2"><input type="submit" name="btnchangePwd" class="button" value="修改模板"></TD>
			  </TR>
			</TABLE>
		</form>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
