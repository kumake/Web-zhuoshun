<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="../Sp_inc/class_Template.asp"-->
<%
	Dim Errormsg
	Dim msgArray
	if config_IsIndexStatic = 1 then
		''''开始执行生成静态页面操作
		if config_IndexTemplate<>"" then
			''调用程序生成
			Dim template
			Set template = New Cls_Tempalte
			dim filestream:filestream = template.Parser(config_IndexTemplate)
			'''
			F.WriteTextFile "/index.html",filestream
			Errormsg = "生成成功!"
		else
			Errormsg = "你没有设置静态首页的模板!"
		end if
	else
		Errormsg = "你没有设置生成静态页面.请现在全局设置中设置可以生成静态首页!"
	end if
%>
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">生成首页</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
	  <DIV>
	  	<%
		  if Errormsg<>"" then
		  	msgArray = split(Errormsg,"$")
			'''显示信息
			for i=0 to ubound(msgArray)
				response.Write (i+1)&"."& msgArray(i) &"<br>"
			next
		  end if
		%>
	  </DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
