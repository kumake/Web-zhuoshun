<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	dim a,b,c,title
	a=request("page")
	b=request("wd")
	c=request("long")
	if a="" then a="1"
	if b<>"" then
		url = "http://www.baidu.com/s?lm="&c&"&si=&rn=20&tn=baiduadv&ie=gb2312&ct=0&wd=site%3A"&b&"&pn="&a&"&cl=3"
	else
		TheBody =""
	end if
	
	ComStr = GetPage(url)
	Set Re=new RegExp
	Re.Global = True
	TheBody = GetContent(Comstr,"<table border=""0"" cellpadding=""0"" cellspacing=""0""><tr>","<form name=f2 action=""/s"">",0)
	total = GetContent(Comstr,"百度一下，找到相关网页","篇，用时",0)
	Re.Pattern="<font size=-1(.+?)点击此处</a>查看全部搜索结果</font><br>"
	TheBody=Re.Replace(TheBody,"")
	Re.Pattern="- <a href=""http://cache.baidu.com/c(.+?)"" target=""_blank"" class=m>百度快照</a>"
	TheBody=Re.Replace(TheBody,"")
	Re.Pattern="<a href=s\?lm=(.+?)&si=&rn=20&tn=baiduadv&ie=gb2312&ct=0&wd=site%3A(.+?)&pn=(.+?)&ver=0&cl=3&uim=0&usm=0>"
	TheBody=Re.Replace(TheBody,"<a href=?wd=$2&long=$1&page=$3>")
	Set Re=Nothing
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">百度收录查询</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
	  <DIV>
		<form name="f" method="post" action="manage.index.baidu.asp">网站域名：
			<input type=text name="wd" class="input" size="21" maxlength="100"  value="www.8-du.net">
			<select name="long" style="background-color:#CCCCCC;">
				<option selected="selected" value="1">请选择要查询近几天的收录量</option>
				<option value=1>查询昨日收录情况</option>
				<option value=7>查询最近1星期收录情况</option>
				<option value=30>查询最近1月收录情况</option>
				<option value=360>查询最近1年收录情况</option>
				<option value=0>查询总的（所有日期）收录情况</option>
			</select>
			<input type="submit" class="buttonlen" value="立即查询百度近日收录情况！">
			<br>
		</form>
	 <br>
		<table border="0" cellpadding="0" cellspacing="0">
			<%=TheBody%>
		</table>
	  </DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
