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
	total = GetContent(Comstr,"�ٶ�һ�£��ҵ������ҳ","ƪ����ʱ",0)
	Re.Pattern="<font size=-1(.+?)����˴�</a>�鿴ȫ���������</font><br>"
	TheBody=Re.Replace(TheBody,"")
	Re.Pattern="- <a href=""http://cache.baidu.com/c(.+?)"" target=""_blank"" class=m>�ٶȿ���</a>"
	TheBody=Re.Replace(TheBody,"")
	Re.Pattern="<a href=s\?lm=(.+?)&si=&rn=20&tn=baiduadv&ie=gb2312&ct=0&wd=site%3A(.+?)&pn=(.+?)&ver=0&cl=3&uim=0&usm=0>"
	TheBody=Re.Replace(TheBody,"<a href=?wd=$2&long=$1&page=$3>")
	Set Re=Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�ٶ���¼��ѯ</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
	  <DIV>
		<form name="f" method="post" action="manage.index.baidu.asp">��վ������
			<input type=text name="wd" class="input" size="21" maxlength="100"  value="www.8-du.net">
			<select name="long" style="background-color:#CCCCCC;">
				<option selected="selected" value="1">��ѡ��Ҫ��ѯ���������¼��</option>
				<option value=1>��ѯ������¼���</option>
				<option value=7>��ѯ���1������¼���</option>
				<option value=30>��ѯ���1����¼���</option>
				<option value=360>��ѯ���1����¼���</option>
				<option value=0>��ѯ�ܵģ��������ڣ���¼���</option>
			</select>
			<input type="submit" class="buttonlen" value="������ѯ�ٶȽ�����¼�����">
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
