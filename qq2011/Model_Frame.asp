<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	'action
	Dim leftUrl:leftUrl = VerificationUrlParam("leftUrl","string","")
	''''����ģ����ϢID
	DIm MainUrl:MainUrl = VerificationUrlParam("MainUrl","string","")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ģ�Ϳ�ܹ���</title>
</head>
<frameset cols="190,*" frameborder="NO" border="0" framespacing="0">
  <frame src="<%=leftUrl%>" id="leftUrl" scrolling="NO" noresize>
  <frame src="<%=MainUrl%>" id="MainUrl" scrolling="yes" name="mainModel">
</frameset>
<noframes><body>
</body></noframes>
</html>
