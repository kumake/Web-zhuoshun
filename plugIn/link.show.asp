<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="site.css" rel="stylesheet" type="text/css">
<!--#include file="../conn.asp"-->
<div id="FriendLink">
<%
	'''��ϲ�ѯsql
	Dim sql:sql = "select urltype,sitename,siteurl,sitelogourl from Sp_FriendLink where ISdisplay=1"
	'''ִ��sqlȡ������
	set rs = UICon.Query(sql)
	if rs.recordcount<>0 then
	do while not rs.eof
		'''ѭ��ȡ����Ϣ
		response.Write "<li><a href='"&rs("siteurl")&"'>"
		if rs("urltype")=1 then
		response.Write "<img align='absmiddle' src='"&rs("sitelogourl")&"' border='0'>"
		else
		response.Write rs("sitename")
		end if
		response.Write "</a></li>"
	rs.movenext
	loop
	rs.close
	set rs=nothing
	end if			
%>
</div>
