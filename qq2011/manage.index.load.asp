<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../Sp_inc/class_Md5.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<%
	'''''1, 2, 3, 4, 5, 6, 7, 8, 9, 10,1, 2, 3, 4, 5
	'''''定义函数去除重复的权限
	function splitUserpower(userpower)
		Dim temparray
		Dim tempstr
		temparray = split(trim(userpower),",")
		for i=0 to ubound(temparray)	
			if instr(tempstr,temparray(i))>0 then 
			else
				if tempstr<>"" then
				tempstr = tempstr &","&temparray(i)
				else
				tempstr = temparray(i)
				end if
			end if
		next
		splitUserpower = tempstr
	end function
	'''''取详细权限
	set rs1 = Con.Query("select RolePower from Sp_ManageRole where id in ("&CK("userrole")&")")
	if rs1.recordcount<>0 then
		do while not rs1.eof
			if userpower<>"" then
			userpower = userpower & "," &rs1("RolePower")
			else
			userpower = rs1("RolePower")
			end if
		rs1.movenext
		loop
	end if
	CK("userpower") = splitUserpower(userpower)
	'''''取详细菜单
	set rs2 = Con.Query("select ModelCode from Sp_ManagePower where id in ("&CK("userpower")&") and ParentID=0")
	if rs2.recordcount<>0 then
		do while not rs2.eof
			if userpower<>"" then
			usermenu = usermenu & "," &rs2("ModelCode")
			else
			usermenu = rs2("ModelCode")
			end if
		rs2.movenext
		loop
	end if
	CK("usermenu") = splitUserpower(usermenu)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>加载权限：5秒钟转入</title>
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<script language="javascript" type="text/javascript">
		var i=0;
		function test()
		{
			i++;
			document.getElementById("loading").style.width = i + "%";
			if(i<100)
			{
				setTimeout("test()",19);
			}
			else
			{
				location.href= "manage.index.asp";
			}
		}
		setTimeout("test()",19);
	</script>
</head>

<body>
<table width="100%" height="0"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="110">&nbsp;</td>
  </tr>
  <tr>
    <td height="60" align="center">
	<div id="load">
	<div id="loading"></div>
	</div>
	</td>
  </tr>
</table>

</body>
</html>
