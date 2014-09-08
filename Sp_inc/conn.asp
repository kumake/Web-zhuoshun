<!--#include file="class_Cookie.asp"-->
<!--#include file="class_Application.asp"-->
<!--#include file="class_Session.asp"-->
<!--#include file="class_Function.asp"-->
<!--#include file="class_Conn.asp"-->
<!--#include file="User_Function.asp"-->
<%
	Dim Con
	Set Con = New HotCMS_Class_Conn
	'链接Access数据库
	Con.Open "../newpic/new/2009pic/2/3/##SP_cms#.mdb", DTAccess
	'连接sql数据库
	'UIcon.Open "SqlLocalName$SqlDatabaseName$SqlUsername$SqlPassword",DTMSSQL
	'显示调试信息\
	'An.Clear()
	'Response.Write F.trace()
%>
