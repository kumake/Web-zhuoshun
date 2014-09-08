<!--#include file="Sp_inc/class_Application.asp"-->
<!--#include file="Sp_inc/class_Session.asp"-->
<!--#include file="Sp_inc/class_Cookie.asp"-->
<!--#include file="Sp_inc/class_Function.asp"-->
<!--#include file="Sp_inc/class_Conn.asp"-->
<!--#include file="Sp_inc/User_Function.asp"-->
<%
	Dim UICon
	Set UICon = New HotCMS_Class_Conn
	'链接Access数据库
	UICon.Open "newpic/new/2009pic/2/3/##SP_cms#.mdb", DTAccess
	'连接sql数据库
	'UIcon.Open "SqlLocalName$SqlDatabaseName$SqlUsername$SqlPassword",DTMSSQL
	'显示调试信息\
	'An.Clear()
	'Response.Write F.trace()
%>
