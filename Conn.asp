<!--#include file="Sp_inc/class_Application.asp"-->
<!--#include file="Sp_inc/class_Session.asp"-->
<!--#include file="Sp_inc/class_Cookie.asp"-->
<!--#include file="Sp_inc/class_Function.asp"-->
<!--#include file="Sp_inc/class_Conn.asp"-->
<!--#include file="Sp_inc/User_Function.asp"-->
<%
	Dim UICon
	Set UICon = New HotCMS_Class_Conn
	'����Access���ݿ�
	UICon.Open "newpic/new/2009pic/2/3/##SP_cms#.mdb", DTAccess
	'����sql���ݿ�
	'UIcon.Open "SqlLocalName$SqlDatabaseName$SqlUsername$SqlPassword",DTMSSQL
	'��ʾ������Ϣ\
	'An.Clear()
	'Response.Write F.trace()
%>
