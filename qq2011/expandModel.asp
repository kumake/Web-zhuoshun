<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
Dim ModelID:ModelID = VerificationUrlParam("ModelID","int","0")
''''
'����model�����չ����
con.execute("insert into Sp_Model select modelname,modeltable,modelCategoryID,"&ModelID&" as ExpModelID,IsDisplay,0 as IsSystem,IsUserAllowAdd from Sp_Model where ID="&ModelID&"")
''
Dim ExpModelID:ExpModelID = 0
if ModelID<>0 then
	Dim ItemArray:ItemArray = Con.QueryRow("select top 1 id from Sp_Model order by id desc",0)
	ExpModelID = ItemArray(0)
end if
''����ԭ���ݵ��ֶ�����

con.execute("insert into Sp_ModelField select "&ExpModelID&" as ModelID,fieldname,FieldUI,FieldDescription,fieldtype,fieldlength,isNull,Attribute,defaultValue,"&ExpModelID&" as isSubmodelID from Sp_ModelField where ModelID = "&ModelID&"")
''�ɹ�
Alert "��չ�ɹ�","setting.Model.list.asp"
%>