<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<!--#include file="Setting.Config.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim FormID:FormID = VerificationUrlParam("FormID","int","0")
	
	''''
	if FormID=0 then
		response.Write "ģ�ͽ��ͳ��ִ���!"
	else
		Dim FormFieldArray:FormFieldArray = Con.QueryData("select F.fieldname,F.fieldtype,M.FormTable from Sp_FormField F,Sp_Form M where F.FormID=M.ID and F.FormID="&FormID&"")
		Dim FormTable
		'if isobject(FormFieldArray) then
			FormTable = FormFieldArray(2,0)
			if action<>"" and action="save" then
				'''������Ϣ
				dim fsql:fsql=""
				dim vsql:vsql=""
				dim tempvalue:tempvalue = ""
				for i=0 to ubound(FormFieldArray,2)
					if fsql<>"" then 
						fsql = fsql &","&FormFieldArray(0,i)
					else
						fsql = FormFieldArray(0,i)
					end if
					
					if FormFieldArray(1,i)="int" then
						tempvalue = F.ReplaceFormText("txt_"&FormFieldArray(0,i))
					else
						tempvalue = "'"&F.ReplaceFormText("txt_"&FormFieldArray(0,i))&"'"
					end if
					
					if vsql<>"" then 
						vsql = vsql &","&tempvalue
					else
						vsql = tempvalue
					end if
				next	 		 
				'�Զ����ֶ����sql
				'�Զ����ֶ����sql
				''''' 
				dim sql:sql="insert into "&FormTable&" ("&fsql&") values ("&vsql&")"
				'response.Write sql
				'response.End()
				Con.execute(sql)
				Alert "�ύ���ɹ�!",""
			end if
		end if
	'end if	
%>