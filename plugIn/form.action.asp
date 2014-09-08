<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<!--#include file="Setting.Config.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim FormID:FormID = VerificationUrlParam("FormID","int","0")
	
	''''
	if FormID=0 then
		response.Write "模型解释出现错误!"
	else
		Dim FormFieldArray:FormFieldArray = Con.QueryData("select F.fieldname,F.fieldtype,M.FormTable from Sp_FormField F,Sp_Form M where F.FormID=M.ID and F.FormID="&FormID&"")
		Dim FormTable
		'if isobject(FormFieldArray) then
			FormTable = FormFieldArray(2,0)
			if action<>"" and action="save" then
				'''增加信息
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
				'自定义字段组合sql
				'自定义字段组合sql
				''''' 
				dim sql:sql="insert into "&FormTable&" ("&fsql&") values ("&vsql&")"
				'response.Write sql
				'response.End()
				Con.execute(sql)
				Alert "提交表单成功!",""
			end if
		end if
	'end if	
%>