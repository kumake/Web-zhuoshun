<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Field.asp"-->
<%
	response.Expires=0
	Dim action:action = VerificationUrlParam("action","string","")
	Dim FormID:FormID = VerificationUrlParam("FormID","string","")
	Dim Content:Content = ""
	Dim FormFieldArray:FormFieldArray = Con.QueryData("select fieldname,FieldUI,FieldDescription,fieldtype,isNull,Attribute,defaultValue from Sp_FormField where FormID="&FormID&"")
	'if isobject(FormFieldArray) then
		if action<>"" and action="html" then 
			Content = Content & "<script>"&vbCrlf
			Content = Content & "	function checkform()"&vbCrlf
			Content = Content & "	{"&vbCrlf
			''循环
			for i=0 to ubound(FormFieldArray,2) step 1
				if FormFieldArray(4,i)=1 then
					Content = Content & "if(document.getElementById('txt_"&FormFieldArray(0,i)&"').value=='')"&vbCrlf
					Content = Content & "{"&vbCrlf
					Content = Content & "	alert('Sp_CMS提示\r\n\n"&FormFieldArray(2,i)&"必须填写!');"&vbCrlf
					Content = Content & "	return false;"&vbCrlf		
					Content = Content & "}"&vbCrlf
				end if
			next
			Content = Content & "		return true;"&vbCrlf
			Content = Content & "	}"&vbCrlf
			Content = Content & "</script>"&vbCrlf
			Content = Content & "<form action='../plugIn/form.action.asp?action=save&FormID="&FormID&"' method='post'  onSubmit='javascript:return checkform();'>"&vbCrlf
			Content = Content & "<TABLE cellSpacing=1 width='100%'>"&vbCrlf
			for i=0 to ubound(FormFieldArray,2) step 1
			Content = Content & "<TR class=content-td1>"&vbCrlf
			Content = Content & "<TD>"&vbCrlf
			Content = Content & FormFieldArray(2,i)
			Content = Content & "</TD>"&vbCrlf
			Content = Content & "<TD>"&vbCrlf
			Content = Content & FieldUIType(FormFieldArray(0,i),FormFieldArray(1,i),FormFieldArray(6,i),"",FormFieldArray(4,i),FormFieldArray(5,i))
			if Cint(FormFieldArray(4,i))=1 then Content = Content &" <span class='huitext'>必填</span>"
			Content = Content & "</TD>"&vbCrlf
			Content = Content & "</TR>"&vbCrlf
			next
			Content = Content & "<tr>"&vbCrlf
			Content = Content & "<TD colspan='2'>"&vbCrlf
			Content = Content & "<input name='btnactionform' value='提交表单' type='submit'>"&vbCrlf
			Content = Content & "</TD>"&vbCrlf
			Content = Content & "</tr>"&vbCrlf
			Content = Content & "</Table>"&vbCrlf
			Content = Content & "</form>"&vbCrlf
		else
			Content = Content & "<script language='javascript' src='setting.selfForm.js.asp?action=save&FormID="&FormID&"'></script>"
		end if
		'response.Write Content
	'end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">代码调用</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<BR>
		<DIV>
			<TABLE cellSpacing=1 width="100%">
			<TR class=content-td1>
				<TD>
				<textarea cols="67" rows="10"><%=Content%></textarea>
				</TD>
			</TR>
			</TABLE>
		</DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
