<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ID:ID = VerificationUrlParam("ID","int","")
	Dim FieldArray
	if ID<>"" and isnumeric(ID) then
		FieldArray = Con.QueryRow("select M.FormTable,F.ID,F.formID,F.fieldname,F.FieldUI,F.FieldDescription,F.fieldtype,F.fieldlength,F.isNull,F.Attribute,F.defaultValue,M.ID from Sp_FormField F,Sp_Form M where F.FormID=M.ID and F.ID="&ID&"",0)
		'response.Write FieldArray(8)
	else
		Alert "���ݲ�ѯʧ��",""
	end if
	if action<>"" and action="save" then
	  'Dim ModelID:ModelID = F.ReplaceFormText("txt_ModelID")
	  'Dim fieldname:fieldname = F.ReplaceFormText("txt_fieldname")
	  Dim FieldUI:FieldUI = F.ReplaceFormText("txt_FieldUI")
	  Dim FieldDescription:FieldDescription = F.ReplaceFormText("txt_FieldDescription")
	  'Dim fieldtype:fieldtype = F.ReplaceFormText("txt_fieldtype")
	  'Dim fieldlength:fieldlength = F.ReplaceFormText("txt_fieldlength")
	  Dim isNull:isNull = F.ReplaceFormText("txt_isNull")
	  if isNull ="" then isNull = 0
	  Dim Attribute:Attribute = F.ReplaceFormText("txt_Attribute")
	  Dim defaultValue:defaultValue = F.ReplaceFormText("txt_defaultValue")
	  	'''���Ӽ�¼
		Con.Execute("update Sp_FormField  set FieldUI ='"&FieldUI&"',FieldDescription='"&FieldDescription&"',isNull="&isNull&",Attribute='"&Attribute&"',defaultValue='"&defaultValue&"' where ID="&FieldArray(1)&"")
		'''���������ṹ
		''''
		Alert "�޸ı��ֶγɹ�","setting.selfForm.Field.list.asp?FormID="&FieldArray(11)&""
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/HtmlEditor.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_ModelNmae").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\nģ�����Ʊ�����д!");
				return false;			
			}
			if(document.getElementById("txt_ModelTable").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\nģ�ͱ������д!");
				return false;			
			}
			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=save&ID=<%=ID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�޸�ģ���ֶ�</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">ģ�����ƣ�</TD>
				<TD>
				<select name="txt_ModelID" disabled>
				<%
					set Rs = Con.Query("select id,Formname from Sp_Form")
					if rs.recordcount<>0 then
					do while not rs.eof
					response.Write "<option value="&rs("id")&""
					if Cint(FieldArray(2))=Cint(rs("id")) then response.Write " selected"
					response.Write ">"&rs("Formname")&"</option>"
					rs.movenext
					loop
					end if
				%>
				</select>&nbsp;<a href="setting.selfForm.add.asp">����±�</a>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ����ƣ�</TD>
				<TD><input name="txt_fieldname" class="input" type="text" value="<%=FieldArray(3)%>" disabled><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƿ���</TD>
				<TD><input name="txt_isNull" class="input" type="checkbox" value="1" <%if FieldArray(8)=1 then response.Write "checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ�ҳ��������</TD>
				<TD><input name="txt_FieldDescription" class="input" type="text" value="<%=FieldArray(5)%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ����ͣ�</TD>
				<TD>
				<select name="txt_fieldtype" disabled>
				<option value="nvarchar"  <%if FieldArray(6)="nvarchar" then response.Write "selected"%>>�ı�</option>
				<option value="datetime"  <%if FieldArray(6)="datetime" then response.Write "selected"%>>ʱ��</option>
				<option value="int"  <%if FieldArray(6)="int" then response.Write "selected"%>>����</option>
				<option value="ntext"  <%if FieldArray(6)="ntext" then response.Write "selected"%>>��ע</option>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶγ��ȣ�</TD>
				<TD><input name="txt_fieldlength" class="input" type="text"  value="<%=FieldArray(7)%>" disabled></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ�ҳ����ʾ��ʽ��</TD>
				<TD>
				<input name="txt_FieldUI" type="radio" value="text"  <%if FieldArray(4)="text" then response.Write "checked"%>>�ı�
				<input type="radio" name="txt_FieldUI" value="textarea"  <%if FieldArray(4)="textarea" then response.Write "checked"%>>��ͨ�����ı���
				<input type="radio" name="txt_FieldUI" value="html"  <%if FieldArray(4)="html" then response.Write "checked"%>>���༭�����ı���
				<input type="radio" name="txt_FieldUI" value="datetime"  <%if FieldArray(4)="datetime" then response.Write "checked"%>>ʱ��
				<input type="radio" name="txt_FieldUI" value="checkbox"  <%if FieldArray(4)="checkbox" then response.Write "checked"%>>��ѡ��
				<input type="radio" name="txt_FieldUI" value="select"  <%if FieldArray(4)="select" then response.Write "checked"%>>ѡ���
				<input type="radio" name="txt_FieldUI" value="file"  <%if FieldArray(4)="file" then response.Write "checked"%>>�ļ�ѡ��
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ����ԣ�</TD>
				<TD><input name="txt_Attribute" class="input" type="text" value="<%=FieldArray(9)%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ�Ĭ��ֵ��</TD>
				<TD><textarea name="txt_defaultValue" cols="20" rows="8"><%=FieldArray(10)%></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�޸��ֶ�" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
