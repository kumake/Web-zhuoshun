<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then
	  Dim formID:formID = F.ReplaceFormText("txt_formID")
	  Dim fieldname:fieldname = F.ReplaceFormText("txt_fieldname")
	  Dim FieldUI:FieldUI = F.ReplaceFormText("txt_FieldUI")
	  Dim FieldDescription:FieldDescription = F.ReplaceFormText("txt_FieldDescription")
	  Dim fieldtype:fieldtype = F.ReplaceFormText("txt_fieldtype")
	  Dim fieldlength:fieldlength = F.ReplaceFormText("txt_fieldlength")
	  Dim isNull:isNull = F.ReplaceFormText("txt_isNull")
	  Dim Attribute:Attribute = F.ReplaceFormText("txt_Attribute")
	  Dim defaultValue:defaultValue = F.ReplaceFormText("txt_defaultValue")
	  if isNull ="" then isNull = 0
	  	'''���Ӽ�¼
		Con.Execute("insert into Sp_FormField (FormID,fieldname,FieldUI,FieldDescription,fieldtype,fieldlength,isNull,Attribute,defaultValue) values ("&formID&",'"&fieldname&"','"&FieldUI&"','"&FieldDescription&"','"&fieldtype&"',"&fieldlength&","&isNull&",'"&Attribute&"','"&defaultValue&"')")
		'''ȡģ�͵ı�����
		Dim FormArray:FormArray = Con.QueryRow("select FormTable from Sp_Form where ID="&formID&"",0)
		'''���������ṹ
		Dim fieldSql
		select case fieldtype
			case "char","nvarchar"
				fieldSql = "alter table "&FormArray(0)&" add "&fieldname&" "&fieldtype&"("&fieldlength&")"
			case "int"
				fieldSql = "alter table "&FormArray(0)&" add "&fieldname&" "&fieldtype&""
			case "ntext"
				fieldSql = "alter table "&FormArray(0)&" add "&fieldname&" "&fieldtype&""
			case "datetime"
				fieldSql = "alter table "&FormArray(0)&" add "&fieldname&" "&fieldtype&""
			case else 
				fieldSql = "alter table "&FormArray(0)&" add "&fieldname&" "&fieldtype&"(255)"
		end select	  
		Con.execute(fieldSql)
		''''
		Alert "���ӱ��Զ����ֶγɹ�","setting.selfform.Field.list.asp"
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
<form action="?action=save" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">���ӱ��ֶ�</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">�����ƣ�</TD>
				<TD>
				<select name="txt_formID">
				<%
					set Rs = Con.Query("select id,Formname from Sp_Form")
					if rs.recordcount<>0 then
					do while not rs.eof
					response.Write "<option value="&rs("id")&">"&rs("Formname")&"</option>"
					rs.movenext
					loop
					end if
				%>
				</select>&nbsp;<a href="setting.selfForm.add.asp">����µ��Զ����</a>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ����ƣ�</TD>
				<TD><input name="txt_fieldname" class="input" type="text" value="Field_"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƿ���</TD>
				<TD><input name="txt_isNull" class="input" type="checkbox" value="1"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ�ҳ��������</TD>
				<TD><input name="txt_FieldDescription" class="input" type="text" ></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ����ͣ�</TD>
				<TD>
				<select name="txt_fieldtype">
				<option value="datetime">ʱ��</option>
				<option value="nvarchar">�ı�</option>
				<option value="int">����</option>
				<option value="ntext">��ע</option>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶγ��ȣ�</TD>
				<TD><input name="txt_fieldlength" class="input" type="text" ></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ�ҳ����ʾ��ʽ��</TD>
				<TD>
				<input name="txt_FieldUI" type="radio" value="text" checked>�ı�
				<input type="radio" name="txt_FieldUI" value="textarea">��ͨ�����ı���
				<input type="radio" name="txt_FieldUI" value="html">���༭�����ı���
				<input type="radio" name="txt_FieldUI" value="datetime">ʱ��
				<input type="radio" name="txt_FieldUI" value="checkbox">��ѡ��
				<input type="radio" name="txt_FieldUI" value="select">ѡ���
				<input type="radio" name="txt_FieldUI" value="file">�ļ�ѡ��
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ����ԣ�</TD>
				<TD><input name="txt_Attribute" class="input" type="text" ></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ�Ĭ��ֵ��</TD>
				<TD><textarea name="txt_defaultValue" cols="20" rows="8"></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�����ֶ�" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
