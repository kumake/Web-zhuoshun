<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then
	  Dim ModelID:ModelID = F.ReplaceFormText("txt_ModelID")
	  Dim fieldname:fieldname = F.ReplaceFormText("txt_fieldname")
	  Dim FieldUI:FieldUI = F.ReplaceFormText("txt_FieldUI")
	  Dim FieldDescription:FieldDescription = F.ReplaceFormText("txt_FieldDescription")
	  Dim fieldtype:fieldtype = F.ReplaceFormText("txt_fieldtype")
	  Dim fieldlength:fieldlength = F.ReplaceFormText("txt_fieldlength")
	  Dim isNull:isNull = F.ReplaceFormText("txt_isNull")
	  Dim Attribute:Attribute = F.ReplaceFormText("txt_Attribute")
	  Dim defaultValue:defaultValue = F.ReplaceFormText("txt_defaultValue")
	  Dim IsSearchForm:IsSearchForm = F.ReplaceFormText("txt_IsSearchForm")
	  Dim IndexID:IndexID = F.ReplaceFormText("txt_IndexID")
	  Dim DefaultType:DefaultType = F.ReplaceFormText("txt_DefaultType")
	  Dim DictionaryID:DictionaryID = F.ReplaceFormText("txt_DictionaryID")
	  if DefaultType<>"" then DefaultType=0
	  if DictionaryID<>"" then DictionaryID=0
	  if IndexID = "" then IndexID = 0
	  if isNull = "" then isNull = 0 
	  if fieldlength="" then fieldlength = 0
	  if IsSearchForm = "" then IsSearchForm = 0
	  	'''���Ӽ�¼
		Con.Execute("insert into Sp_ModelField (ModelID,fieldname,FieldUI,FieldDescription,fieldtype,fieldlength,isNull,Attribute,defaultValue,IsSearchForm,DefaultType,DictionaryID,IndexID) values ("&ModelID&",'"&fieldname&"','"&FieldUI&"','"&FieldDescription&"','"&fieldtype&"',"&fieldlength&","&isNull&",'"&Attribute&"','"&defaultValue&"',"&IsSearchForm&","&DefaultType&","&DictionaryID&","&IndexID&")")
		'''ȡģ�͵ı�����
		Dim ModelArray:ModelArray = Con.QueryRow("select ModelTable from Sp_Model where ID="&ModelID&"",0)
		'''���������ṹ
		Dim fieldSql
		Dim fieldSql_En
		select case fieldtype
			case "char","nvarchar"
				fieldSql = "alter table "&ModelArray(0)&" add "&fieldname&" "&fieldtype&"("&fieldlength&")"
				fieldSql_En = "alter table "&ModelArray(0)&" add "&fieldname&"_en "&fieldtype&"("&fieldlength&")"
			case "int"
				fieldSql = "alter table "&ModelArray(0)&" add "&fieldname&" "&fieldtype&""
				fieldSql_En = "alter table "&ModelArray(0)&" add "&fieldname&"_en "&fieldtype&""
			case "ntext"
				fieldSql = "alter table "&ModelArray(0)&" add "&fieldname&" "&fieldtype&""
				fieldSql_En = "alter table "&ModelArray(0)&" add "&fieldname&"_en "&fieldtype&""
			case "datetime"
				fieldSql = "alter table "&ModelArray(0)&" add "&fieldname&" "&fieldtype&""
				fieldSql_En = "alter table "&ModelArray(0)&" add "&fieldname&"_en "&fieldtype&""
			case else 
				fieldSql = "alter table "&ModelArray(0)&" add "&fieldname&" "&fieldtype&"(255)"
				fieldSql_En = "alter table "&ModelArray(0)&" add "&fieldname&"_en "&fieldtype&"(255)"
		end select	  
		Con.execute(fieldSql)
		if config_Verison= 2 then
		Con.execute(fieldSql_En)
		end if
		''''
		If Err.number<>0 Then
		Alert "�����쳣!","setting.Model.Field.list.asp"
		else
		Alert "����ģ���ֶγɹ�","setting.Model.Field.list.asp"
		end if
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
			if(document.getElementById("txt_fieldname").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�ֶ����Ʊ�����д!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">����ģ���ֶ�</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">ģ�����ƣ�</TD>
				<TD>
				<select name="txt_ModelID">
				<%
					set Rs = Con.Query("select id,Modelname from Sp_Model")
					if rs.recordcount<>0 then
					do while not rs.eof
					response.Write "<option value="&rs("id")&">"&rs("Modelname")&"</option>"
					rs.movenext
					loop
					end if
				%>
				</select>&nbsp;<a href="setting.Model.add.asp">�����ģ��</a>
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
				<option value="nvarchar">�ı�</option>
				<option value="ntext">��ע</option>
				<option value="int">����</option>
				<option value="datetime">ʱ��</option>
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
				<input type="radio" name="txt_FieldUI" value="eWebeditor">eWebeditor�༭��
				<input type="radio" name="txt_FieldUI" value="datetime">ʱ��
				<input type="radio" name="txt_FieldUI" value="checkbox">��ѡ��
				<input type="radio" name="txt_FieldUI" value="select">ѡ���
				<input type="radio" name="txt_FieldUI" value="file">�ļ�ѡ��
				<input type="radio" name="txt_FieldUI" value="checkboxList">ѡ���checkbox
				<input type="radio" name="txt_FieldUI" value="textaddSelect">�ı��������������
				<input type="radio" name="txt_FieldUI" value="SelectToText">������ѡ�����ݽ��ı���
				<input type="radio" name="txt_FieldUI" value="Maptext">��ͼ��ע��
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ����ԣ�</TD>
				<TD><input name="txt_Attribute" class="input" type="text" ></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�Ƿ��������ֶΣ�</TD>
				<TD><input name="txt_IsSearchForm" class="input" type="checkbox" value="1"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>Ĭ��ֵ����</TD>
				<TD>
				<input type="radio" name="txt_DefaultType" value="0" checked>�ֹ�����&nbsp;
				<input type="radio" name="txt_DefaultType" value="1">�����ֵ�
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�ֶ�Ĭ��ֵ��</TD>
				<TD><textarea name="txt_defaultValue" cols="20" rows="8"></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>�����ֵ�ֵ��</TD>
				<TD><input type="text" name="txt_DictionaryID" class="input" value="0" size="5"></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>��ʾ˳��</TD>
			    <TD><input name="txt_IndexID" type="text" class="input" id="txt_IndexID" value="0" size="5" ></TD>
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
