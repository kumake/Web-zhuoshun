<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then
	  Dim GroupID:GroupID = F.ReplaceFormText("txt_GroupID")
	  Dim fieldname:fieldname = F.ReplaceFormText("txt_fieldname")
	  Dim FieldUI:FieldUI = F.ReplaceFormText("txt_FieldUI")
	  Dim FieldDescription:FieldDescription = F.ReplaceFormText("txt_FieldDescription")
	  Dim fieldtype:fieldtype = F.ReplaceFormText("txt_fieldtype")
	  Dim fieldlength:fieldlength = F.ReplaceFormText("txt_fieldlength")
	  Dim isNull:isNull = F.ReplaceFormText("txt_isNull")
	  Dim isCreateField:isCreateField = F.ReplaceFormText("txt_isCreateField")
	  Dim Attribute:Attribute = F.ReplaceFormText("txt_Attribute")
	  Dim defaultValue:defaultValue = F.ReplaceFormText("txt_defaultValue")
	  if isNull="" then isNull = 0
	  if isCreateField="" then isCreateField = 0
	  	'''增加记录
		Con.Execute("insert into Sp_UserExpField (GroupID,fieldname,FieldUI,FieldDescription,fieldtype,fieldlength,isNull,Attribute,defaultValue) values ("&GroupID&",'"&fieldname&"','"&FieldUI&"','"&FieldDescription&"','"&fieldtype&"',"&fieldlength&","&isNull&",'"&Attribute&"','"&defaultValue&"')")
		'''增加物理表结构
		Dim fieldSql
		if isCreateField<>0 then
		'''''创建实体表字段
			select case fieldtype
				case "char","nvarchar"
					fieldSql = "alter table Sp_User add "&fieldname&" "&fieldtype&"("&fieldlength&")"
				case "int"
					fieldSql = "alter table Sp_User add "&fieldname&" "&fieldtype&""
				case "ntext"
					fieldSql = "alter table Sp_User add "&fieldname&" "&fieldtype&""
				case "datetime"
					fieldSql = "alter table Sp_User add "&fieldname&" "&fieldtype&""
				case else 
					fieldSql = "alter table Sp_User add "&fieldname&" "&fieldtype&"(255)"
			end select	  
			Con.execute(fieldSql)
		end if
		''''
		Alert "增加字段成功","setting.user.Field.expand.asp"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/HtmlEditor.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_ModelNmae").value=="")
			{
				alert("Sp_CMS提示\r\n\n模型名称必须填写!");
				return false;			
			}
			if(document.getElementById("txt_ModelTable").value=="")
			{
				alert("Sp_CMS提示\r\n\n模型表必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">增加用户字段扩展</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">表单名称：</TD>
				<TD>
				<select name="txt_GroupID">
				<%
					set Rs = Con.Query("select id,UserGroup from Sp_UserGroup")
					if rs.recordcount<>0 then
					do while not rs.eof
					response.Write "<option value="&rs("id")&">"&rs("UserGroup")&"</option>"
					rs.movenext
					loop
					end if
				%>
				</select>&nbsp;<a href="pulg.user.group.add.asp">添加新用户组</a>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段名称：</TD>
				<TD><input name="txt_fieldname" class="input" type="text" value="Field_"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否必填：</TD>
				<TD><input name="txt_isNull" class="input" type="checkbox" value="1"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段页面描述：</TD>
				<TD><input name="txt_FieldDescription" class="input" type="text" ></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段类型：</TD>
				<TD>
				<select name="txt_fieldtype">
				<option value="datetime">时间</option>
				<option value="nvarchar">文本</option>
				<option value="int">数字</option>
				<option value="ntext">备注</option>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段长度：</TD>
				<TD><input name="txt_fieldlength" class="input" type="text" ></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段页面显示形式：</TD>
				<TD>
				<input name="txt_FieldUI" type="radio" value="text" checked>文本
				<input type="radio" name="txt_FieldUI" value="textarea">普通多行文本框
				<input type="radio" name="txt_FieldUI" value="html">带编辑器的文本框
				<input type="radio" name="txt_FieldUI" value="datetime">时间
				<input type="radio" name="txt_FieldUI" value="checkbox">单选框
				<input type="radio" name="txt_FieldUI" value="select">选择框
				<input type="radio" name="txt_FieldUI" value="file">文件选择
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段属性：</TD>
				<TD><input name="txt_Attribute" class="input" type="text" ></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段默认值：</TD>
				<TD><textarea name="txt_defaultValue" cols="20" rows="8"></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>创建实体表字段：</TD>
				<TD><input name="txt_isCreateField" class="input" checked type="checkbox" value="1"></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="增加字段" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
