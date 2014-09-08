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
	  	'''增加记录
		Con.Execute("insert into Sp_ModelField (ModelID,fieldname,FieldUI,FieldDescription,fieldtype,fieldlength,isNull,Attribute,defaultValue,IsSearchForm,DefaultType,DictionaryID,IndexID) values ("&ModelID&",'"&fieldname&"','"&FieldUI&"','"&FieldDescription&"','"&fieldtype&"',"&fieldlength&","&isNull&",'"&Attribute&"','"&defaultValue&"',"&IsSearchForm&","&DefaultType&","&DictionaryID&","&IndexID&")")
		'''取模型的表名称
		Dim ModelArray:ModelArray = Con.QueryRow("select ModelTable from Sp_Model where ID="&ModelID&"",0)
		'''增加物理表结构
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
		Alert "出现异常!","setting.Model.Field.list.asp"
		else
		Alert "增加模型字段成功","setting.Model.Field.list.asp"
		end if
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
			if(document.getElementById("txt_fieldname").value=="")
			{
				alert("Sp_CMS提示\r\n\n字段名称必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">增加模型字段</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">模型名称：</TD>
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
				</select>&nbsp;<a href="setting.Model.add.asp">添加新模型</a>
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
				<option value="nvarchar">文本</option>
				<option value="ntext">备注</option>
				<option value="int">数字</option>
				<option value="datetime">时间</option>
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
				<input type="radio" name="txt_FieldUI" value="eWebeditor">eWebeditor编辑器
				<input type="radio" name="txt_FieldUI" value="datetime">时间
				<input type="radio" name="txt_FieldUI" value="checkbox">单选框
				<input type="radio" name="txt_FieldUI" value="select">选择框
				<input type="radio" name="txt_FieldUI" value="file">文件选择
				<input type="radio" name="txt_FieldUI" value="checkboxList">选择框checkbox
				<input type="radio" name="txt_FieldUI" value="textaddSelect">文本框与下拉框组合
				<input type="radio" name="txt_FieldUI" value="SelectToText">下拉框选择内容进文本框
				<input type="radio" name="txt_FieldUI" value="Maptext">地图标注框
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段属性：</TD>
				<TD><input name="txt_Attribute" class="input" type="text" ></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否是搜索字段：</TD>
				<TD><input name="txt_IsSearchForm" class="input" type="checkbox" value="1"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>默认值类型</TD>
				<TD>
				<input type="radio" name="txt_DefaultType" value="0" checked>手工输入&nbsp;
				<input type="radio" name="txt_DefaultType" value="1">数据字典
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段默认值：</TD>
				<TD><textarea name="txt_defaultValue" cols="20" rows="8"></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>数据字典值：</TD>
				<TD><input type="text" name="txt_DictionaryID" class="input" value="0" size="5"></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>显示顺序：</TD>
			    <TD><input name="txt_IndexID" type="text" class="input" id="txt_IndexID" value="0" size="5" ></TD>
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
