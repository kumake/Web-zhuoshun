<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ID:ID = VerificationUrlParam("ID","int","")
	Dim FieldArray
	if ID<>"" and isnumeric(ID) then
		FieldArray = Con.QueryRow("select M.ModelTable,F.ID,F.ModelID,F.fieldname,F.FieldUI,F.FieldDescription,F.fieldtype,F.fieldlength,F.isNull,F.Attribute,F.defaultValue,M.ID,F.IsSearchForm,F.IndexID,F.DefaultType,F.DictionaryID from Sp_ModelField F,Sp_Model M where F.ModelID=M.ID and F.ID="&ID&"",0)
		'response.Write FieldArray(8)
	else
		Alert "数据查询失败",""
	end if
	if action<>"" and action="save" then
	  'Dim ModelID:ModelID = F.ReplaceFormText("txt_ModelID")
	  'Dim fieldname:fieldname = F.ReplaceFormText("txt_fieldname")
	  Dim FieldUI:FieldUI = F.ReplaceFormText("txt_FieldUI")
	  Dim FieldDescription:FieldDescription = F.ReplaceFormText("txt_FieldDescription")
	  'Dim fieldtype:fieldtype = F.ReplaceFormText("txt_fieldtype")
	  'Dim fieldlength:fieldlength = F.ReplaceFormText("txt_fieldlength")
	  Dim IsSearchForm:IsSearchForm = F.ReplaceFormText("txt_IsSearchForm")
	  Dim isNull:isNull = F.ReplaceFormText("txt_isNull")
	  if isNull ="" then isNull = 0
	  if IsSearchForm = "" then IsSearchForm = 0
	  Dim Attribute:Attribute = F.ReplaceFormText("txt_Attribute")
	  Dim defaultValue:defaultValue = F.ReplaceFormText("txt_defaultValue")
	  Dim IndexID:IndexID = F.ReplaceFormText("txt_IndexID")
	  Dim DefaultType:DefaultType = F.ReplaceFormText("txt_DefaultType")
	  Dim DictionaryID:DictionaryID = F.ReplaceFormText("txt_DictionaryID")
	  if DefaultType<>"" then DefaultType=0
	  if DictionaryID<>"" then DictionaryID=0
	  
	  if IndexID = "" then IndexID = 0
	  	'''增加记录
		Con.Execute("update Sp_ModelField  set FieldUI ='"&FieldUI&"',FieldDescription='"&FieldDescription&"',isNull="&isNull&",Attribute='"&Attribute&"',defaultValue='"&defaultValue&"',IsSearchForm="&IsSearchForm&",DictionaryID="&DictionaryID&",DefaultType="&DefaultType&",IndexID="&IndexID&" where ID="&FieldArray(1)&"")
		'''增加物理表结构
		''''
		Alert "增加模型字段成功","setting.Model.Field.list.asp?ModelID="&FieldArray(11)&""
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
<form action="?action=save&ID=<%=ID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">修改模型字段</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">模型名称：</TD>
				<TD>
				<select name="txt_ModelID" disabled>
				<%
					set Rs = Con.Query("select id,Modelname from Sp_Model")
					if rs.recordcount<>0 then
					do while not rs.eof
					response.Write "<option value="&rs("id")&""
					if Cint(FieldArray(2))=Cint(rs("id")) then response.Write " selected"
					response.Write ">"&rs("Modelname")&"</option>"
					rs.movenext
					loop
					end if
				%>
				</select>&nbsp;<a href="setting.Model.add.asp">添加新模型</a>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段名称：</TD>
				<TD><input name="txt_fieldname" class="input" type="text" value="<%=FieldArray(3)%>" disabled><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否必填：</TD>
				<TD><input name="txt_isNull" class="input" type="checkbox" value="1" <%if FieldArray(8)=1 then response.Write "checked"%>></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段页面描述：</TD>
				<TD><input name="txt_FieldDescription" class="input" type="text" value="<%=FieldArray(5)%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段类型：</TD>
				<TD>
				<select name="txt_fieldtype" disabled>
				<option value="nvarchar"  <%if FieldArray(6)="nvarchar" then response.Write "selected"%>>文本</option>
				<option value="ntext"  <%if FieldArray(6)="ntext" then response.Write "selected"%>>备注</option>
				<option value="int"  <%if FieldArray(6)="int" then response.Write "selected"%>>数字</option>
				<option value="datetime"  <%if FieldArray(6)="datetime" then response.Write "selected"%>>时间</option>
				</select>
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段长度：</TD>
				<TD><input name="txt_fieldlength" class="input" type="text"  value="<%=FieldArray(7)%>" disabled></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段页面显示形式：</TD>
				<TD>
				<input name="txt_FieldUI" type="radio" value="text"  <%if FieldArray(4)="text" then response.Write "checked"%>>文本
				<input type="radio" name="txt_FieldUI" value="textarea"  <%if FieldArray(4)="textarea" then response.Write "checked"%>>普通多行文本框
				<input type="radio" name="txt_FieldUI" value="html"  <%if FieldArray(4)="html" then response.Write "checked"%>>带编辑器的文本框
				<input type="radio" name="txt_FieldUI" value="eWebeditor"  <%if FieldArray(4)="eWebeditor" then response.Write "checked"%>>eWebeditor编辑器
				<input type="radio" name="txt_FieldUI" value="datetime"  <%if FieldArray(4)="datetime" then response.Write "checked"%>>时间
				<input type="radio" name="txt_FieldUI" value="checkbox"  <%if FieldArray(4)="checkbox" then response.Write "checked"%>>单选框
				<input type="radio" name="txt_FieldUI" value="select"  <%if FieldArray(4)="select" then response.Write "checked"%>>选择框
				<input type="radio" name="txt_FieldUI" value="file"  <%if FieldArray(4)="file" then response.Write "checked"%>>文件选择
				<input type="radio" name="txt_FieldUI" value="checkboxList"  <%if FieldArray(4)="checkboxList" then response.Write "checked"%>>选择框checkbox
				<input type="radio" name="txt_FieldUI" value="textaddSelect"  <%if FieldArray(4)="textaddSelect" then response.Write "checked"%>>文本框与下拉框组合
				<input type="radio" name="txt_FieldUI" value="SelectToText"  <%if FieldArray(4)="SelectToText" then response.Write "checked"%>>下拉框选择内容进文本框
				<input type="radio" name="txt_FieldUI" value="Maptext"  <%if FieldArray(4)="Maptext" then response.Write "checked"%>>地图标注框
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段属性：</TD>
				<TD><input name="txt_Attribute" class="input" type="text" value="<%=FieldArray(9)%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>是否是搜索字段：</TD>
				<TD><input name="txt_IsSearchForm" class="input" type="checkbox" value="1" <%if FieldArray(12)=1 then response.Write " checked"%>></TD>
			  </TR>
			  
			  <TR class="content-td1">
				<TD>默认值类型</TD>
				<TD>
				<input type="radio" name="txt_DefaultType" value="0" <%if FieldArray(14)=0 then response.Write "checked"%>>手工输入&nbsp;
				<input type="radio" name="txt_DefaultType" value="1" <%if FieldArray(14)=1 then response.Write "checked"%>>数据字典
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段默认值：</TD>
				<TD><textarea name="txt_defaultValue" cols="20" rows="8"><%=FieldArray(10)%></textarea></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>数据字典值：</TD>
				<TD><input type="text" name="txt_DictionaryID" class="input" value="<%=FieldArray(14)%>" size="5"></TD>
			  </TR>
			  <TR class="content-td1">
			    <TD>显示顺序：</TD>
			    <TD><input name="txt_IndexID" type="text" class="input" id="txt_IndexID" value="<%=FieldArray(13)%>" size="5" ></TD>
		      </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="修改字段" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
