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
		Alert "数据查询失败",""
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
	  	'''增加记录
		Con.Execute("update Sp_FormField  set FieldUI ='"&FieldUI&"',FieldDescription='"&FieldDescription&"',isNull="&isNull&",Attribute='"&Attribute&"',defaultValue='"&defaultValue&"' where ID="&FieldArray(1)&"")
		'''增加物理表结构
		''''
		Alert "修改表单字段成功","setting.selfForm.Field.list.asp?FormID="&FieldArray(11)&""
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
				</select>&nbsp;<a href="setting.selfForm.add.asp">添加新表单</a>
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
				<option value="datetime"  <%if FieldArray(6)="datetime" then response.Write "selected"%>>时间</option>
				<option value="int"  <%if FieldArray(6)="int" then response.Write "selected"%>>数字</option>
				<option value="ntext"  <%if FieldArray(6)="ntext" then response.Write "selected"%>>备注</option>
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
				<input type="radio" name="txt_FieldUI" value="datetime"  <%if FieldArray(4)="datetime" then response.Write "checked"%>>时间
				<input type="radio" name="txt_FieldUI" value="checkbox"  <%if FieldArray(4)="checkbox" then response.Write "checked"%>>单选框
				<input type="radio" name="txt_FieldUI" value="select"  <%if FieldArray(4)="select" then response.Write "checked"%>>选择框
				<input type="radio" name="txt_FieldUI" value="file"  <%if FieldArray(4)="file" then response.Write "checked"%>>文件选择
				</TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段属性：</TD>
				<TD><input name="txt_Attribute" class="input" type="text" value="<%=FieldArray(9)%>"></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>字段默认值：</TD>
				<TD><textarea name="txt_defaultValue" cols="20" rows="8"><%=FieldArray(10)%></textarea></TD>
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
