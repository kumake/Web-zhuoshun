<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
  		dim FormName:FormName = F.ReplaceFormText("txt_FormName")
		dim FormTable:FormTable = F.ReplaceFormText("txt_FormTable")
		dim memo:memo = F.ReplaceFormText("txt_memo")
		Dim sql:sql = "insert into Sp_Form (formname,FormTable,[memo]) values ('"&FormName&"','"&FormTable&"','"&memo&"')"
		Con.Execute(sql)
		''''取最大值
    	Dim ItemArray:ItemArray = Con.QueryRow("select top 1 id from Sp_Form order by id desc",0)
		'''增加表单的对应记录
		Con.Execute("insert into Sp_FormField (FormID,fieldname,FieldUI,FieldDescription,fieldtype,fieldlength,isNull,Attribute,defaultValue) values ("&ItemArray(0)&",'title','text','标题','nvarchar',255,1,'','')")
		'response.Write sql
		'response.End()
		''''增加物理表
		Con.execute("select * into "&FormTable&" from Sp_FormTableStructure where 1<>1")
		Con.execute("alter table "&FormTable&" add primary key (id)")
		'''修改字段默认值
		Con.execute("alter table "&FormTable&" alter posttime datetime default now()")
		Alert "增加自定义表单成功","setting.selfForm.list.asp"
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
			if(document.getElementById("txt_FormNmae").value=="")
			{
				alert("Sp_CMS提示\r\n\n表单名称必须填写!");
				return false;			
			}
			if(document.getElementById("txt_FormTable").value=="")
			{
				alert("Sp_CMS提示\r\n\n表单表必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">增加自定义表单</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">表单名称：</TD>
				<TD><input name="txt_FormName" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>表单表名称：</TD>
				<TD><input name="txt_FormTable" class="input" type="text" value="form_"><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>表单描述：</TD>
				<TD><textarea name="txt_memo" cols="50" rows="5" class="input"></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="增加模型" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
