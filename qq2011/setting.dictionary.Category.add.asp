<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
  		dim txtdictionary:txtdictionary = F.ReplaceFormText("txt_dictionary")
		dim txtdictionary_en:txtdictionary_en = F.ReplaceFormText("txt_dictionary_en")
		Con.execute("insert into Sp_dictionaryCategory (dictionary,dictionary_en) values ('"&txtdictionary&"','"&txtdictionary_en&"')")
		'''''取最新的ID
		Dim ItemArray:ItemArray = Con.QueryRow("select top 1 id from Sp_dictionaryCategory order by id desc",0)
		''''
		Con.execute("insert into Sp_dictionary (categoryID,categoryname,Pathstr,parentID,Depth) values ("&ItemArray(0)&",'默认类别','0',0,1)")
		'''''取最新的ID
		ItemArray = Con.QueryRow("select top 1 id from Sp_dictionary order by id desc",0)
		''''
		Con.execute("update Sp_dictionary set Pathstr='"&ItemArray(0)&"' where ID="&ItemArray(0)&"")
		Alert "增加类别成功","setting.dictionary.Category.asp"
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
			if(document.getElementById("txt_dictionary").value=="")
			{
				alert("Sp_CMS提示\r\n\n类别名称必须填写!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">增加字典类别</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">类别名称：</TD>
				<TD><input name="txt_dictionary" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>类别名称(英文)：</TD>
				<TD><input name="txt_dictionary_en" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="增加类别" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
