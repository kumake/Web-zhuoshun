<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim FormID:FormID = VerificationUrlParam("FormID","int","")
	'response.Write ID
	'response.End()
	Dim FormArray
	if FormID<>"" and isnumeric(FormID) then
		FormArray = Con.QueryRow("select ID,formname,FormTable,memo from Sp_Form where ID="&FormID&"",0)
		'response.Write ModelArray(2)
		'response.End()
	end if
	if action<>"" and action="save" then 
  		dim FormName:FormName = F.ReplaceFormText("txt_FormName")
		dim FormTable:FormTable = F.ReplaceFormText("txt_FormTable")
		dim memo:memo = F.ReplaceFormText("txt_memo")
		if IsDisplay="" then IsDisplay = 0
		Con.execute("update Sp_Form set formname='"&FormName&"',[memo]='"&memo&"' where ID="&FormID&"")
		Alert "�޸ı��ɹ�","setting.selfform.list.asp"
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
<form action="?action=save&FormID=<%=FormID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�޸�ģ��</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">�����ƣ�</TD>
				<TD><input name="txt_FormName" class="input" type="text" value="<%=FormArray(1)%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>ģ�ͱ����ƣ�</TD>
				<TD><input name="txt_FormTable" class="input" type="text" value="<%=FormArray(2)%>" disabled><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>��ע��</TD>
				<TD><textarea name="txt_memo" cols="50" rows="5" class="input"><%=FormArray(3)%></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�޸ı�" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
