<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="../plugIn/Setting.Config.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ItemID:ItemID = VerificationUrlParam("ItemID","int","0")
	if action<>"" and action="save" then 	
		Dim Labname:Labname = F.ReplaceFormText("txt_name")
		Dim Labcontent:Labcontent = F.ReplaceFormText("txt_content")
		Con.execute("update sp_JsTemplate set name='"&Labname&"',code='"&Labcontent&"' where ID="&ItemID&"")
		Alert "�޸ĳɹ�","setting.Model.js.asp"
	end if
	if ItemID = 0 then
		Alert "��������ʧ��",""
	else
		set rs = Con.Query("select * from sp_JsTemplate where ID="&ItemID&"")
		if rs.recordcount=0 then
			Alert "�Ҳ�����¼",""
		else
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>��̨��ҳ</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
	<script src="Scripts/calender.js" type="text/javascript"></script>
	<script language="javascript" type="text/javascript">
		function checkform()
		{
			if(document.getElementById("txt_name").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n���Ʊ�����д!");
				return false;			
			}
			if(document.getElementById("txt_content").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n������д!");
				return false;			
			}

			return true;
		}
	</script>
</HEAD>
<BODY>
<form action="?action=save&ItemID=<%=ItemID%>" method="post" onSubmit="javascript:return checkform();">
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">����javascriptģ��</A></LI>
	  <LI><A href="setting.Model.js.asp">javascriptģ�����</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="10%">jsģ�����ƣ�</TD>
				<TD><input name="txt_name" type="text" value="<%=rs("name")%>" class="input"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD>���ݣ�</TD>
				<TD><textarea name="txt_content" cols="150" rows="20" class="input"><%=rs("code")%></textarea><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�޸ı�ǩ" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
<%
	end if
	end if
%>
