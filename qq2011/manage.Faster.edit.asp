<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	Dim ID:ID = VerificationUrlParam("ID","int","0")
	if action<>"" and action="save" then 
  		dim MenuName:MenuName = F.ReplaceFormText("txt_MenuName")
		dim MenuUrl:MenuUrl = F.ReplaceFormText("txt_MenuUrl")
		dim Memo:Memo = F.ReplaceFormText("txt_Memo")
		'response.Write rolepower
		'response.End()
		Con.execute("update Sp_FasterMenu set MenuName='"&MenuName&"',MenuUrl='"&MenuUrl&"',[Memo]='"&Memo&"' where ID="&ID&" ")
		Alert "�޸Ľ�ɫ�ɹ�","manage.Faster.list.asp"
	end if
	if ID=0 then
		Alert "��������ʧ��",""
	else
		set rs = Con.Query("select * from Sp_FasterMenu where ID="&ID&"")
		if rs.recordcount=0 then
			Alert "��Ӧ��Ϣ�Ҳ���",""
		else
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
			if(document.getElementById("txt_MenuName").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�˵����Ʊ�����д!");
				return false;			
			}
			if(document.getElementById("txt_MenuUrl").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�˵�URL������д!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" >�޸Ľ�ɫ</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="14%">�˵����ƣ�</TD>
				<TD><input name="txt_MenuName"  class="input" type="text" value="<%=rs("MenuName")%>"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="14%">�˵�URL��</TD>
				<TD><input name="txt_MenuUrl" type="text" class="input" value="<%=rs("MenuUrl")%>" size="70">
				<span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="14%">Memo��</TD>
				<TD><textarea name="txt_Memo" cols="70" rows="8" class="input"><%=rs("Memo")%></textarea>
				<span class="huitext">&nbsp;����</span></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�޸Ĳ˵�" class="button" type="submit">
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
