<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
  		dim UserGroup:UserGroup = F.ReplaceFormText("txt_UserGroup")
  		dim discount:discount = F.ReplaceFormText("txt_discount")
  		dim OverScore:OverScore = F.ReplaceFormText("txt_OverScore")
  		dim Overdiscount:Overdiscount = F.ReplaceFormText("txt_Overdiscount")
		dim upgradeNeedScore:upgradeNeedScore = F.ReplaceFormText("txt_upgradeNeedScore")
		
		Con.execute("insert into Sp_UserGroup (UserGroup,discount,OverScore,Overdiscount,upgradeNeedScore) values ('"&UserGroup&"',"&discount&","&OverScore&","&Overdiscount&","&upgradeNeedScore&")")
		Alert "�����û���ɹ�","pulg.User.Group.asp"
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
			if(document.getElementById("txt_UserGroup").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n�û������Ʊ�����д!");
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
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">�����û���</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">�û������ƣ�</TD>
				<TD colspan="3"><input name="txt_UserGroup" class="input" type="text"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="15%">�����ۿۣ�</TD>
				<TD colspan="3"><input name="txt_discount" class="input" type="text" value="100"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="15%">�����ִﵽ��</TD>
				<TD width="17%"><input name="txt_OverScore" class="input" type="text" value="0"><span class="huitext">&nbsp;����</span></TD>
			    <TD width="9%">�����ۿۣ�</TD>
			    <TD width="59%"><input name="txt_Overdiscount" class="input" type="text" value="0">
		        <span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="15%">������Ҫ���֣�</TD>
				<TD colspan="3"><input name="txt_upgradeNeedScore" class="input" type="text" value="0"><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="�����û���" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
