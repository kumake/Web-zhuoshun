<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
  		dim paytype:paytype = F.ReplaceFormText("txt_paytype")
  		dim payfield1:payfield1 = F.ReplaceFormText("txt_payfield1")
  		dim payfield2:payfield2 = F.ReplaceFormText("txt_payfield2")
  		dim payfield3:payfield3 = F.ReplaceFormText("txt_payfield3")
  		dim payfield4:payfield4 = F.ReplaceFormText("txt_payfield4")
  		dim payfield5:payfield5 = F.ReplaceFormText("txt_payfield5")
		dim memo:memo = F.ReplaceFormText("txt_memo")
		Con.execute("insert into Sp_orderPay ([paytype],payfield1,payfield2,payfield3,payfield4,payfield5,[memo]) values ('"&paytype&"','"&payfield1&"','"&payfield2&"','"&payfield3&"','"&payfield4&"','"&payfield5&"','"&memo&"')")
		Alert "����֧����ʽ�ɹ�","pulg.order.pay.list.asp"
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
			if(document.getElementById("txt_paytype").value=="")
			{
				alert("Sp_CMS��ʾ\r\n\n֧����ʽ������д!");
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
	  <LI><A href="pulg.order.pay.list.asp">֧����ʽ����</A></LI>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">����֧����ʽ</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--��1����ǩ//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">֧����ʽ��</TD>
				<TD><input name="txt_paytype" class="input" type="text" value=""><span class="huitext">&nbsp;����</span></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="15%">֧����ע�ֶ�һ��</TD>
				<TD><input name="txt_payfield1" class="input" type="text" value=""></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="15%">֧����ע�ֶζ���</TD>
				<TD><input name="txt_payfield2" class="input" type="text" value=""></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="15%">֧����ע�ֶ�����</TD>
				<TD><input name="txt_payfield3" class="input" type="text" value=""></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="15%">֧����ע�ֶ��ģ�</TD>
				<TD><input name="txt_payfield4" class="input" type="text" value=""></TD>
			  </TR>
			  <TR class="content-td1">
				<TD width="15%">֧����ע�ֶ��壺</TD>
				<TD><input name="txt_payfield5" class="input" type="text" value=""></TD>
			  </TR>
			  <TR class="content-td1">
			    <td>˵����</td>
				<TD colspan="2"><textarea name="txt_memo" cols="80" rows="8"></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="����֧����ʽ" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
