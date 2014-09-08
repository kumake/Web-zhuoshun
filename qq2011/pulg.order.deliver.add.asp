<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="../Sp_inc/class_Page.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Dim action:action = VerificationUrlParam("action","string","")
	if action<>"" and action="save" then 
  		dim deliverytype:deliverytype = F.ReplaceFormText("txt_deliverytype")
		dim Delivermoney:Delivermoney = F.ReplaceFormText("txt_Delivermoney")		
		dim memo:memo = F.ReplaceFormText("txt_memo")		
		Con.execute("insert into Sp_orderDeliver (deliverytype,Delivermoney,[memo]) values ('"&deliverytype&"','"&Delivermoney&"','"&memo&"')")
		Alert "增加配送方式成功","pulg.order.deliver.list.asp"
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
			if(document.getElementById("txt_deliverytype").value=="")
			{
				alert("Sp_CMS提示\r\n\n支付方式必须填写!");
				return false;			
			}
			if(document.getElementById("txt_Delivermoney").value=="")
			{
				alert("Sp_CMS提示\r\n\n配送费用必须填写!");
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
	  <LI><A href="pulg.order.deliver.list.asp">配送方式管理</A></LI>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">增加配送方式</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>
		<!--第1个标签//-->
		<DIV>
			<TABLE cellSpacing="1" width="100%">
			  <TR class="content-td1">
				<TD width="15%">配送方式：</TD>
				<TD><input name="txt_deliverytype" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
			    <td>配送费用：</td>
				<TD colspan="2"><input name="txt_Delivermoney" class="input" type="text" value=""><span class="huitext">&nbsp;必填</span></TD>
			  </TR>
			  <TR class="content-td1">
			    <td>说明：</td>
				<TD colspan="2"><textarea name="txt_memo" cols="80" rows="8"></textarea></TD>
			  </TR>
			</TABLE>
		</DIV>
		<div class="divpadding">
		  <input name="btnsearch" value="增加配送方式" class="button" type="submit">
		</div>
	</DIV>
	</DIV>
</DIV>
</form>
</BODY>
</HTML>
