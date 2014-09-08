<!--#include file="../Sp_inc/conn.asp"-->
<!--#include file="UserAuthority.asp"-->
<%
	Con.Execute("delete * from Sp_Comment")
	'Con.Execute("delete * from Sp_Common")
	Con.Execute("delete * from Sp_EmailFeed")
	Con.Execute("delete * from Sp_Form")
	Con.Execute("delete * from Sp_FormField")
	Con.Execute("delete * from Sp_GuestBook")
	Con.Execute("delete * from Sp_Job")
	Con.Execute("delete * from sp_JsTemplate")
	Con.Execute("delete * from sp_Label")
	Con.Execute("delete * from Sp_log")
	Con.Execute("delete * from Sp_Message")
	Con.Execute("delete * from Sp_ModelPurchaseLog")
	Con.Execute("delete * from Sp_orderDeliver")
	Con.Execute("delete * from Sp_OrderForm")
	Con.Execute("delete * from Sp_OrderItem")
	Con.Execute("delete * from Sp_orderOut")
	Con.Execute("delete * from Sp_orderPay")
	Con.Execute("delete * from Sp_Resume")
	Con.Execute("delete * from Sp_ShopCart")
	Con.Execute("delete * from Sp_Tags")
	Con.Execute("delete * from Sp_TransactionLog")
	Con.Execute("delete * from Sp_UploadFile")
	Con.Execute("delete * from Sp_Vote")
	Con.Execute("delete * from Sp_VoteItem")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>后台首页</TITLE>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="images/style.css" type=text/css rel=stylesheet>
	<META content="MSHTML 6.00.3790.4237" name=GENERATOR>
</HEAD>
<BODY>
<DIV id=wrap>
	<UL id=tags>
	  <LI class=selectTag style="LEFT: 0px; TOP: 0px"><A onfocus="this.blur()" onclick="selectTag('tagContent0',this)" href="javascript:void(0)">清除表记录</A></LI>
	</UL>
	<DIV id=tagContent>
	<DIV class="tagContent selectTag content" id=tagContent1>

		<%
		'Set conn = Server.CreateObject("ADODB.Connection")
		'connstr="provider=microsoft.jet.oledb.4.0;data source=" &server.mappath("../Sp_data/sp_cms.mdb")
		'conn.open connstr
		'Set rstSchema=conn.OpenSchema(20)
		'Do Until rstSchema.EOF
		'Response.Write   "Con.Execute(""delete * from "&rstSchema(2)&""")<br>"
		'rstSchema.MoveNext
		'Loop
		%>	 
	 
	  <DIV>
	  	所有表记录清楚成功......
	  </DIV>
	</DIV>
	</DIV>
</DIV>
</BODY>
</HTML>
